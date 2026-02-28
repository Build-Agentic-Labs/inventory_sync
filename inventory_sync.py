import os
import sys
import json
import time
import threading
import ctypes
import shutil
import subprocess
import winreg
import pandas as pd
from pathlib import Path
from datetime import datetime
from supabase import create_client, Client
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item
import sv_ttk
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black

# App version - auto-injected by GitHub Actions on each release build
APP_VERSION = "1.0.0"

# Auto-updater
try:
    import auto_updater
    HAS_UPDATER = True
except ImportError:
    HAS_UPDATER = False
    print("Warning: auto_updater module not found. Auto-updates disabled.")

# FedEx shipping integration
try:
    import fedex_shipping
    HAS_FEDEX = True
except ImportError:
    HAS_FEDEX = False
    print("Warning: fedex_shipping module not found. Shipping labels disabled.")

try:
    import win32print
    import win32api
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import ghostscript
    HAS_GHOSTSCRIPT = True
except ImportError:
    HAS_GHOSTSCRIPT = False

# Enable DPI awareness for crisp UI on Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass

# Set AppUserModelID so Windows notifications show "Inventory Sync" instead of "Python"
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("InventorySync.App")
except:
    pass

TABLE_NAME = "inventory"
ORDERS_TABLE_NAME = "orders"
SALES_TABLE_NAME = "daily_sales"
SALES_FILE_PATTERN = "Sales by Transaction"

# Config file path - use AppData when running as exe
def get_config_dir():
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        return Path(os.environ['LOCALAPPDATA']) / 'InventorySync'
    else:
        # Running as script
        return Path(__file__).parent

CONFIG_DIR = get_config_dir()
CONFIG_FILE = CONFIG_DIR / "config.json"

# PDF output directory
PDF_OUTPUT_DIR = CONFIG_DIR / "order_pdfs"

# Polling interval in seconds
POLL_INTERVAL = 15

# Global variables
tray_icon = None
config = None
supabase: Client = None
polling_active = False
settings_window = None
orders_window = None
pending_action = None  # Used to communicate between tray thread and main thread
main_root = None  # Hidden root for main thread tkinter operations


def init_supabase(url, key):
    """Initialize Supabase client with credentials from config."""
    global supabase
    if url and key:
        supabase = create_client(url, key)
        return True
    return False

# UI Colors
COLORS = {
    "bg": "#1a1a2e",
    "secondary_bg": "#16213e",
    "accent": "#0f3460",
    "primary": "#e94560",
    "text": "#ffffff",
    "text_secondary": "#a0a0a0",
    "success": "#4ecca3",
    "error": "#e94560",
    "valid": "#4ecca3",
    "border": "#0f3460"
}


def load_config():
    """Load configuration from file."""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return None


def save_config(store_name, watch_folder, file_pattern, supabase_url, supabase_key,
                enable_printer=False, printer_name=None,
                fedex_api_key=None, fedex_secret_key=None, fedex_account_number=None,
                shipper_addresses=None, fedex_use_sandbox=False):
    """Save configuration to file."""
    # Ensure config directory exists
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    config = {
        "store_name": store_name,
        "watch_folder": watch_folder,
        "file_pattern": file_pattern,
        "supabase_url": supabase_url,
        "supabase_key": supabase_key,
        "enable_printer": enable_printer,
        "printer_name": printer_name,
        "fedex_api_key": fedex_api_key,
        "fedex_secret_key": fedex_secret_key,
        "fedex_account_number": fedex_account_number,
        "shipper_addresses": shipper_addresses or {},
        "fedex_use_sandbox": fedex_use_sandbox
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)
    return config


def create_tray_icon():
    """Create a high-resolution icon for the system tray."""
    size = 256  # Higher resolution
    image = Image.new('RGBA', (size, size), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # Draw circular background
    draw.ellipse([10, 10, size-10, size-10], fill=(233, 69, 96))

    # Draw sync arrows
    center = size // 2
    arrow_color = (255, 255, 255)

    # Top arrow (curved)
    draw.arc([50, 50, size-50, size-50], start=200, end=340, fill=arrow_color, width=20)
    # Arrow head top
    draw.polygon([(size-70, 70), (size-40, 50), (size-50, 90)], fill=arrow_color)

    # Bottom arrow (curved)
    draw.arc([50, 50, size-50, size-50], start=20, end=160, fill=arrow_color, width=20)
    # Arrow head bottom
    draw.polygon([(70, size-70), (40, size-50), (90, size-50)], fill=arrow_color)

    return image


def find_all_inventory_files(watch_folder, file_pattern):
    """Find all inventory files in watch folder."""
    inventory_files = []
    try:
        for file_name in os.listdir(watch_folder):
            if file_pattern in file_name and file_name.endswith(".xlsx"):
                file_path = os.path.join(watch_folder, file_name)
                mod_time = os.path.getmtime(file_path)
                inventory_files.append((file_path, mod_time, file_name))
    except Exception as e:
        print(f"Error scanning folder: {e}")
    return inventory_files


def get_latest_inventory_file(watch_folder, file_pattern):
    """Get the most recently modified inventory file."""
    files = find_all_inventory_files(watch_folder, file_pattern)
    if not files:
        return None, []
    files.sort(key=lambda x: x[1], reverse=True)
    latest_file = files[0][0]
    all_files = [f[0] for f in files]
    return latest_file, all_files


def delete_all_inventory_files(file_list):
    """Delete all inventory files in the list."""
    for file_path in file_list:
        try:
            os.remove(file_path)
            print(f"Deleted: {os.path.basename(file_path)}")
        except Exception as e:
            print(f"Warning: Could not delete {os.path.basename(file_path)}: {e}")


def find_sales_files(watch_folder):
    """Find all sales transaction files in watch folder."""
    sales_files = []
    try:
        for file_name in os.listdir(watch_folder):
            if SALES_FILE_PATTERN in file_name and file_name.endswith(".xlsx"):
                file_path = os.path.join(watch_folder, file_name)
                mod_time = os.path.getmtime(file_path)
                sales_files.append((file_path, mod_time, file_name))
    except Exception as e:
        print(f"Error scanning folder for sales files: {e}")
    return sales_files


def get_latest_sales_file(watch_folder):
    """Get the most recently modified sales file."""
    files = find_sales_files(watch_folder)
    if not files:
        return None, []
    files.sort(key=lambda x: x[1], reverse=True)
    latest_file = files[0][0]
    all_files = [f[0] for f in files]
    return latest_file, all_files


def process_sales_file(file_path, store_name):
    """Read sales file and extract daily totals."""
    try:
        df = pd.read_excel(file_path)
        print(f"Read {len(df)} rows from sales file")

        # Find the Total row or calculate totals
        total_row = df[df['Trans ID'] == 'Total']

        # Get the date from transaction rows (exclude Total row)
        transaction_rows = df[df['Trans ID'] != 'Total']

        if transaction_rows.empty:
            print("No transaction rows found in sales file")
            return None

        # Extract date from first transaction
        date_str = transaction_rows.iloc[0]['Date']
        try:
            # Parse MM/DD/YYYY format
            report_date = datetime.strptime(str(date_str), "%m/%d/%Y").date()
        except (ValueError, TypeError):
            # Try alternative format or use today
            try:
                report_date = pd.to_datetime(date_str).date()
            except:
                report_date = datetime.now().date()
                print(f"Warning: Could not parse date '{date_str}', using today's date")

        # Calculate totals
        if not total_row.empty:
            # Use the Total row if available
            row = total_row.iloc[0]
            total_transactions = len(transaction_rows)
            total_qty_sold = int(row.get('Qty Sold', 0)) if pd.notna(row.get('Qty Sold')) else 0
            total_sales = float(row.get('Sales', 0)) if pd.notna(row.get('Sales')) else 0
            total_cogs = float(row.get('COGS', 0)) if pd.notna(row.get('COGS')) else 0
            total_gross_profit = float(row.get('Gross Profit', 0)) if pd.notna(row.get('Gross Profit')) else 0
            total_discounts = float(row.get('Disc&Mkd', 0)) if pd.notna(row.get('Disc&Mkd')) else 0
            total_tax = float(row.get('Tax', 0)) if pd.notna(row.get('Tax')) else 0
            total_receipts = float(row.get('Receipt Total', 0)) if pd.notna(row.get('Receipt Total')) else 0
        else:
            # Calculate totals manually from transaction rows
            total_transactions = len(transaction_rows)
            total_qty_sold = int(transaction_rows['Qty Sold'].sum())
            total_sales = float(transaction_rows['Sales'].sum())
            total_cogs = float(transaction_rows['COGS'].sum())
            total_gross_profit = float(transaction_rows['Gross Profit'].sum())
            total_discounts = float(transaction_rows['Disc&Mkd'].sum())
            total_tax = float(transaction_rows['Tax'].sum())
            total_receipts = float(transaction_rows['Receipt Total'].sum())

        # Calculate average gross margin
        if total_sales > 0:
            avg_gross_margin = (total_gross_profit / total_sales) * 100
        else:
            avg_gross_margin = 0

        record = {
            "store_name": store_name,
            "report_date": report_date.isoformat(),
            "total_transactions": total_transactions,
            "total_qty_sold": total_qty_sold,
            "total_sales": round(total_sales, 2),
            "total_cogs": round(total_cogs, 2),
            "total_gross_profit": round(total_gross_profit, 2),
            "avg_gross_margin": round(avg_gross_margin, 2),
            "total_discounts": round(total_discounts, 2),
            "total_tax": round(total_tax, 2),
            "total_receipts": round(total_receipts, 2)
        }

        print(f"Processed sales for {report_date}: {total_transactions} transactions, ${total_receipts:.2f} total")
        return record

    except Exception as e:
        print(f"Error processing sales file: {e}")
        import traceback
        traceback.print_exc()
        return None


def clean_data(df: pd.DataFrame) -> list[dict]:
    """Clean and prepare data for Supabase insert."""
    records = []
    for _, row in df.iterrows():
        gross_margin = row.get("Gross Margin", "0%")
        if isinstance(gross_margin, str):
            gross_margin = float(gross_margin.replace("%", "")) / 100
        else:
            gross_margin = float(gross_margin) if pd.notna(gross_margin) else 0

        record = {
            "product_name": str(row.get("Product Name", "")) if pd.notna(row.get("Product Name")) else None,
            "sku": str(row.get("SKU")).strip() if pd.notna(row.get("SKU")) else None,
            "vendor": str(row.get("Vendor", "")) if pd.notna(row.get("Vendor")) else None,
            "brand": str(row.get("Brand", "")) if pd.notna(row.get("Brand")) else None,
            "price": float(row.get("Price", 0)) if pd.notna(row.get("Price")) else 0,
            "cost": float(row.get("Cost", 0)) if pd.notna(row.get("Cost")) else 0,
            "total_stock": int(row.get("Total Stock", 0)) if pd.notna(row.get("Total Stock")) else 0,
            "committed": int(row.get("Committed", 0)) if pd.notna(row.get("Committed")) else 0,
            "open_stock": int(row.get("Open Stock", 0)) if pd.notna(row.get("Open Stock")) else 0,
            "qty_on_order": int(row.get("Qty On Order", 0)) if pd.notna(row.get("Qty On Order")) else 0,
            "gross_margin": gross_margin,
            "total_retail": float(row.get("Total Retail", 0)) if pd.notna(row.get("Total Retail")) else 0,
            "total_cost": float(row.get("Total Cost", 0)) if pd.notna(row.get("Total Cost")) else 0,
        }
        if record["sku"]:
            records.append(record)
    return records


def validate_inventory_file(df):
    """
    Validate inventory file by checking for marker products.

    Both markers should always be present:
    - toppenish (SKU 99999) - should have qty = 1 for correct file
    - yakima (SKU 9999) - should have qty = 0 for correct file

    Returns:
        (is_valid, error_type, message)
        - is_valid: True if file should be synced
        - error_type: None, "wrong_file", or "no_marker"
        - message: Description of validation result
    """
    toppenish_qty = None
    yakima_qty = None

    # Look for both validation markers in the dataframe
    for _, row in df.iterrows():
        product_name = str(row.get("Product Name", "")).strip().lower()
        sku = str(row.get("SKU", "")).strip()
        qty = int(row.get("Total Stock", 0)) if pd.notna(row.get("Total Stock")) else 0

        # Check for toppenish marker (SKU 99999 - 5 nines)
        if "toppenish" in product_name and sku == "99999":
            toppenish_qty = qty

        # Check for yakima marker (SKU 9999 - 4 nines)
        if "yakima" in product_name and sku == "9999":
            yakima_qty = qty

    # Check if both markers were found
    if toppenish_qty is None or yakima_qty is None:
        return False, "no_marker", "No inventory ID markers found"

    # Valid: toppenish = 1, yakima = 0
    if toppenish_qty == 1 and yakima_qty == 0:
        return True, None, "Correct inventory file (Toppenish=1, Yakima=0)"

    # Invalid: yakima has qty of 1 (wrong file)
    if yakima_qty == 1:
        return False, "wrong_file", f"Wrong inventory file detected (Yakima={yakima_qty}, Toppenish={toppenish_qty})"

    # Any other combination is invalid
    return False, "wrong_file", f"Invalid marker quantities (Toppenish={toppenish_qty}, Yakima={yakima_qty})"


def sync_inventory(watch_folder, file_pattern, show_notification=True):
    """Find latest inventory file, sync to Supabase, delete all inventory files."""
    global tray_icon

    latest_file, all_files = get_latest_inventory_file(watch_folder, file_pattern)

    if not latest_file:
        return False

    print(f"\n{'='*50}")
    print(f"Found {len(all_files)} inventory file(s)")
    print(f"Using latest: {os.path.basename(latest_file)}")
    print(f"{'='*50}")

    try:
        time.sleep(2)

        df = pd.read_excel(latest_file)
        print(f"Read {len(df)} rows from Excel")

        # Validate the inventory file
        is_valid, error_type, validation_msg = validate_inventory_file(df)
        print(f"Validation: {validation_msg}")

        if error_type == "wrong_file":
            # Wrong file detected - notify user and delete without syncing
            print(f"[ERROR] {validation_msg} - NOT syncing!")
            if tray_icon:
                tray_icon.notify(
                    "Wrong inventory file! This appears to be the Yakima inventory. File deleted.",
                    "Inventory Sync Error"
                )
            # Delete the wrong file(s)
            print(f"\nDeleting wrong file(s)...")
            delete_all_inventory_files(all_files)
            print("Wrong file(s) deleted!")
            return False

        if error_type == "no_marker":
            # No marker found - warn user and delete without syncing
            print(f"[WARNING] {validation_msg} - NOT syncing!")
            if tray_icon:
                tray_icon.notify(
                    "Are you sure this is Toppenish inventory? Could not detect correct inventory ID. File deleted.",
                    "Inventory Sync Warning"
                )
            # Delete the unverified file(s)
            print(f"\nDeleting unverified file(s)...")
            delete_all_inventory_files(all_files)
            print("Unverified file(s) deleted!")
            return False

        # Filter out both validation marker products (toppenish SKU 99999 and yakima SKU 9999)
        df_filtered = df[~(
            ((df['Product Name'].str.lower().str.contains('toppenish', na=False)) &
             (df['SKU'].astype(str).str.strip() == '99999')) |
            ((df['Product Name'].str.lower().str.contains('yakima', na=False)) &
             (df['SKU'].astype(str).str.strip() == '9999'))
        )]
        print(f"Filtered out validation markers, {len(df_filtered)} products remaining")

        records = clean_data(df_filtered)
        print(f"Prepared {len(records)} valid records")

        if not records:
            print("No valid records to sync")
            return False

        response = supabase.table(TABLE_NAME).upsert(
            records,
            on_conflict="sku"
        ).execute()

        print(f"Successfully synced {len(records)} records to Supabase!")

        print(f"\nCleaning up {len(all_files)} file(s)...")
        delete_all_inventory_files(all_files)
        print("Cleanup complete!")

        if show_notification and tray_icon:
            store_name = config.get("store_name", "Store") if config else "Store"
            tray_icon.notify(
                f"{store_name} inventory synced to site",
                "Inventory Sync"
            )
        return True

    except Exception as e:
        print(f"Error syncing inventory: {e}")
        if tray_icon:
            tray_icon.notify(f"Sync error: {str(e)[:50]}", "Inventory Sync Error")
        return False


def sync_sales(watch_folder, store_name, show_notification=True):
    """Find sales files, extract totals, sync to Supabase, delete files."""
    global tray_icon

    latest_file, all_files = get_latest_sales_file(watch_folder)

    if not latest_file:
        return False

    print(f"\n{'='*50}")
    print(f"SALES SYNC: Processing {os.path.basename(latest_file)}")
    print(f"{'='*50}")

    try:
        # Process the sales file
        record = process_sales_file(latest_file, store_name)

        if not record:
            print("Failed to process sales file")
            return False

        # Sync to Supabase
        print(f"\nSyncing daily sales to Supabase...")
        response = supabase.table(SALES_TABLE_NAME).upsert(
            record,
            on_conflict="store_name,report_date"
        ).execute()

        print(f"Successfully synced sales data for {record['report_date']}!")

        # Delete all sales files
        print(f"\nCleaning up {len(all_files)} sales file(s)...")
        for file_path in all_files:
            try:
                os.remove(file_path)
                print(f"Deleted: {os.path.basename(file_path)}")
            except Exception as e:
                print(f"Warning: Could not delete {os.path.basename(file_path)}: {e}")
        print("Cleanup complete!")

        if show_notification and tray_icon:
            tray_icon.notify(
                f"Sales synced: {record['total_transactions']} transactions, ${record['total_receipts']:.2f}",
                "Sales Sync"
            )
        return True

    except Exception as e:
        print(f"Error syncing sales: {e}")
        import traceback
        traceback.print_exc()
        if tray_icon:
            tray_icon.notify(f"Sales sync error: {str(e)[:50]}", "Sales Sync Error")
        return False


def polling_loop():
    """Main polling loop - checks for files and orders every POLL_INTERVAL seconds."""
    global config, polling_active

    polling_active = True
    print(f"Polling every {POLL_INTERVAL} seconds...")

    while polling_active:
        try:
            if config:
                # Check for inventory files
                files = find_all_inventory_files(config["watch_folder"], config["file_pattern"])
                if files:
                    print(f"\n[POLL] Found {len(files)} inventory file(s)")
                    sync_inventory(config["watch_folder"], config["file_pattern"])

                # Check for sales files (combined report for all stores)
                sales_files = find_sales_files(config["watch_folder"])
                if sales_files:
                    print(f"\n[POLL] Found {len(sales_files)} sales file(s)")
                    sync_sales(config["watch_folder"], "All Stores")

                # Check for new orders
                new_orders = poll_for_orders_once()
                if new_orders > 0:
                    print(f"\n[POLL] Processed {new_orders} new order(s)")

        except Exception as e:
            print(f"Polling error: {e}")

        time.sleep(POLL_INTERVAL)


def stop_polling():
    """Stop the polling loop."""
    global polling_active
    polling_active = False


def get_available_printers():
    """Get list of actual printer names from Windows"""
    try:
        ps_script = '''
$printers = Get-Printer -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
$printers | ForEach-Object { Write-Output $_ }
'''
        result = subprocess.run(
            ['powershell', '-NoProfile', '-Command', ps_script],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.returncode == 0:
            printers = [p.strip() for p in result.stdout.strip().split('\n') if p.strip()]
            return printers
        return []
    except Exception as e:
        print(f"Error getting available printers: {e}")
        return []

def get_actual_printer_name(printer_name):
    """Get the actual Windows printer name that matches the stored printer name"""
    available_printers = get_available_printers()

    if not available_printers:
        print(f"Warning: Could not retrieve printer list, using provided name: {printer_name}")
        return printer_name

    # Exact match
    if printer_name in available_printers:
        print(f"Found exact match: {printer_name}")
        return printer_name

    # Case-insensitive match
    for p in available_printers:
        if p.lower() == printer_name.lower():
            print(f"Found case-insensitive match: {p}")
            return p

    # Partial match (contains)
    for p in available_printers:
        if printer_name.lower() in p.lower():
            print(f"Found partial match: {p}")
            return p

    # No match found
    print(f"Warning: Printer '{printer_name}' not found in Windows. Available printers: {available_printers}")
    return printer_name

def get_default_printer():
    """Get the default printer name"""
    try:
        if HAS_WIN32:
            return win32print.GetDefaultPrinter()
        else:
            # Fallback: try to get from Windows registry
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows NT\CurrentVersion\Devices')
            default_printer = winreg.QueryValueEx(key, 'Device')[0].split(',')[0]
            winreg.CloseKey(key)
            return default_printer
    except Exception as e:
        print(f"Error getting default printer: {e}")
        return None


def get_printer_ip_address(printer_name):
    """Get the network IP address of a network printer from Windows registry
    Works with: WSD printers, TCP/IP printers, and various driver types"""
    import re

    try:
        # Try to get from printer registry
        key_path = r'System\CurrentControlSet\Control\Print\Printers'
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)

        try:
            subkey = winreg.OpenKey(key, printer_name)

            # Method 1: Check Location field (WSD printers - Web Services for Devices)
            # Used by: Canon, HP, Brother, Xerox, modern network printers
            try:
                location = winreg.QueryValueEx(subkey, 'Location')[0]
                if location:
                    # Extract IP from URL like: http://192.168.86.140:80/wsd/...
                    match = re.search(r'(\d+\.\d+\.\d+\.\d+)', location)
                    if match:
                        ip = match.group(1)
                        print(f"[Auto-detected] WSD printer IP: {ip}")
                        return ip
            except:
                pass

            # Method 2: Check PortName for direct IP (Standard TCP/IP printers)
            # Used by: HP LaserJet, Xerox, Ricoh, enterprise printers
            try:
                port_name = winreg.QueryValueEx(subkey, 'PortName')[0]
                if port_name:
                    # Handle various port name formats
                    # Format 1: "192.168.1.100" (direct IP)
                    # Format 2: "IP_192.168.1.100" (prefixed)
                    # Format 3: "192.168.1.100:9100" (with port)
                    match = re.search(r'(\d+\.\d+\.\d+\.\d+)', port_name)
                    if match:
                        ip = match.group(1)
                        print(f"[Auto-detected] TCP/IP port: {ip}")
                        return ip
            except:
                pass

            # Method 3: Check PrinterDriverData for network configuration
            # Used by: Various specialized printers, driver-specific configs
            try:
                try:
                    driver_data_key = winreg.OpenKey(subkey, r'PrinterDriverData')
                    # Look for common IP address keys across different drivers
                    for ip_key in ['IPAddress', 'IP', 'HostAddress', 'NetworkAddress',
                                  'HostIPAddress', 'PrinterIP', 'ServerAddress',
                                  'DeviceIPAddress', 'NetworkIP']:
                        try:
                            ip_value = winreg.QueryValueEx(driver_data_key, ip_key)[0]
                            if ip_value and '.' in str(ip_value):
                                match = re.search(r'(\d+\.\d+\.\d+\.\d+)', str(ip_value))
                                if match:
                                    ip = match.group(1)
                                    print(f"[Auto-detected] Driver data IP: {ip}")
                                    return ip
                        except:
                            pass
                    winreg.CloseKey(driver_data_key)
                except:
                    pass
            except:
                pass

            # Method 4: Check for port monitor configuration
            # Some printers store config in Monitors section
            try:
                try:
                    port_monitors_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                        r'System\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports')
                    # Enumerate all TCP/IP ports
                    i = 0
                    while True:
                        try:
                            port_name_enum = winreg.EnumKey(port_monitors_key, i)
                            # Check if this port belongs to our printer
                            if '.' in port_name_enum:  # Likely an IP
                                match = re.search(r'(\d+\.\d+\.\d+\.\d+)', port_name_enum)
                                if match:
                                    return match.group(1)
                            i += 1
                        except OSError:
                            break
                    winreg.CloseKey(port_monitors_key)
                except:
                    pass
            except:
                pass

            winreg.CloseKey(subkey)
        except WindowsError:
            print(f"Note: Printer '{printer_name}' may be local (USB) or not found in registry")
        except Exception as e:
            print(f"Note: Could not read printer registry: {e}")

        winreg.CloseKey(key)
    except Exception as e:
        print(f"Note: Printer registry access limited: {e}")

    # Note: Not finding an IP doesn't necessarily mean printing will fail
    # Local/USB printers don't have IPs but still work through Windows queue
    print(f"Note: Could not auto-detect IP for '{printer_name}'")
    print(f"      (This is normal for USB/local printers)")
    print(f"      Printing will still work through Windows print queue")
    return None


def detect_ghostscript_path():
    """Detect Ghostscript executable path"""
    try:
        # First check for bundled Ghostscript (when running as exe)
        if getattr(sys, 'frozen', False):
            # Running as compiled exe - check for bundled GS
            bundle_dir = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(sys.executable).parent
            bundled_gs = bundle_dir / "gs" / "bin" / "gswin64c.exe"
            if bundled_gs.exists():
                print(f"Found bundled Ghostscript: {bundled_gs}")
                return str(bundled_gs)
            # Also check in app install directory
            install_dir = Path(sys.executable).parent
            installed_gs = install_dir / "gs" / "bin" / "gswin64c.exe"
            if installed_gs.exists():
                print(f"Found installed Ghostscript: {installed_gs}")
                return str(installed_gs)

        # Check common Ghostscript installation paths (newest versions first)
        gs_paths = [
            r"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe",
            r"C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe",
            r"C:\Program Files\gs\gs10.0.0\bin\gswin64c.exe",
            r"C:\Program Files (x86)\gs\gs10.06.0\bin\gswin32c.exe",
            r"C:\Program Files (x86)\gs\gs10.01.2\bin\gswin32c.exe",
            r"C:\Program Files (x86)\gs\gs10.0.0\bin\gswin32c.exe",
            r"C:\Program Files\ghostscript\gs10.06.0\bin\gswin64c.exe",
        ]

        # Check if any path exists
        for path in gs_paths:
            if os.path.exists(path):
                print(f"Found Ghostscript: {path}")
                return path

        # Try to find via system PATH
        result = subprocess.run(['where', 'gswin64c.exe'], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode == 0:
            gs_path = result.stdout.strip().split('\n')[0]
            print(f"Found Ghostscript in PATH: {gs_path}")
            return gs_path

        result = subprocess.run(['where', 'gswin32c.exe'], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode == 0:
            gs_path = result.stdout.strip().split('\n')[0]
            print(f"Found Ghostscript in PATH: {gs_path}")
            return gs_path

    except Exception as e:
        print(f"Error detecting Ghostscript: {e}")

    return None


def send_via_ghostscript(pdf_path, printer_name):
    """Send PDF directly to printer using Ghostscript (simple, direct)"""
    try:
        gs_path = detect_ghostscript_path()
        if not gs_path:
            print("Ghostscript not found")
            return False

        pdf_path = str(pdf_path)
        print(f"Sending PDF to {printer_name}...")

        # Verify PDF exists before printing
        if not os.path.exists(pdf_path):
            print(f"[ERROR] PDF file does not exist: {pdf_path}")
            return False

        print(f"Submitting to Windows print queue...")

        # Use Ghostscript directly with mswinpr2 device
        cmd = [gs_path]
        cmd.extend([
            "-sDEVICE=mswinpr2",
            f'-sOutputFile=%printer%{printer_name}',
            "-dBATCH",
            "-dNOPAUSE",
            "-dQUIET",
            pdf_path
        ])

        print(f"[DEBUG] Ghostscript command: {' '.join(cmd[:4])}...")

        result = subprocess.run(cmd, capture_output=True, timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)

        if result.returncode == 0:
            print(f"[OK] PDF sent to printer")
            return True
        else:
            stderr_msg = result.stderr.decode().strip()
            print(f"[ERROR] Ghostscript failed (code {result.returncode}): {stderr_msg}")
            return False

    except Exception as e:
        print(f"Error in send_via_ghostscript: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_pdf_to_printer(pdf_path, printer_name=None):
    """Send PDF to physical printer using Ghostscript through Windows print queue"""
    try:
        pdf_path = str(pdf_path)

        # Validate PDF exists
        if not os.path.exists(pdf_path):
            print(f"Error: PDF file not found: {pdf_path}")
            return False

        if printer_name is None:
            printer_name = get_default_printer()

        if printer_name is None:
            print("Error: No printer found or specified")
            return False

        # Get the actual Windows printer name (exact match, case-insensitive)
        actual_printer_name = get_actual_printer_name(printer_name)

        print(f"Sending to printer: {actual_printer_name}")

        # Send via Ghostscript to Windows print queue
        if send_via_ghostscript(pdf_path, actual_printer_name):
            return True
        else:
            print(f"Failed to send PDF via Ghostscript")
            return False

    except Exception as e:
        print(f"Error in send_pdf_to_printer: {e}")
        import traceback
        traceback.print_exc()
        return False


def check_printer_setup():
    """Diagnostic function: Check if printer and Ghostscript are ready"""
    print("\n" + "=" * 70)
    print("PRINTER SETUP CHECK")
    print("=" * 70)

    # Check Ghostscript
    print("\nChecking Ghostscript...")
    gs_path = detect_ghostscript_path()
    if gs_path:
        print(f"  [OK] Ghostscript ready")
    else:
        print(f"  [ERROR] Ghostscript not installed")
        return False

    # Check printers
    print("\nScanning installed printers...")
    printers = get_available_printers()
    if not printers:
        print(f"  [ERROR] No printers found")
        return False
    print(f"  [OK] Found {len(printers)} printer(s)")

    # Check for Canon printer
    print("\nDetecting Canon TR4700...")
    actual_name = get_actual_printer_name("Canon TR4700 series")
    if actual_name not in printers:
        print(f"  [WARN] Canon TR4700 not found in Windows")
        print(f"  Available printers: {', '.join(printers[:5])}...")
        return False

    # Get IP address
    print(f"  [OK] Found: {actual_name}")
    ip_address = get_printer_ip_address(actual_name)
    if not ip_address:
        print(f"  [WARN] Could not auto-detect IP address")
        return False

    print(f"  [OK] IP auto-detected: {ip_address}")

    print("\n" + "=" * 70)
    print("STATUS: Printer setup is ready for automatic printing!")
    print("=" * 70 + "\n")
    return True


def format_fulfillment(item):
    """Format fulfillment information for printing"""
    fulfillment = item.get('fulfillment', {})
    method = fulfillment.get('method', 'N/A')

    if method == 'pickup':
        location = fulfillment.get('location', 'Unknown')
        location_name = "Yakima" if location == 1 else "Toppenish"
        return f"PICKUP at {location_name}"
    elif method == 'delivery':
        address = fulfillment.get('address', {})
        return f"DELIVERY to {address.get('city', 'Unknown')}"
    elif method == 'shipping':
        address = fulfillment.get('address', {})
        return f"SHIPPING to {address.get('city', 'Unknown')}, {address.get('state', 'Unknown')}"
    else:
        return method.upper()


def create_pdf_order(order, pdf_path):
    """Generate a professional PDF order form for workers"""
    try:
        c = canvas.Canvas(str(pdf_path), pagesize=letter)
        width, height = letter

        # Colors
        gray_bar = HexColor("#999999")
        text_color = black
        light_gray = HexColor("#eeeeee")

        # Add logo at top left
        logo_path = Path(__file__).parent / "CASCADELOGO.png"
        if logo_path.exists():
            c.drawImage(str(logo_path), 0.75 * inch, height - 1.2 * inch, width=1.5 * inch, height=0.6 * inch, preserveAspectRatio=True)

        y_position = height - 1.4 * inch

        # Header - "INVOICE" on the right
        c.setFont("Helvetica-Bold", 42)
        c.setFillColor(HexColor("#1a1a1a"))
        c.drawRightString(width - 0.75 * inch, y_position, "INVOICE")

        y_position -= 0.35 * inch

        # Header line
        c.setStrokeColor(HexColor("#2C3E50"))
        c.setLineWidth(2)
        c.line(0.75 * inch, y_position, width - 0.75 * inch, y_position)

        y_position -= 0.25 * inch
        c.setFont("Helvetica", 13)
        c.setFillColor(HexColor("#555555"))
        c.drawRightString(width - 0.75 * inch, y_position, f"Order #{order['order_number']}")

        # Customer Information - below header line with better typography
        y_position -= 0.15 * inch

        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(HexColor("#666666"))
        c.drawString(0.75 * inch, y_position, "NAME")
        c.setFont("Helvetica", 11)
        c.setFillColor(text_color)
        c.drawString(1.8 * inch, y_position, f"{order['customer_first_name']} {order['customer_last_name']}")

        y_position -= 0.25 * inch
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(HexColor("#666666"))
        c.drawString(0.75 * inch, y_position, "EMAIL")
        c.setFont("Helvetica", 11)
        c.setFillColor(text_color)
        c.drawString(1.8 * inch, y_position, order['customer_email'])

        y_position -= 0.25 * inch
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(HexColor("#666666"))
        c.drawString(0.75 * inch, y_position, "NUMBER")
        c.setFont("Helvetica", 11)
        c.setFillColor(text_color)
        c.drawString(1.8 * inch, y_position, order.get('customer_phone', 'N/A'))

        # Shipping Address (if provided)
        if order.get('customer_shipping_address'):
            addr = order['customer_shipping_address']
            y_position -= 0.5 * inch
            c.setFont("Helvetica-Bold", 10)
            c.setFillColor(HexColor("#666666"))
            c.drawString(0.75 * inch, y_position, "SHIPPING ADDRESS")

            y_position -= 0.25 * inch
            c.setFont("Helvetica", 10)
            c.setFillColor(text_color)
            c.drawString(0.75 * inch, y_position, addr.get('address1', '').upper())

            if addr.get('address2'):
                y_position -= 0.2 * inch
                c.drawString(0.75 * inch, y_position, addr['address2'].upper())

            y_position -= 0.2 * inch
            c.drawString(0.75 * inch, y_position, f"{addr.get('city', '').upper()} {addr.get('state', '').upper()} {addr.get('zipCode', '')}")

        y_position -= 0.6 * inch

        # Group items by fulfillment method and location
        items_by_fulfillment = {}
        for item in order['items']:
            fulfillment = item.get('fulfillment', {})
            method = fulfillment.get('method', 'unknown')

            # For pickup, also group by location
            if method == 'pickup':
                # Check both 'location' (int) and 'pickupLocation' (string) fields
                location = fulfillment.get('location') or fulfillment.get('pickupLocation')

                # Handle both int and string location values
                if location == 1 or location == '1' or str(location).lower() == 'yakima':
                    location_name = "Yakima"
                elif location == 2 or location == '2' or str(location).lower() == 'toppenish':
                    location_name = "Toppenish"
                else:
                    location_name = "Toppenish"  # Default to Toppenish

                key = f"pickup_{location_name}"
            else:
                key = method

            if key not in items_by_fulfillment:
                items_by_fulfillment[key] = []
            items_by_fulfillment[key].append(item)

        # Define order and colors for fulfillment types
        fulfillment_order = ['shipping', 'delivery', 'pickup_Yakima', 'pickup_Toppenish']
        fulfillment_colors = {
            'shipping': HexColor("#8B0000"),      # Dark red
            'delivery': HexColor("#2C3E50"),      # Deep slate
            'pickup_Yakima': HexColor("#00008B"),    # Dark blue
            'pickup_Toppenish': HexColor("#00008B")  # Dark blue
        }

        # Add any unknown fulfillment types at the end
        for fulfillment_type in items_by_fulfillment.keys():
            if fulfillment_type not in fulfillment_order:
                fulfillment_order.append(fulfillment_type)

        # Draw items grouped by fulfillment in the specified order
        for fulfillment_key in fulfillment_order:
            if fulfillment_key not in items_by_fulfillment:
                continue

            items = items_by_fulfillment[fulfillment_key]
            # Fulfillment type bar with color
            bar_color = fulfillment_colors.get(fulfillment_key, gray_bar)
            c.setFillColor(bar_color)
            c.roundRect(0.75 * inch, y_position - 0.3 * inch, width - 1.5 * inch, 0.35 * inch, 0.05 * inch, fill=True, stroke=False)

            c.setFillColor(HexColor("#ffffff"))
            c.setFont("Helvetica-Bold", 11)

            # Get fulfillment label
            if fulfillment_key == 'shipping':
                label = "SHIPPING"
                location_label = None
            elif fulfillment_key.startswith('pickup_'):
                label = "LOCAL PICKUP"
                location_name = fulfillment_key.replace('pickup_', '')
                location_label = location_name.upper()
            elif fulfillment_key == 'delivery':
                label = "LOCAL DELIVERY"
                location_label = None
            else:
                label = f"{fulfillment_key.upper()}"
                location_label = None

            # Left-aligned label, with location on right if applicable
            c.drawString(0.85 * inch, y_position - 0.15 * inch, label)
            if location_label:
                c.drawRightString(width - 0.85 * inch, y_position - 0.15 * inch, location_label)

            y_position -= 0.5 * inch

            # Table headers
            c.setFillColor(HexColor("#666666"))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(0.75 * inch, y_position, "ITEM")
            c.drawString(2.3 * inch, y_position, "QTY")
            c.drawString(2.8 * inch, y_position, "SKU")
            c.drawRightString(width - 3.5 * inch, y_position, "PRICE")
            c.drawRightString(width - 2 * inch, y_position, "SHIPPING")
            c.drawRightString(width - 0.75 * inch, y_position, "AMOUNT")

            y_position -= 0.08 * inch
            c.setStrokeColor(HexColor("#DDDDDD"))
            c.setLineWidth(0.5)
            c.line(0.75 * inch, y_position, width - 0.75 * inch, y_position)

            y_position -= 0.3 * inch

            # Items in this fulfillment group
            c.setFont("Helvetica", 9)
            c.setFillColor(text_color)
            for item in items:
                item_price = float(item['price'])
                item_qty = int(item['quantity'])
                shipping_cost = float(item.get('shippingCost', 0))
                amount = item_price * item_qty + shipping_cost

                # Item name
                item_name = item['name']

                # SKU
                sku = item.get('sku', 'N/A')

                c.drawString(0.75 * inch, y_position, item_name)
                c.drawString(2.3 * inch, y_position, str(item_qty))
                c.drawString(2.8 * inch, y_position, sku)
                c.drawRightString(width - 3.5 * inch, y_position, f"{item_price:.2f}")

                if shipping_cost > 0:
                    c.drawRightString(width - 2 * inch, y_position, f"{shipping_cost:.2f}")
                else:
                    c.drawRightString(width - 2 * inch, y_position, "FREE")

                c.drawRightString(width - 0.75 * inch, y_position, f"{amount:.2f}")

                y_position -= 0.25 * inch

            # Line after items
            c.setLineWidth(0.5)
            c.line(0.75 * inch, y_position, width - 0.75 * inch, y_position)
            y_position -= 0.4 * inch

            # Check if we need a new page
            if y_position < 3 * inch:
                c.showPage()
                y_position = height - inch

        # Totals Section
        y_position -= 0.25 * inch
        c.setFillColor(HexColor("#555555"))
        c.setFont("Helvetica", 10)
        c.drawString(0.75 * inch, y_position, "Sub Total")
        c.drawRightString(width - 0.75 * inch, y_position, f"${float(order['subtotal']):.2f}")

        # Tax if applicable
        tax_amount = float(order.get('tax_amount', 0))
        if tax_amount > 0:
            y_position -= 0.22 * inch
            c.drawString(0.75 * inch, y_position, "Tax (8.5%)")
            c.drawRightString(width - 0.75 * inch, y_position, f"${tax_amount:.2f}")

        # Discount if applicable
        discount = float(order.get('discount', 0))
        if discount > 0:
            y_position -= 0.22 * inch
            c.drawString(0.75 * inch, y_position, "Discount")
            c.drawRightString(width - 0.75 * inch, y_position, f"-${discount:.2f}")

        y_position -= 0.1 * inch
        c.setStrokeColor(HexColor("#2C3E50"))
        c.setLineWidth(2)
        c.line(0.75 * inch, y_position, width - 0.75 * inch, y_position)

        y_position -= 0.25 * inch
        c.setFont("Helvetica-Bold", 16)
        c.setFillColor(HexColor("#1a1a1a"))
        c.drawString(0.75 * inch, y_position, "TOTAL")
        c.drawRightString(width - 0.75 * inch, y_position, f"${float(order['total']):.2f}")

        c.save()
        return True
    except Exception as e:
        print(f"Error generating PDF: {e}")
        import traceback
        traceback.print_exc()
        return False


def print_order_pdf(order, send_to_printer=True, printer_name=None):
    """Generate PDF for an order and optionally send to printer"""
    pdf_filename = f"order_{order['order_number']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    pdf_path = PDF_OUTPUT_DIR / pdf_filename

    # First, create the PDF
    pdf_created = False
    try:
        pdf_created = create_pdf_order(order, pdf_path)
    except Exception as e:
        print(f"Error generating PDF file: {e}")
        return None

    if not pdf_created:
        print(f"Failed to create PDF for order {order['order_number']}")
        return None

    # PDF created successfully, send to printer if enabled
    printer_success = True  # Assume success if not auto-printing
    if send_to_printer:
        printer_result = send_pdf_to_printer(pdf_path, printer_name=printer_name)
        if not printer_result:
            print(f"[ERROR] Failed to send PDF to printer: {pdf_path}")
            print(f"[ERROR] Order will NOT be marked as printed until manual action taken")
            return None  # Return None to indicate failure - do NOT mark as printed
        printer_success = True
        print(f"[OK] PDF successfully sent to printer")

    # Only mark as printed if printing was successful (or if auto-print was disabled)
    if printer_success:
        try:
            # First try with pdf_path
            update_data = {
                'printed': True,
                'printed_at': datetime.now().isoformat(),
                'pdf_path': str(pdf_path)
            }
            supabase.table(ORDERS_TABLE_NAME).update(update_data).eq('id', order['id']).execute()
            return pdf_path
        except Exception as db_error:
            # If pdf_path column doesn't exist, try without it
            error_msg = str(db_error)
            if 'pdf_path' in error_msg or 'PGRST204' in error_msg:
                try:
                    update_data_simple = {
                        'printed': True,
                        'printed_at': datetime.now().isoformat()
                    }
                    supabase.table(ORDERS_TABLE_NAME).update(update_data_simple).eq('id', order['id']).execute()
                    return pdf_path
                except Exception as retry_error:
                    print(f"Error marking order as printed: {retry_error}")
                    return None
            else:
                print(f"Database error: {db_error}")
                return None
    else:
        # Printing failed, don't mark as printed
        return None


def poll_for_orders_once():
    """Poll Supabase for unprinted orders and print them (filtered by location)"""
    global supabase, config
    try:
        # Get store name for location filtering
        store_name = config.get('store_name', '').lower() if config else ''

        # Build query with location filtering
        # Show orders where order_location matches this store OR equals 'both'
        if store_name in ['yakima', 'toppenish']:
            response = supabase.table(ORDERS_TABLE_NAME)\
                .select('*')\
                .eq('printed', False)\
                .or_(f'order_location.eq.{store_name},order_location.eq.both')\
                .order('created_at', desc=True)\
                .execute()
        else:
            # No location filter if store name not set properly
            response = supabase.table(ORDERS_TABLE_NAME)\
                .select('*')\
                .eq('printed', False)\
                .order('created_at', desc=True)\
                .execute()

        if response.data and len(response.data) > 0:
            print(f"\nFound {len(response.data)} unprinted order(s) for {store_name or 'all locations'}")

            # Get printer settings from config
            enable_printer = config.get('enable_printer', False) if config else False
            printer_name = config.get('printer_name') if config else None

            for order in response.data:
                print(f"Processing order #{order['order_number']}...")

                # Note: Shipping labels are printed manually via Orders window
                # to allow entering actual package weight

                # Generate and print order PDF
                pdf_path = print_order_pdf(order, send_to_printer=enable_printer, printer_name=printer_name)
                if pdf_path:
                    print(f"[OK] Order #{order['order_number']} processed")
                    if tray_icon:
                        tray_icon.notify(f"Order #{order['order_number']} ready", "New Order")
                else:
                    print(f"[FAIL] Failed to process order #{order['order_number']}")

            return len(response.data)
        return 0
    except Exception as e:
        print(f"Error polling for orders: {e}")
        import traceback
        traceback.print_exc()
        return 0


class ModernButton(tk.Button):
    """Modern styled button widget."""
    def __init__(self, parent, primary=True, **kwargs):
        self.bg_color = COLORS["primary"] if primary else COLORS["accent"]
        self.hover_color = "#ff6b6b" if primary else "#1a4a6e"

        super().__init__(parent,
                        bg=self.bg_color,
                        fg=COLORS["text"],
                        font=("Segoe UI", 13, "bold"),
                        relief="flat",
                        cursor="hand2",
                        activebackground=self.hover_color,
                        activeforeground=COLORS["text"],
                        padx=25,
                        pady=10,
                        **kwargs)

        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, e):
        self.configure(bg=self.hover_color)

    def _on_leave(self, e):
        self.configure(bg=self.bg_color)


class OrdersWindow:
    """Window to display and manage orders"""

    def __init__(self):
        global main_root
        # Use Toplevel of the main root (created in main thread)
        self.root = tk.Toplevel(main_root)

        self.root.title("Orders")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)  # Set minimum window size
        self.root.configure(bg="#1c1c1c")
        self.root.resizable(True, True)  # Allow resizing

        self.create_widgets()

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Defer initial data load to let window render first
        self.root.after(100, self.refresh_orders)

    def create_widgets(self):
        # Styles are configured globally in main() to avoid threading issues

        # Header
        header_frame = ttk.Frame(self.root, padding=(15, 10))
        header_frame.pack(fill=tk.X)

        title_label = ttk.Label(header_frame,
                               text="Orders Management",
                               font=("Segoe UI", 20, "bold"))
        title_label.pack(side=tk.LEFT)

        refresh_btn = ModernButton(header_frame, text="Refresh", primary=False, command=self.refresh_orders)
        refresh_btn.pack(side=tk.RIGHT, padx=5)

        # Treeview Frame
        tree_frame = ttk.Frame(self.root, padding=(10, 5))
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview
        columns = ("Order #", "Date", "Customer", "Location", "Total", "Payment", "Printed")
        self.tree = ttk.Treeview(tree_frame,
                                columns=columns,
                                show="headings",
                                style="Orders.Treeview",
                                yscrollcommand=vsb.set,
                                xscrollcommand=hsb.set)

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # Configure columns
        self.tree.heading("Order #", text="Order #")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Customer", text="Customer")
        self.tree.heading("Location", text="Location")
        self.tree.heading("Total", text="Total")
        self.tree.heading("Payment", text="Payment")
        self.tree.heading("Printed", text="Printed")

        self.tree.column("Order #", width=100, anchor="center", stretch=True)
        self.tree.column("Date", width=160, anchor="center", stretch=True)
        self.tree.column("Customer", width=200, anchor="center", stretch=True)
        self.tree.column("Location", width=100, anchor="center", stretch=True)
        self.tree.column("Total", width=100, anchor="center", stretch=True)
        self.tree.column("Payment", width=100, anchor="center", stretch=True)
        self.tree.column("Printed", width=120, anchor="center", stretch=True)

        self.tree.pack(fill=tk.BOTH, expand=True)

        # Bind selection change to update shipping button state
        self.tree.bind('<<TreeviewSelect>>', self._on_order_select)

        # Button Frame
        button_frame = ttk.Frame(self.root, padding=(20, 15))
        button_frame.pack(fill=tk.X)

        print_btn = ModernButton(button_frame, text="Print Order", primary=True, command=self.print_selected)
        print_btn.pack(side=tk.LEFT, padx=5)

        # Shipping button - starts disabled until order with shipping is selected
        self.shipping_btn = ModernButton(button_frame, text="Print Shipping Label", primary=True, command=self.print_shipping_label)
        self.shipping_btn.pack(side=tk.LEFT, padx=5)
        self.shipping_btn.configure(state='disabled', bg='#555555', cursor='arrow')

        reprint_btn = ModernButton(button_frame, text="Re-Print Order", primary=False, command=self.reprint_selected)
        reprint_btn.pack(side=tk.LEFT, padx=5)

        view_pdf_btn = ModernButton(button_frame, text="View PDF", primary=False, command=self.view_pdf)
        view_pdf_btn.pack(side=tk.LEFT, padx=5)

    def refresh_orders(self):
        """Refresh the orders list (filtered by location)"""
        global supabase, config
        try:
            # Clear existing items
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Get store name for location filtering
            store_name = config.get('store_name', '').lower() if config else ''

            # Build query with location filtering
            if store_name in ['yakima', 'toppenish']:
                response = supabase.table(ORDERS_TABLE_NAME)\
                    .select('*')\
                    .or_(f'order_location.eq.{store_name},order_location.eq.both')\
                    .order('created_at', desc=True)\
                    .limit(100)\
                    .execute()
            else:
                response = supabase.table(ORDERS_TABLE_NAME)\
                    .select('*')\
                    .order('created_at', desc=True)\
                    .limit(100)\
                    .execute()

            if response.data:
                for order in response.data:
                    order_date = datetime.fromisoformat(order['created_at'].replace('Z', '+00:00')).strftime('%m/%d/%Y %I:%M %p')
                    customer = f"{order['customer_first_name']} {order['customer_last_name']}"
                    location = order.get('order_location', 'N/A').capitalize()
                    total = f"${float(order['total']):.2f}"
                    payment_status = order['payment_status'].upper()
                    printed_status = "Printed" if order.get('printed') else "Not Printed"

                    # Insert with tag for color coding
                    tag = "printed" if order.get('printed') else "not_printed"
                    self.tree.insert("", tk.END,
                                   values=(order['order_number'], order_date, customer, location, total, payment_status, printed_status),
                                   tags=(tag, order['id']))

                # Configure tags
                self.tree.tag_configure("printed", foreground=COLORS["text_secondary"])
                self.tree.tag_configure("not_printed", foreground=COLORS["success"])

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load orders: {e}")

    def _on_order_select(self, event=None):
        """Handle order selection change - enable/disable shipping button"""
        global supabase

        selected = self.tree.selection()
        if not selected:
            # No selection - disable shipping button
            self.shipping_btn.configure(state='disabled', bg='#555555', cursor='arrow')
            return

        try:
            # Get order ID from tags
            item_tags = self.tree.item(selected[0])['tags']
            order_id = item_tags[1]

            # Fetch order data to check for shipping items
            response = supabase.table(ORDERS_TABLE_NAME).select('*').eq('id', order_id).execute()

            if response.data and len(response.data) > 0:
                order = response.data[0]

                # Check if order has shipping fulfillment
                has_shipping = False
                for item in order.get('items', []):
                    fulfillment = item.get('fulfillment', {})
                    if fulfillment.get('method') == 'shipping':
                        has_shipping = True
                        break

                # Also check if order has a shipping address
                has_address = bool(order.get('customer_shipping_address'))

                if has_shipping and has_address:
                    # Enable shipping button
                    self.shipping_btn.configure(state='normal', bg=COLORS["primary"], cursor='hand2')
                else:
                    # Disable shipping button
                    self.shipping_btn.configure(state='disabled', bg='#555555', cursor='arrow')
            else:
                self.shipping_btn.configure(state='disabled', bg='#555555', cursor='arrow')

        except Exception as e:
            print(f"Error checking order for shipping: {e}")
            self.shipping_btn.configure(state='disabled', bg='#555555', cursor='arrow')

    def print_selected(self):
        """Print the selected order"""
        global config

        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an order to print")
            return

        # Check if printer is configured
        if not config or not config.get("printer_name"):
            messagebox.showwarning("No Printer Selected", "Please configure a printer in Settings before printing")
            return

        try:
            # Get order ID from tags
            item_tags = self.tree.item(selected[0])['tags']
            order_id = item_tags[1]  # Second tag is the order ID

            # Fetch full order data
            response = supabase.table(ORDERS_TABLE_NAME).select('*').eq('id', order_id).execute()

            if response.data and len(response.data) > 0:
                order = response.data[0]
                # Send to printer with the configured printer name
                printer_name = config.get("printer_name")
                pdf_path = print_order_pdf(order, send_to_printer=True, printer_name=printer_name)
                if pdf_path:
                    messagebox.showinfo("Success", f"Order printed successfully")
                    self.refresh_orders()
                else:
                    messagebox.showerror("Error", "Failed to print order - check printer connection and settings")
            else:
                messagebox.showerror("Error", "Order not found")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to print order: {e}")

    def reprint_selected(self):
        """Re-print the selected order"""
        self.print_selected()

    def view_pdf(self):
        """Open the PDF for the selected order"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an order to view")
            return

        try:
            # Get order ID from tags
            item_tags = self.tree.item(selected[0])['tags']
            order_id = item_tags[1]

            # Fetch order data
            response = supabase.table(ORDERS_TABLE_NAME).select('*').eq('id', order_id).execute()

            if response.data and len(response.data) > 0:
                order = response.data[0]
                pdf_path = order.get('pdf_path')

                if pdf_path and os.path.exists(pdf_path):
                    os.startfile(pdf_path)
                else:
                    messagebox.showinfo("Info", "PDF not found. Would you like to generate it?")
                    self.print_selected()
            else:
                messagebox.showerror("Error", "Order not found")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")

    def print_shipping_label(self):
        """Print FedEx shipping label for selected order with weight input"""
        global config, supabase

        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an order to print a shipping label")
            return

        # Check if FedEx is configured
        if not HAS_FEDEX:
            messagebox.showerror("FedEx Not Available", "FedEx shipping module is not available")
            return

        if not config.get("fedex_api_key") or not config.get("fedex_secret_key"):
            messagebox.showerror("FedEx Not Configured", "Please configure FedEx API credentials in Settings")
            return

        try:
            # Get order ID from tags
            item_tags = self.tree.item(selected[0])['tags']
            order_id = item_tags[1]

            # Fetch order data
            response = supabase.table(ORDERS_TABLE_NAME).select('*').eq('id', order_id).execute()

            if not response.data or len(response.data) == 0:
                messagebox.showerror("Error", "Order not found")
                return

            order = response.data[0]

            # Check if order has shipping address
            if not order.get('customer_shipping_address'):
                messagebox.showerror("No Shipping Address", "This order does not have a shipping address")
                return

            # Check if order already has a tracking number
            existing_tracking = order.get('tracking_number')
            if existing_tracking:
                result = messagebox.askyesno(
                    "Tracking Number Exists",
                    f"This order already has tracking number:\n{existing_tracking}\n\nCreate a new shipping label anyway?"
                )
                if not result:
                    return

            # Create weight input dialog
            weight_dialog = tk.Toplevel(self.root)
            weight_dialog.title("Enter Package Weight")
            weight_dialog.geometry("350x200")
            weight_dialog.configure(bg="#1c1c1c")
            weight_dialog.transient(self.root)
            weight_dialog.grab_set()

            # Center the dialog
            weight_dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (175)
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (100)
            weight_dialog.geometry(f"+{x}+{y}")

            # Dialog content
            tk.Label(weight_dialog, text="Package Weight (lbs):",
                    font=("Segoe UI", 13, "bold"), bg="#1c1c1c", fg="#ffffff").pack(pady=(30, 10))

            weight_var = tk.StringVar(value="1.0")
            weight_entry = ttk.Entry(weight_dialog, textvariable=weight_var, width=15, font=("Segoe UI", 14))
            weight_entry.pack(pady=10, ipady=8)
            weight_entry.focus()
            weight_entry.select_range(0, tk.END)

            tk.Label(weight_dialog, text="Enter the weight after weighing the package",
                    font=("Segoe UI", 10), bg="#1c1c1c", fg="#888888").pack(pady=(5, 15))

            result_holder = {"weight": None}

            def submit_weight():
                try:
                    weight = float(weight_var.get())
                    if weight <= 0:
                        messagebox.showerror("Invalid Weight", "Weight must be greater than 0")
                        return
                    if weight > 150:
                        messagebox.showerror("Invalid Weight", "Weight cannot exceed 150 lbs for FedEx Ground")
                        return
                    result_holder["weight"] = weight
                    weight_dialog.destroy()
                except ValueError:
                    messagebox.showerror("Invalid Weight", "Please enter a valid number")

            def cancel():
                weight_dialog.destroy()

            # Buttons
            btn_frame = tk.Frame(weight_dialog, bg="#1c1c1c")
            btn_frame.pack(pady=10)

            submit_btn = ModernButton(btn_frame, text="Print Label", primary=True, command=submit_weight)
            submit_btn.pack(side=tk.LEFT, padx=10)

            cancel_btn = ModernButton(btn_frame, text="Cancel", primary=False, command=cancel)
            cancel_btn.pack(side=tk.LEFT, padx=10)

            # Bind Enter key
            weight_entry.bind("<Return>", lambda e: submit_weight())
            weight_dialog.bind("<Escape>", lambda e: cancel())

            # Wait for dialog to close
            self.root.wait_window(weight_dialog)

            # Check if weight was entered
            if result_holder["weight"] is None:
                return

            weight = result_holder["weight"]

            # Determine ship-from location
            store_name = config.get('store_name', '').lower()
            if store_name == 'yakima':
                ship_from = "Yakima"
            elif store_name == 'toppenish':
                ship_from = "Toppenish"
            else:
                ship_from = fedex_shipping.get_ship_from_location(order, "Toppenish")

            # Show progress
            self.root.config(cursor="wait")
            self.root.update()

            try:
                # Get FedEx credentials
                api_key = config.get("fedex_api_key")
                secret_key = config.get("fedex_secret_key")
                account_number = config.get("fedex_account_number")
                shipper_addresses = config.get("shipper_addresses", {})
                use_sandbox = config.get("fedex_use_sandbox", False)

                # Get shipper address
                shipper = shipper_addresses.get(ship_from)
                if not shipper:
                    messagebox.showerror("Configuration Error", f"No shipper address configured for {ship_from}")
                    return

                # Build recipient from order
                shipping_address = order.get("customer_shipping_address", {})
                recipient = {
                    "name": f"{order.get('customer_first_name', '')} {order.get('customer_last_name', '')}".strip(),
                    "phone": order.get("customer_phone", ""),
                    "address1": shipping_address.get("address1", ""),
                    "address2": shipping_address.get("address2", ""),
                    "city": shipping_address.get("city", ""),
                    "state": shipping_address.get("state", ""),
                    "zip": shipping_address.get("zipCode", "")
                }

                package_details = {"weight": weight}

                # Get token
                token = fedex_shipping.get_fedex_token(api_key, secret_key, use_sandbox)
                if not token:
                    messagebox.showerror("FedEx Error", "Failed to authenticate with FedEx API")
                    return

                # Create shipment
                result = fedex_shipping.create_shipment(
                    token=token,
                    account_number=account_number,
                    shipper=shipper,
                    recipient=recipient,
                    package_details=package_details,
                    use_sandbox=use_sandbox
                )

                if not result:
                    messagebox.showerror("FedEx Error", "Failed to create shipment. Check the console for details.")
                    return

                tracking_number = result.get("tracking_number")
                label_data = result.get("label_data")

                # Save label PDF
                label_filename = f"shipping_label_{order['order_number']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                label_path = PDF_OUTPUT_DIR / label_filename

                if label_data:
                    fedex_shipping.save_label_pdf(label_data, label_path)

                # Save tracking number to database
                if tracking_number:
                    supabase.table(ORDERS_TABLE_NAME).update({
                        'tracking_number': tracking_number
                    }).eq('id', order['id']).execute()

                # Print the label
                printer_name = config.get('printer_name')
                if printer_name and label_path.exists():
                    print_result = send_pdf_to_printer(str(label_path), printer_name)
                    if print_result:
                        messagebox.showinfo("Success",
                            f"Shipping label printed!\n\nTracking: {tracking_number}\nWeight: {weight} lbs\nShip From: {ship_from}")
                    else:
                        messagebox.showwarning("Partial Success",
                            f"Label created but failed to print.\n\nTracking: {tracking_number}\nLabel saved to: {label_path}")
                else:
                    messagebox.showinfo("Success",
                        f"Shipping label created!\n\nTracking: {tracking_number}\nLabel saved to: {label_path}\n\nNote: No printer configured for auto-print")

                self.refresh_orders()

            finally:
                self.root.config(cursor="")

        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("Error", f"Failed to create shipping label: {e}")
            import traceback
            traceback.print_exc()

    def _on_close(self):
        """Handle window close - clear global reference"""
        global orders_window
        orders_window = None
        self.root.destroy()


class SetupWindow:
    """Setup window for configuration (used for both first-time setup and settings)."""

    def __init__(self, on_complete, existing_config=None):
        self.on_complete = on_complete
        self.existing_config = existing_config
        self.is_settings_mode = existing_config is not None

        # Use Toplevel if settings mode (tray already running), otherwise Tk
        if self.is_settings_mode:
            global main_root
            self.root = tk.Toplevel(main_root)
        else:
            self.root = tk.Tk()
            # Apply theme and styles when creating a new Tk instance (first-time setup)
            sv_ttk.set_theme("dark")
            init_styles()

        self.root.title("Inventory Sync")
        self.root.geometry("900x950")
        self.root.resizable(True, True)

        # Set background to match Sun Valley dark theme
        self.root.configure(bg="#1c1c1c")

        # Override close button to minimize to tray
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        # Override minimize button to also minimize to tray
        self.root.bind("<Unmap>", self.on_minimize)

        if self.is_settings_mode:
            self.root.transient()
            self.root.grab_set()

        self.create_widgets()

        # Bind click event to remove focus from input fields when clicking outside
        self.root.bind("<Button-1>", self._on_click)

    def create_widgets(self):
        # Styles are configured globally in main() to avoid threading issues

        # Create main container frame
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ttk.Label(container,
                               text="Inventory Sync",
                               font=("Segoe UI", 24, "bold"))
        title_label.pack(pady=(0, 5))

        # Subtitle
        subtitle_label = ttk.Label(container,
                                  text="Configure your sync settings",
                                  font=("Segoe UI", 11))
        subtitle_label.pack(pady=(0, 20))

        # Create notebook for tabs
        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True, pady=(0, 20))

        # Create tab frames
        tab1 = ttk.Frame(notebook)
        tab2 = ttk.Frame(notebook)
        tab3 = ttk.Frame(notebook)
        tab4 = ttk.Frame(notebook)

        # Add tabs to notebook
        notebook.add(tab1, text="  Inventory Sync  ")
        notebook.add(tab2, text="  Printer Settings  ")
        notebook.add(tab3, text="  FedEx Shipping  ")
        notebook.add(tab4, text="  Supabase  ")

        # ===== TAB 1: INVENTORY SYNC =====
        tab1_content = ttk.Frame(tab1, padding=30)
        tab1_content.pack(fill=tk.BOTH, expand=True)

        # Store Location (Dropdown)
        store_label = ttk.Label(tab1_content,
                               text="Store Location",
                               font=("Segoe UI", 13, "bold"))
        store_label.pack(anchor="w")

        # Map existing config value to proper case for display
        store_value = self.existing_config.get("store_name", "") if self.existing_config else ""
        if store_value.lower() == "yakima":
            store_value = "Yakima"
        elif store_value.lower() == "toppenish":
            store_value = "Toppenish"
        else:
            store_value = ""

        self.store_var = tk.StringVar(value=store_value)
        self.store_combo = ttk.Combobox(tab1_content,
                                        textvariable=self.store_var,
                                        values=["Yakima", "Toppenish"],
                                        state="readonly",
                                        font=("Segoe UI", 12),
                                        width=50)
        self.store_combo.pack(fill=tk.X, pady=(5, 10), ipady=8)
        # Disable scrolling through options with mouse wheel
        self.store_combo.bind("<MouseWheel>", lambda e: "break")
        self.store_combo.bind('<<ComboboxSelected>>', lambda e: self.root.focus())

        store_help_label = ttk.Label(tab1_content,
                                    text="Select which store this app instance is for",
                                    font=("Segoe UI", 11))
        store_help_label.pack(anchor="w", pady=(0, 15))

        # Watch Folder
        folder_label = ttk.Label(tab1_content,
                                text="Watch Folder",
                                font=("Segoe UI", 13, "bold"))
        folder_label.pack(anchor="w")

        folder_frame = ttk.Frame(tab1_content)
        folder_frame.pack(fill=tk.X, pady=(5, 20))

        folder_value = self.existing_config.get("watch_folder", "") if self.existing_config else str(Path.home() / "Downloads")
        self.folder_var = tk.StringVar(value=folder_value)
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=30)
        self.folder_entry.pack(side=tk.LEFT, ipady=8)
        # Clear selection when field gets focus
        self.folder_entry.bind("<FocusIn>", self._clear_entry_selection)

        select_btn = ModernButton(folder_frame, text="Browse", primary=False, command=self.browse_folder)
        select_btn.pack(side=tk.LEFT, padx=(10, 0))

        # File Pattern
        pattern_label = ttk.Label(tab1_content,
                                 text="File Name Contains",
                                 font=("Segoe UI", 13, "bold"))
        pattern_label.pack(anchor="w")

        pattern_value = self.existing_config.get("file_pattern", "") if self.existing_config else "Inventory by Product"
        self.pattern_var = tk.StringVar(value=pattern_value)
        self.pattern_entry = ttk.Entry(tab1_content, textvariable=self.pattern_var, width=50)
        self.pattern_entry.pack(fill=tk.X, pady=(5, 10), ipady=8)
        # Clear selection when field gets focus
        self.pattern_entry.bind("<FocusIn>", self._clear_entry_selection)

        # Info
        info_label = ttk.Label(tab1_content,
                              text="Watches for .xlsx files containing this text",
                              font=("Segoe UI", 12))
        info_label.pack(anchor="w", pady=(0, 30))

        # Status bar (only show in settings mode when running)
        if self.is_settings_mode:
            status_frame = ttk.Frame(tab1_content)
            status_frame.pack(fill=tk.X, pady=(0, 0))

            status_dot = ttk.Label(status_frame, text="●", font=("Segoe UI", 14))
            status_dot.pack(side=tk.LEFT)

            status_text = ttk.Label(status_frame,
                                   text=f"Running  •  Checking every {POLL_INTERVAL} seconds",
                                   font=("Segoe UI", 11))
            status_text.pack(side=tk.LEFT, padx=(10, 0))

        # ===== TAB 2: PRINTER SETTINGS =====
        tab2_content = ttk.Frame(tab2, padding=30)
        tab2_content.pack(fill=tk.BOTH, expand=True)

        # Enable printer checkbox (proper toggle with Checkbutton)
        enable_printer_value = self.existing_config.get("enable_printer", False) if self.existing_config else False
        self.enable_printer_var = tk.IntVar(value=1 if enable_printer_value else 0)

        enable_printer_check = ttk.Checkbutton(
            tab2_content,
            text="Enable automatic printing to default printer",
            variable=self.enable_printer_var
        )
        enable_printer_check.pack(anchor="w", pady=15)

        # Printer name dropdown
        printer_name_value = self.existing_config.get("printer_name", "") if self.existing_config else ""
        self.printer_name_var = tk.StringVar(value=printer_name_value)

        printer_name_label = ttk.Label(tab2_content,
                                      text="Select Printer",
                                      font=("Segoe UI", 13, "bold"))
        printer_name_label.pack(anchor="w", pady=(20, 5))

        # Get initial printer list
        initial_printers = self._get_printer_list()

        self.printer_combo = ttk.Combobox(tab2_content,
                                          textvariable=self.printer_name_var,
                                          values=initial_printers,
                                          state="readonly",
                                          font=("Segoe UI", 12),
                                          width=50)
        self.printer_combo.pack(fill=tk.X, pady=(5, 10), ipady=8)

        # Bind to dropdown click to refresh printers and unfocus after selection
        self.printer_combo.bind("<Button-1>", self._refresh_printer_list)
        self.printer_combo.bind('<<ComboboxSelected>>', self._on_printer_select)
        # Disable scrolling through printers with mouse wheel
        self.printer_combo.bind("<MouseWheel>", self._disable_printer_scroll)

        # Help text
        printer_help_label = ttk.Label(tab2_content,
                                      text="The dropdown will detect available printers when clicked",
                                      font=("Segoe UI", 11))
        printer_help_label.pack(anchor="w", pady=(0, 25))

        # ===== TAB 3: FEDEX SHIPPING =====
        tab3_content = ttk.Frame(tab3, padding=20)
        tab3_content.pack(fill=tk.BOTH, expand=True)

        # Create scrollable frame for FedEx settings
        fedex_canvas = tk.Canvas(tab3_content, bg="#1c1c1c", highlightthickness=0)
        fedex_scrollbar = ttk.Scrollbar(tab3_content, orient="vertical", command=fedex_canvas.yview)
        fedex_scrollable = ttk.Frame(fedex_canvas)

        fedex_scrollable.bind(
            "<Configure>",
            lambda e: fedex_canvas.configure(scrollregion=fedex_canvas.bbox("all"))
        )

        fedex_canvas.create_window((0, 0), window=fedex_scrollable, anchor="nw")
        fedex_canvas.configure(yscrollcommand=fedex_scrollbar.set)

        fedex_canvas.pack(side="left", fill="both", expand=True)
        fedex_scrollbar.pack(side="right", fill="y")

        # Bind mousewheel to scroll
        def _on_fedex_mousewheel(event):
            fedex_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        fedex_canvas.bind_all("<MouseWheel>", _on_fedex_mousewheel)

        # FedEx API Credentials Section
        creds_label = ttk.Label(fedex_scrollable,
                               text="FedEx API Credentials",
                               font=("Segoe UI", 14, "bold"))
        creds_label.pack(anchor="w", pady=(10, 15))

        # FedEx API Key
        api_key_label = ttk.Label(fedex_scrollable,
                                 text="API Key (Client ID)",
                                 font=("Segoe UI", 12, "bold"))
        api_key_label.pack(anchor="w")

        fedex_api_key_value = self.existing_config.get("fedex_api_key", "") if self.existing_config else ""
        self.fedex_api_key_var = tk.StringVar(value=fedex_api_key_value)
        self.fedex_api_key_entry = ttk.Entry(fedex_scrollable, textvariable=self.fedex_api_key_var, width=60)
        self.fedex_api_key_entry.pack(fill=tk.X, pady=(5, 15), ipady=6)

        # FedEx Secret Key
        secret_key_label = ttk.Label(fedex_scrollable,
                                    text="Secret Key (Client Secret)",
                                    font=("Segoe UI", 12, "bold"))
        secret_key_label.pack(anchor="w")

        fedex_secret_key_value = self.existing_config.get("fedex_secret_key", "") if self.existing_config else ""
        self.fedex_secret_key_var = tk.StringVar(value=fedex_secret_key_value)
        self.fedex_secret_key_entry = ttk.Entry(fedex_scrollable, textvariable=self.fedex_secret_key_var, width=60, show="*")
        self.fedex_secret_key_entry.pack(fill=tk.X, pady=(5, 15), ipady=6)

        # FedEx Account Number
        account_label = ttk.Label(fedex_scrollable,
                                 text="Account Number",
                                 font=("Segoe UI", 12, "bold"))
        account_label.pack(anchor="w")

        fedex_account_value = self.existing_config.get("fedex_account_number", "") if self.existing_config else ""
        self.fedex_account_var = tk.StringVar(value=fedex_account_value)
        self.fedex_account_entry = ttk.Entry(fedex_scrollable, textvariable=self.fedex_account_var, width=60)
        self.fedex_account_entry.pack(fill=tk.X, pady=(5, 15), ipady=6)

        # Sandbox mode checkbox
        sandbox_value = self.existing_config.get("fedex_use_sandbox", False) if self.existing_config else False
        self.fedex_sandbox_var = tk.IntVar(value=1 if sandbox_value else 0)
        sandbox_check = ttk.Checkbutton(
            fedex_scrollable,
            text="Use FedEx Sandbox (for testing)",
            variable=self.fedex_sandbox_var
        )
        sandbox_check.pack(anchor="w", pady=(0, 20))

        # Separator
        ttk.Separator(fedex_scrollable, orient="horizontal").pack(fill="x", pady=10)

        # Shipper Addresses Section
        shipper_label = ttk.Label(fedex_scrollable,
                                 text="Shipper Addresses",
                                 font=("Segoe UI", 14, "bold"))
        shipper_label.pack(anchor="w", pady=(10, 15))

        # Get existing shipper addresses
        shipper_addresses = self.existing_config.get("shipper_addresses", {}) if self.existing_config else {}

        # Yakima Address
        yakima_frame = tk.LabelFrame(fedex_scrollable, text="  Yakima Location  ",
                                     bg="#1c1c1c", fg="#ffffff", font=("Segoe UI", 12, "bold"),
                                     bd=1, relief="solid", highlightbackground="#3a3a3a",
                                     highlightcolor="#3a3a3a", highlightthickness=1,
                                     padx=15, pady=15)
        yakima_frame.pack(fill=tk.X, pady=(0, 15))

        yakima_addr = shipper_addresses.get("Yakima", {})

        tk.Label(yakima_frame, text="Company Name", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(anchor="w")
        self.yakima_company_var = tk.StringVar(value=yakima_addr.get("company", ""))
        ttk.Entry(yakima_frame, textvariable=self.yakima_company_var, width=50).pack(fill=tk.X, pady=(2, 10), ipady=4)

        tk.Label(yakima_frame, text="Street Address", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(anchor="w")
        self.yakima_street_var = tk.StringVar(value=yakima_addr.get("street", ""))
        ttk.Entry(yakima_frame, textvariable=self.yakima_street_var, width=50).pack(fill=tk.X, pady=(2, 10), ipady=4)

        # City, State, Zip row
        yakima_csz_frame = tk.Frame(yakima_frame, bg="#1c1c1c")
        yakima_csz_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(yakima_csz_frame, text="City", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(side=tk.LEFT)
        self.yakima_city_var = tk.StringVar(value=yakima_addr.get("city", "Yakima"))
        ttk.Entry(yakima_csz_frame, textvariable=self.yakima_city_var, width=20).pack(side=tk.LEFT, padx=(5, 15), ipady=4)

        tk.Label(yakima_csz_frame, text="State", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(side=tk.LEFT)
        self.yakima_state_var = tk.StringVar(value=yakima_addr.get("state", "WA"))
        ttk.Entry(yakima_csz_frame, textvariable=self.yakima_state_var, width=5).pack(side=tk.LEFT, padx=(5, 15), ipady=4)

        tk.Label(yakima_csz_frame, text="ZIP", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(side=tk.LEFT)
        self.yakima_zip_var = tk.StringVar(value=yakima_addr.get("zip", ""))
        ttk.Entry(yakima_csz_frame, textvariable=self.yakima_zip_var, width=10).pack(side=tk.LEFT, padx=(5, 0), ipady=4)

        tk.Label(yakima_frame, text="Phone", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(anchor="w")
        self.yakima_phone_var = tk.StringVar(value=yakima_addr.get("phone", ""))
        ttk.Entry(yakima_frame, textvariable=self.yakima_phone_var, width=20).pack(anchor="w", pady=(2, 0), ipady=4)

        # Toppenish Address
        toppenish_frame = tk.LabelFrame(fedex_scrollable, text="  Toppenish Location  ",
                                        bg="#1c1c1c", fg="#ffffff", font=("Segoe UI", 12, "bold"),
                                        bd=1, relief="solid", highlightbackground="#3a3a3a",
                                        highlightcolor="#3a3a3a", highlightthickness=1,
                                        padx=15, pady=15)
        toppenish_frame.pack(fill=tk.X, pady=(0, 15))

        toppenish_addr = shipper_addresses.get("Toppenish", {})

        tk.Label(toppenish_frame, text="Company Name", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(anchor="w")
        self.toppenish_company_var = tk.StringVar(value=toppenish_addr.get("company", ""))
        ttk.Entry(toppenish_frame, textvariable=self.toppenish_company_var, width=50).pack(fill=tk.X, pady=(2, 10), ipady=4)

        tk.Label(toppenish_frame, text="Street Address", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(anchor="w")
        self.toppenish_street_var = tk.StringVar(value=toppenish_addr.get("street", ""))
        ttk.Entry(toppenish_frame, textvariable=self.toppenish_street_var, width=50).pack(fill=tk.X, pady=(2, 10), ipady=4)

        # City, State, Zip row
        toppenish_csz_frame = tk.Frame(toppenish_frame, bg="#1c1c1c")
        toppenish_csz_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(toppenish_csz_frame, text="City", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(side=tk.LEFT)
        self.toppenish_city_var = tk.StringVar(value=toppenish_addr.get("city", "Toppenish"))
        ttk.Entry(toppenish_csz_frame, textvariable=self.toppenish_city_var, width=20).pack(side=tk.LEFT, padx=(5, 15), ipady=4)

        tk.Label(toppenish_csz_frame, text="State", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(side=tk.LEFT)
        self.toppenish_state_var = tk.StringVar(value=toppenish_addr.get("state", "WA"))
        ttk.Entry(toppenish_csz_frame, textvariable=self.toppenish_state_var, width=5).pack(side=tk.LEFT, padx=(5, 15), ipady=4)

        tk.Label(toppenish_csz_frame, text="ZIP", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(side=tk.LEFT)
        self.toppenish_zip_var = tk.StringVar(value=toppenish_addr.get("zip", ""))
        ttk.Entry(toppenish_csz_frame, textvariable=self.toppenish_zip_var, width=10).pack(side=tk.LEFT, padx=(5, 0), ipady=4)

        tk.Label(toppenish_frame, text="Phone", font=("Segoe UI", 11), bg="#1c1c1c", fg="#ffffff").pack(anchor="w")
        self.toppenish_phone_var = tk.StringVar(value=toppenish_addr.get("phone", ""))
        ttk.Entry(toppenish_frame, textvariable=self.toppenish_phone_var, width=20).pack(anchor="w", pady=(2, 0), ipady=4)

        # Info note
        info_note = ttk.Label(fedex_scrollable,
                             text="Get FedEx API credentials at developer.fedex.com",
                             font=("Segoe UI", 10))
        info_note.pack(anchor="w", pady=(10, 20))

        # ===== TAB 4: SUPABASE =====
        tab4_content = ttk.Frame(tab4, padding=30)
        tab4_content.pack(fill=tk.BOTH, expand=True)

        # Supabase URL
        url_label = ttk.Label(tab4_content,
                             text="Supabase URL",
                             font=("Segoe UI", 13, "bold"))
        url_label.pack(anchor="w")

        url_value = self.existing_config.get("supabase_url", "") if self.existing_config else ""
        self.url_var = tk.StringVar(value=url_value)
        self.url_entry = ttk.Entry(tab4_content, textvariable=self.url_var, width=50)
        self.url_entry.pack(fill=tk.X, pady=(5, 25), ipady=8)
        # Clear selection when field gets focus
        self.url_entry.bind("<FocusIn>", self._clear_entry_selection)

        # Supabase Key
        key_label = ttk.Label(tab4_content,
                             text="Supabase Key",
                             font=("Segoe UI", 13, "bold"))
        key_label.pack(anchor="w")

        key_value = self.existing_config.get("supabase_key", "") if self.existing_config else ""
        self.key_var = tk.StringVar(value=key_value)
        self.key_entry = ttk.Entry(tab4_content, textvariable=self.key_var, width=50, show="*")
        self.key_entry.pack(fill=tk.X, pady=(5, 10), ipady=8)
        # Clear selection when field gets focus
        self.key_entry.bind("<FocusIn>", self._clear_entry_selection)

        # ===== BOTTOM SECTION (OUTSIDE TABS) =====
        # Save Button
        save_btn = ModernButton(container, text="Save Settings", primary=True, command=self.save_and_start)
        save_btn.pack()

    def _clear_entry_selection(self, event=None):
        """Clear text selection/highlight from entry fields"""
        if event and hasattr(event.widget, 'selection_clear'):
            event.widget.selection_clear()

    def _on_click(self, event):
        """Remove focus from input fields when clicking outside them."""
        # Get the widget that was clicked
        widget = event.widget

        # Check if the clicked widget is an input field (Entry)
        if not isinstance(widget, (ttk.Entry, tk.Entry)):
            # Click was outside input fields, remove focus
            self.root.focus()

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get())
        if folder:
            self.folder_var.set(folder)

    def _get_printer_list(self):
        """Get list of available printers"""
        try:
            printers = []

            if HAS_WIN32:
                try:
                    printer_data = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
                    printers = [p[2] for p in printer_data]
                except:
                    pass

            if not printers:
                # Fallback: try to get default printer
                default = get_default_printer()
                if default:
                    printers = [default]
                else:
                    printers = ["Default Printer"]

            return printers
        except Exception as e:
            return ["Error detecting printers"]

    def _refresh_printer_list(self, event=None):
        """Refresh the printer list when dropdown is clicked"""
        printers = self._get_printer_list()
        self.printer_combo['values'] = printers

    def _on_printer_select(self, event=None):
        """Handle printer selection - unfocus dropdown and clear highlight"""
        # Clear any text selection in the combobox
        self.printer_combo.selection_clear()
        # Remove focus from combobox
        self.root.focus()

    def _disable_printer_scroll(self, event=None):
        """Disable scrolling through printers with mouse wheel"""
        return "break"  # Prevent default scroll behavior

    def detect_printers(self):
        """Detect available printers"""
        try:
            printers = []

            if HAS_WIN32:
                try:
                    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
                    printer_names = [p[2] for p in printers]
                except:
                    pass

            if not printers:
                # Fallback: try to get default printer
                default = get_default_printer()
                if default:
                    printer_names = [default]
                else:
                    printer_names = ["No printers detected"]

            # Show printer selection dialog
            printer_window = tk.Toplevel(self.root)
            printer_window.title("Select Printer")
            printer_window.geometry("400x300")
            printer_window.configure(bg=COLORS["bg"])

            label = tk.Label(printer_window,
                           text="Available Printers:",
                           font=("Segoe UI", 13, "bold"),
                           bg=COLORS["bg"],
                           fg=COLORS["text"])
            label.pack(pady=10)

            # Listbox
            listbox = tk.Listbox(printer_window,
                                bg=COLORS["secondary_bg"],
                                fg=COLORS["text"],
                                font=("Segoe UI", 12),
                                height=10)
            listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

            for printer in printer_names:
                listbox.insert(tk.END, printer)

            def select_printer():
                selection = listbox.curselection()
                if selection:
                    self.printer_name_var.set(printer_names[selection[0]])
                    printer_window.destroy()

            btn = ModernButton(printer_window, text="Select", primary=True, command=select_printer)
            btn.pack(pady=10)

            messagebox.showinfo("Printers", f"Found {len(printer_names)} printer(s)")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to detect printers: {e}")

    def validate_all_fields(self):
        """Validate all fields and show errors."""
        errors = []

        store_value = self.store_var.get().strip()
        if not store_value or store_value not in ["Yakima", "Toppenish"]:
            errors.append("Please select a Store Location (Yakima or Toppenish)")

        watch_folder = self.folder_var.get()
        if not watch_folder or not os.path.isdir(watch_folder):
            errors.append("Please select a valid folder")

        pattern_value = self.pattern_var.get().strip()
        if not pattern_value or not any(c.isalnum() for c in pattern_value):
            errors.append("File Name pattern is required")

        url_value = self.url_var.get().strip()
        if not url_value or not any(c.isalnum() for c in url_value):
            errors.append("Supabase URL is required")

        key_value = self.key_var.get().strip()
        if not key_value or not any(c.isalnum() for c in key_value):
            errors.append("Supabase Key is required")

        if errors:
            messagebox.showerror("Missing Information", "\n".join(errors))
            return False
        return True

    def _save_settings(self):
        """Save settings and close window."""
        if not self.validate_all_fields():
            return

        # Convert IntVar to boolean for printer setting
        enable_printer = bool(self.enable_printer_var.get())
        fedex_use_sandbox = bool(self.fedex_sandbox_var.get())

        # Build shipper addresses
        shipper_addresses = {
            "Yakima": {
                "company": self.yakima_company_var.get(),
                "street": self.yakima_street_var.get(),
                "city": self.yakima_city_var.get(),
                "state": self.yakima_state_var.get(),
                "zip": self.yakima_zip_var.get(),
                "phone": self.yakima_phone_var.get()
            },
            "Toppenish": {
                "company": self.toppenish_company_var.get(),
                "street": self.toppenish_street_var.get(),
                "city": self.toppenish_city_var.get(),
                "state": self.toppenish_state_var.get(),
                "zip": self.toppenish_zip_var.get(),
                "phone": self.toppenish_phone_var.get()
            }
        }

        config = save_config(
            self.store_var.get(),
            self.folder_var.get(),
            self.pattern_var.get(),
            self.url_var.get(),
            self.key_var.get(),
            enable_printer,
            self.printer_name_var.get() if self.printer_name_var.get() else None,
            self.fedex_api_key_var.get() if self.fedex_api_key_var.get() else None,
            self.fedex_secret_key_var.get() if self.fedex_secret_key_var.get() else None,
            self.fedex_account_var.get() if self.fedex_account_var.get() else None,
            shipper_addresses,
            fedex_use_sandbox
        )
        self.root.destroy()
        self.on_complete(config)

    def on_minimize(self, event):
        """Handle minimize button - save and go to tray instead."""
        if self.root.state() == 'iconic':
            self.root.deiconify()  # Restore first to prevent weird state
            self.minimize_to_tray()

    def minimize_to_tray(self):
        """Save settings and minimize to system tray."""
        global settings_window
        settings_window = None
        self._save_settings()

    def save_and_start(self):
        """Save settings and start the application."""
        self._save_settings()

    def run(self):
        self.root.mainloop()


def request_show_settings():
    """Request to show settings window (called from tray menu thread)."""
    global pending_action
    pending_action = "settings"


def request_show_orders():
    """Request to show orders window (called from tray menu thread)."""
    global pending_action
    pending_action = "orders"


def do_show_settings():
    """Actually show the settings window (called from main thread)."""
    global config, settings_window, orders_window, main_root

    # If settings window is already open, bring it to focus
    if settings_window is not None:
        try:
            settings_window.root.lift()
            settings_window.root.focus_force()
            return
        except:
            settings_window = None

    # Only allow one window at a time - if orders is open, close it first
    if orders_window is not None:
        try:
            orders_window.root.destroy()
        except:
            pass
        orders_window = None

    def on_save(new_config):
        global config
        config = new_config

    settings = SetupWindow(on_save, existing_config=config)
    settings_window = settings


def do_show_orders():
    """Actually show the orders window (called from main thread)."""
    global orders_window, settings_window

    # If orders window is already open, bring it to focus
    if orders_window is not None:
        try:
            orders_window.root.lift()
            orders_window.root.focus_force()
            return
        except:
            orders_window = None

    # Only allow one window at a time - if settings is open, close it first
    if settings_window is not None:
        try:
            settings_window.root.destroy()
        except:
            pass
        settings_window = None

    # Create orders window as Toplevel of main_root
    orders_window = OrdersWindow()


def sync_now():
    """Manually trigger a sync."""
    global config
    if config:
        threading.Thread(
            target=sync_inventory,
            args=(config["watch_folder"], config["file_pattern"]),
            daemon=True
        ).start()


def quit_app(icon):
    """Quit the application."""
    global main_root
    stop_polling()
    icon.stop()
    # Destroy main root to exit tkinter mainloop
    if main_root:
        try:
            main_root.quit()
            main_root.destroy()
        except:
            pass


def check_pending_actions():
    """Check for pending actions from tray menu and execute them in main thread."""
    global pending_action, main_root

    if pending_action == "settings":
        pending_action = None
        do_show_settings()
    elif pending_action == "orders":
        pending_action = None
        do_show_orders()
    elif pending_action and isinstance(pending_action, tuple) and pending_action[0] == "update":
        _, latest_version, download_url = pending_action
        pending_action = None
        prompt_update(latest_version, download_url)

    # Schedule next check
    if main_root:
        main_root.after(100, check_pending_actions)


def prompt_update(latest_version, download_url):
    """Show update dialog to the user (must be called from main thread)."""
    answer = messagebox.askyesno(
        "Update Available",
        f"A new version of Inventory Sync is available!\n\n"
        f"Current version: {APP_VERSION}\n"
        f"New version: {latest_version}\n\n"
        f"Would you like to update now?",
        parent=main_root
    )

    if not answer:
        return

    # Show progress window
    progress_win = tk.Toplevel(main_root)
    progress_win.title("Updating...")
    progress_win.geometry("350x120")
    progress_win.resizable(False, False)
    progress_win.transient(main_root)
    progress_win.grab_set()

    tk.Label(progress_win, text="Downloading update...", font=("Segoe UI", 11)).pack(pady=(15, 5))
    progress_var = tk.DoubleVar(value=0)
    progress_bar = ttk.Progressbar(progress_win, variable=progress_var, maximum=100, length=280)
    progress_bar.pack(pady=10, padx=20)
    status_label = tk.Label(progress_win, text="0%", font=("Segoe UI", 9))
    status_label.pack()

    def do_download():
        def on_progress(downloaded, total):
            pct = (downloaded / total) * 100
            progress_var.set(pct)
            try:
                status_label.config(text=f"{pct:.0f}% ({downloaded // 1024} KB / {total // 1024} KB)")
            except tk.TclError:
                pass

        temp_path = auto_updater.download_update(download_url, progress_callback=on_progress)

        if temp_path and auto_updater.apply_update(temp_path):
            try:
                progress_win.destroy()
            except tk.TclError:
                pass
            # Quit the app so the updater batch script can replace the exe
            if tray_icon:
                tray_icon.stop()
            if main_root:
                main_root.quit()
                main_root.destroy()
            sys.exit(0)
        else:
            try:
                progress_win.destroy()
            except tk.TclError:
                pass
            messagebox.showerror(
                "Update Failed",
                "Could not download or apply the update.\nPlease try again later.",
                parent=main_root
            )

    threading.Thread(target=do_download, daemon=True).start()


def on_update_available(latest_version, download_url):
    """Callback from auto_updater when an update is found (called from background thread)."""
    global pending_action
    pending_action = ("update", latest_version, download_url)


def run_tray(config_data):
    """Run the system tray application."""
    global tray_icon, config, PDF_OUTPUT_DIR, main_root
    config = config_data

    # Initialize Supabase with credentials from config
    if not init_supabase(config.get("supabase_url"), config.get("supabase_key")):
        print("Error: Could not initialize Supabase. Check your credentials.")
        return

    # Create PDF output directory
    PDF_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Create hidden main root for tkinter operations (must be in main thread)
    main_root = tk.Tk()
    main_root.withdraw()

    # Apply theme and styles to main root
    sv_ttk.set_theme("dark")
    init_styles()

    # Start polling thread
    threading.Thread(target=polling_loop, daemon=True).start()

    # Check for existing files immediately
    threading.Thread(
        target=sync_inventory,
        args=(config["watch_folder"], config["file_pattern"]),
        daemon=True
    ).start()

    # Create system tray icon
    icon_image = create_tray_icon()

    menu = pystray.Menu(
        item('Inventory Sync', None, enabled=False),
        item('─────────────', None, enabled=False),
        item('Sync Now', lambda: sync_now()),
        item('View Orders', lambda: request_show_orders()),
        item('Settings', lambda: request_show_settings()),
        item('─────────────', None, enabled=False),
        item('Exit', quit_app)
    )

    tray_icon = pystray.Icon("inventory_sync", icon_image, "Inventory Sync", menu)

    print("\n" + "="*50)
    print("Inventory Sync is running in the system tray")
    print("="*50)
    print(f"Watching: {config['watch_folder']}")
    print(f"Pattern: {config['file_pattern']}")
    print(f"Polling orders from Supabase")
    print(f"PDF output: {PDF_OUTPUT_DIR}")
    print(f"Checking every {POLL_INTERVAL} seconds")
    print("\nRight-click the tray icon for options:")
    print("  • Sync Now - Manually trigger inventory sync")
    print("  • View Orders - See all orders and their print status")
    print("  • Settings - Update configuration")
    print("="*50)

    # Start checking for pending actions
    main_root.after(100, check_pending_actions)

    # Check for updates in background
    if HAS_UPDATER:
        auto_updater.run_update_check(APP_VERSION, on_update_available)

    # Run tray icon in background thread
    threading.Thread(target=tray_icon.run, daemon=True).start()

    # Run tkinter main loop (this keeps the main thread alive and handles GUI)
    main_root.mainloop()


def init_styles():
    """Initialize all ttk styles globally in the main thread"""
    style = ttk.Style()
    theme_bg = "#1c1c1c"
    theme_fg = "#ffffff"

    # Configure base styles
    style.configure("TFrame", background=theme_bg)
    style.configure("TLabel", background=theme_bg, foreground=theme_fg, font=("Segoe UI", 11))
    style.configure("TCheckbutton", background=theme_bg, foreground=theme_fg, font=("Segoe UI", 11))
    style.configure("TRadiobutton", font=("Segoe UI", 12))
    style.configure("TNotebook", background=theme_bg)
    style.configure("TNotebook.Tab", font=("Segoe UI", 13), padding=[15, 12], foreground=theme_fg)
    style.configure("TEntry", font=("Segoe UI", 12), padding=8)
    style.configure("TCombobox", font=("Segoe UI", 12), padding=8)
    style.configure("TSeparator", background="#3a3a3a")

    # Configure LabelFrame styling for dark theme (for any ttk.LabelFrame usage)
    style.configure("TLabelframe", background=theme_bg)
    style.configure("TLabelframe.Label", background=theme_bg, foreground=theme_fg, font=("Segoe UI", 12, "bold"))

    # Configure Orders Treeview style
    style.configure("Orders.Treeview",
                   background="#2a2a2a",
                   foreground="#ffffff",
                   fieldbackground="#2a2a2a",
                   borderwidth=0,
                   rowheight=28,
                   font=("Segoe UI", 11))
    style.configure("Orders.Treeview.Heading",
                   background="#1f6aa0",
                   foreground="#ffffff",
                   borderwidth=0,
                   font=("Segoe UI", 10, "bold"))
    style.map("Orders.Treeview",
             background=[("selected", "#0f3460")],
             foreground=[("selected", "#ffffff")])


def check_and_install():
    """Check if running from install location, if not install and relaunch."""
    if not getattr(sys, 'frozen', False):
        return  # Not running as exe, skip installation

    install_dir = Path(os.environ['LOCALAPPDATA']) / 'InventorySync'
    installed_exe = install_dir / 'InventorySync.exe'
    current_exe = Path(sys.executable)

    # Check if already running from install location
    try:
        if current_exe.resolve() == installed_exe.resolve():
            return  # Already installed, continue normally
    except:
        pass

    # Also check by parent directory
    if current_exe.parent == install_dir:
        return  # Already installed

    # First run - install the application
    try:
        install_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy(current_exe, installed_exe)

        # Add to Windows startup registry
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(key, "InventorySync", 0, winreg.REG_SZ, str(installed_exe))
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Warning: Could not add to startup: {e}")

        # Launch the installed version
        subprocess.Popen([str(installed_exe)])
        sys.exit(0)

    except Exception as e:
        print(f"Installation error: {e}")
        # Continue running from current location if install fails


def main():
    # Check and install if needed (only when running as exe)
    check_and_install()

    config = load_config()

    if config:
        # Config exists - go straight to tray mode
        # Theme/styles will be applied when windows are opened later
        run_tray(config)
    else:
        # No config - show initial setup window
        # SetupWindow will apply theme/styles since it creates a new Tk
        def on_setup_complete(new_config):
            run_tray(new_config)

        setup = SetupWindow(on_setup_complete)
        setup.run()


if __name__ == "__main__":
    main()
