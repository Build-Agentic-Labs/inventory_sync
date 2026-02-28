"""
Order Printing System - Supabase Polling Script
This script polls Supabase for new orders and prints them as PDF.
Add this to your existing inventory script or run it separately.
"""

import os
import time
import json
from datetime import datetime
from pathlib import Path
from supabase import create_client, Client
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Supabase configuration
SUPABASE_URL = os.getenv("SUPABASE_URL", "your-supabase-url")
SUPABASE_KEY = os.getenv("SUPABASE_SERVICE_KEY", "your-supabase-service-key")

# Polling interval in seconds
POLL_INTERVAL = 15  # Check every 15 seconds

# PDF output directory
PDF_OUTPUT_DIR = Path(__file__).parent / "order_pdfs"
PDF_OUTPUT_DIR.mkdir(exist_ok=True)


def init_supabase() -> Client:
    """Initialize Supabase client"""
    return create_client(SUPABASE_URL, SUPABASE_KEY)


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
    """
    Generate a professional PDF order form for workers
    """
    c = canvas.Canvas(str(pdf_path), pagesize=letter)
    width, height = letter

    # Colors
    primary_color = HexColor("#e94560")
    secondary_color = HexColor("#0f3460")
    text_color = black
    gray = HexColor("#666666")
    light_gray = HexColor("#f5f5f5")

    y_position = height - 0.75 * inch

    # Header - Order Number
    c.setFillColor(primary_color)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(0.75 * inch, y_position, f"ORDER #{order['order_number']}")

    y_position -= 0.4 * inch
    c.setFillColor(text_color)
    c.setFont("Helvetica", 10)
    order_date = datetime.fromisoformat(order['created_at'].replace('Z', '+00:00')).strftime('%B %d, %Y at %I:%M %p')
    c.drawString(0.75 * inch, y_position, f"Received: {order_date}")

    # Payment Status Badge
    c.setFillColor(HexColor("#4ecca3") if order['payment_status'] == 'paid' else HexColor("#ff6b6b"))
    c.setFont("Helvetica-Bold", 10)
    status_text = order['payment_status'].upper()
    c.drawString(width - 1.5 * inch, y_position, status_text)

    y_position -= 0.6 * inch

    # Customer Information Section
    c.setFillColor(secondary_color)
    c.rect(0.75 * inch, y_position - 0.2 * inch, width - 1.5 * inch, 0.3 * inch, fill=True, stroke=False)
    c.setFillColor(HexColor("#ffffff"))
    c.setFont("Helvetica-Bold", 12)
    c.drawString(0.85 * inch, y_position - 0.05 * inch, "CUSTOMER INFORMATION")

    y_position -= 0.5 * inch
    c.setFillColor(text_color)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.85 * inch, y_position, "Name:")
    c.setFont("Helvetica", 11)
    c.drawString(1.5 * inch, y_position, f"{order['customer_first_name']} {order['customer_last_name']}")

    y_position -= 0.25 * inch
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.85 * inch, y_position, "Email:")
    c.setFont("Helvetica", 11)
    c.drawString(1.5 * inch, y_position, order['customer_email'])

    y_position -= 0.25 * inch
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.85 * inch, y_position, "Phone:")
    c.setFont("Helvetica", 11)
    c.drawString(1.5 * inch, y_position, order.get('customer_phone', 'N/A'))

    # Shipping Address (if provided)
    if order.get('customer_shipping_address'):
        addr = order['customer_shipping_address']
        y_position -= 0.35 * inch
        c.setFont("Helvetica-Bold", 11)
        c.drawString(0.85 * inch, y_position, "Shipping Address:")
        y_position -= 0.2 * inch
        c.setFont("Helvetica", 10)
        c.drawString(1.5 * inch, y_position, addr.get('address1', ''))
        if addr.get('address2'):
            y_position -= 0.15 * inch
            c.drawString(1.5 * inch, y_position, addr['address2'])
        y_position -= 0.15 * inch
        c.drawString(1.5 * inch, y_position, f"{addr.get('city', '')}, {addr.get('state', '')} {addr.get('zipCode', '')}")

    y_position -= 0.6 * inch

    # Order Items Section
    c.setFillColor(secondary_color)
    c.rect(0.75 * inch, y_position - 0.2 * inch, width - 1.5 * inch, 0.3 * inch, fill=True, stroke=False)
    c.setFillColor(HexColor("#ffffff"))
    c.setFont("Helvetica-Bold", 12)
    c.drawString(0.85 * inch, y_position - 0.05 * inch, "ORDER ITEMS - FULFILLMENT INSTRUCTIONS")

    y_position -= 0.5 * inch

    # Draw items with fulfillment details
    for idx, item in enumerate(order['items']):
        # Item background (alternating)
        if idx % 2 == 0:
            c.setFillColor(light_gray)
            c.rect(0.75 * inch, y_position - 0.8 * inch, width - 1.5 * inch, 0.9 * inch, fill=True, stroke=False)

        # Item details
        c.setFillColor(text_color)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(0.85 * inch, y_position, f"{item['quantity']}x {item['name']}")

        y_position -= 0.25 * inch
        c.setFont("Helvetica", 10)
        c.drawString(0.95 * inch, y_position, f"Price: ${item['price']:.2f} each  |  Subtotal: ${item['price'] * item['quantity']:.2f}")

        # Fulfillment - LARGE AND BOLD for worker clarity
        y_position -= 0.3 * inch
        fulfillment_text = format_fulfillment(item)
        c.setFillColor(primary_color)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(0.95 * inch, y_position, f"ACTION REQUIRED: {fulfillment_text}")

        # Additional fulfillment details
        fulfillment = item.get('fulfillment', {})
        method = fulfillment.get('method', 'N/A')

        y_position -= 0.2 * inch
        c.setFillColor(gray)
        c.setFont("Helvetica", 9)

        if method == 'pickup':
            location = fulfillment.get('location', 'Unknown')
            location_name = "Yakima" if location == 1 else "Toppenish"
            c.drawString(0.95 * inch, y_position, f"→ Prepare for customer pickup at {location_name} location")
        elif method == 'delivery':
            address = fulfillment.get('address', {})
            c.drawString(0.95 * inch, y_position, f"→ Deliver to: {address.get('street', '')}, {address.get('city', '')}")
        elif method == 'shipping':
            address = fulfillment.get('address', {})
            c.drawString(0.95 * inch, y_position, f"→ Ship to: {address.get('city', '')}, {address.get('state', '')} {address.get('zipCode', '')}")
            if item.get('shippingCost'):
                y_position -= 0.15 * inch
                c.drawString(0.95 * inch, y_position, f"   Shipping cost: ${item['shippingCost']:.2f}")

        y_position -= 0.5 * inch

        # Check if we need a new page
        if y_position < 2 * inch:
            c.showPage()
            y_position = height - inch

    # Totals Section
    y_position -= 0.3 * inch
    c.setStrokeColor(gray)
    c.setLineWidth(1)
    c.line(0.75 * inch, y_position, width - 0.75 * inch, y_position)

    y_position -= 0.3 * inch
    c.setFillColor(text_color)
    c.setFont("Helvetica", 11)
    c.drawRightString(width - 1.5 * inch, y_position, "Subtotal:")
    c.drawRightString(width - 0.75 * inch, y_position, f"${float(order['subtotal']):.2f}")

    y_position -= 0.2 * inch
    c.drawRightString(width - 1.5 * inch, y_position, "Shipping:")
    c.drawRightString(width - 0.75 * inch, y_position, f"${float(order['shipping_cost']):.2f}")

    y_position -= 0.2 * inch
    c.drawRightString(width - 1.5 * inch, y_position, "Tax (8.5%):")
    c.drawRightString(width - 0.75 * inch, y_position, f"${float(order['tax_amount']):.2f}")

    y_position -= 0.3 * inch
    c.setFont("Helvetica-Bold", 13)
    c.drawRightString(width - 1.5 * inch, y_position, "TOTAL:")
    c.drawRightString(width - 0.75 * inch, y_position, f"${float(order['total']):.2f}")

    # Footer
    c.setFont("Helvetica", 8)
    c.setFillColor(gray)
    c.drawCentredString(width / 2, 0.5 * inch, f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    c.save()
    return pdf_path


def print_order(order):
    """
    Generate PDF and optionally send to printer
    """
    # Generate PDF
    pdf_filename = f"order_{order['order_number']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    pdf_path = PDF_OUTPUT_DIR / pdf_filename

    try:
        create_pdf_order(order, pdf_path)
        print(f"✓ PDF created: {pdf_path}")

        # TODO: Add actual printer code here when ready
        # from win32print import GetDefaultPrinter
        # import subprocess
        # subprocess.run(['print', '/D:"%s"' % GetDefaultPrinter(), str(pdf_path)], shell=True)

        return True
    except Exception as e:
        print(f"✗ Error creating PDF: {e}")
        return False


def mark_order_printed(supabase: Client, order_id: str, pdf_path: str = None):
    """Mark order as printed in Supabase"""
    try:
        update_data = {
            'printed': True,
            'printed_at': datetime.utcnow().isoformat()
        }

        # Try to update with pdf_path if provided
        if pdf_path:
            try:
                update_data['pdf_path'] = str(pdf_path)
                result = supabase.table('orders').update(update_data).eq('id', order_id).execute()
            except Exception as db_error:
                # If pdf_path column doesn't exist, update without it
                if 'pdf_path' in str(db_error):
                    del update_data['pdf_path']
                    result = supabase.table('orders').update(update_data).eq('id', order_id).execute()
                else:
                    raise
        else:
            result = supabase.table('orders').update(update_data).eq('id', order_id).execute()

        return True
    except Exception as e:
        print(f"Error marking order as printed: {e}")
        return False


def poll_for_orders(supabase: Client):
    """Poll Supabase for unprinted orders"""
    try:
        # Get all unprinted orders
        response = supabase.table('orders').select('*').eq('printed', False).order('created_at', desc=True).execute()

        if response.data and len(response.data) > 0:
            print(f"\n{'='*50}")
            print(f"Found {len(response.data)} unprinted order(s)")
            print(f"{'='*50}")

            for order in response.data:
                print(f"\nProcessing order #{order['order_number']}...")

                # Generate PDF
                pdf_filename = f"order_{order['order_number']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                pdf_path = PDF_OUTPUT_DIR / pdf_filename

                if create_pdf_order(order, pdf_path):
                    # Mark as printed with PDF path
                    if mark_order_printed(supabase, order['id'], str(pdf_path)):
                        print(f"[OK] Order #{order['order_number']} printed and marked as complete")
                        print(f"  PDF saved to: {pdf_path}")
                    else:
                        print(f"[FAIL] Failed to mark order #{order['order_number']} as printed")
                else:
                    print(f"[FAIL] Failed to generate PDF for order #{order['order_number']}")

        else:
            # Uncomment to see polling activity
            # print(f"[{datetime.now().strftime('%H:%M:%S')}] No new orders")
            pass

    except Exception as e:
        print(f"Error polling for orders: {e}")


def main():
    """Main polling loop"""
    print("Order Printing System Started")
    print(f"Polling interval: {POLL_INTERVAL} seconds")
    print("Press Ctrl+C to stop\n")

    supabase = init_supabase()

    try:
        while True:
            poll_for_orders(supabase)
            time.sleep(POLL_INTERVAL)
    except KeyboardInterrupt:
        print("\n\nStopping order polling system...")
        print("Goodbye!")


if __name__ == "__main__":
    main()
