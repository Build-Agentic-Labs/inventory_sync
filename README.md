# Inventory Sync & Order Management System

A Windows desktop application that syncs inventory from Excel files to Supabase and automatically processes and prints incoming orders.

## Features

- **Inventory Sync**: Automatically watches a folder for inventory Excel files and syncs them to Supabase
- **Order Processing**: Polls Supabase for new orders and generates professional PDF order forms
- **Order Management**: View all orders, check print status, and print/reprint orders on demand
- **System Tray Integration**: Runs quietly in the background with easy access from the system tray
- **Beautiful UI**: Modern, dark-themed interface for settings and order management

## Installation

1. Install Python 3.8 or higher
2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Running the Application

### As a Python Script
```bash
python inventory_sync.py
```

### First-Time Setup
On first run, you'll be prompted to configure:
- **Store Location**: Your store name (e.g., "Toppenish", "Yakima")
- **Watch Folder**: Folder to monitor for inventory files (default: Downloads)
- **File Name Pattern**: Text to identify inventory files (default: "Inventory by Product")
- **Supabase URL**: Your Supabase project URL
- **Supabase Key**: Your Supabase service role key

## Features Overview

### Inventory Sync
- Automatically detects new inventory Excel files in the watch folder
- Syncs data to Supabase `inventory` table
- Deletes processed files after successful sync
- Polls every 15 seconds for new files

### Order Processing
- Polls Supabase `orders` table for unprinted orders every 15 seconds
- Generates professional PDF order forms with:
  - Clear order number and customer information
  - Detailed fulfillment instructions for each item
  - Pickup location, delivery address, or shipping details highlighted
  - Order totals and payment status
- Automatically marks orders as printed in the database
- Saves PDFs to `order_pdfs` directory (in AppData when running as exe)

### Order Management UI
Access from system tray: **Right-click → View Orders**

- View all orders in a sortable list
- See print status at a glance
- Print or re-print any order
- View saved PDFs
- Refresh to see latest orders

### Order PDF Format
Each PDF includes:
- **Header**: Large, bold order number with date and payment status
- **Customer Info**: Name, email, phone, and shipping address (if applicable)
- **Item Details**: Quantity, product name, price, and subtotal
- **Fulfillment Instructions**: Large, highlighted action items for workers
  - `PICKUP at [Location]` - Shows which store location
  - `DELIVERY to [Address]` - Shows delivery address
  - `SHIPPING to [Address]` - Shows shipping address and cost
- **Order Totals**: Subtotal, shipping, tax, and total

## System Tray Menu

Right-click the tray icon to access:
- **Sync Now**: Manually trigger inventory sync
- **View Orders**: Open the orders management window
- **Settings**: Update configuration
- **Exit**: Close the application

## Database Schema

### Orders Table
The application expects an `orders` table with the following structure:
```sql
{
  "id": "uuid",
  "order_number": "text",
  "created_at": "timestamp",
  "customer_first_name": "text",
  "customer_last_name": "text",
  "customer_email": "text",
  "customer_phone": "text",
  "customer_shipping_address": "jsonb",
  "items": "jsonb[]",
  "subtotal": "numeric",
  "shipping_cost": "numeric",
  "tax_amount": "numeric",
  "total": "numeric",
  "payment_status": "text",
  "printed": "boolean",
  "printed_at": "timestamp",
  "pdf_path": "text" (optional but recommended)
}
```

**Note:** The `pdf_path` column is optional. If it doesn't exist, the app will still work but won't track where PDFs are saved. To add it to your existing table, run the SQL script in `add_pdf_path_column.sql` in your Supabase SQL Editor.

### Item Structure
Each item in the `items` array should have:
```json
{
  "name": "Product Name",
  "quantity": 2,
  "price": 29.99,
  "shippingCost": 5.00,
  "fulfillment": {
    "method": "pickup|delivery|shipping",
    "location": 1,  // for pickup: 1=Yakima, 2=Toppenish
    "address": {    // for delivery/shipping
      "street": "123 Main St",
      "city": "City",
      "state": "WA",
      "zipCode": "98901"
    }
  }
}
```

## Future Enhancements

- [ ] Direct printer integration (currently generates PDFs only)
- [ ] Email notifications for new orders
- [ ] Order fulfillment tracking
- [ ] Multi-store management
- [ ] Custom PDF templates

## Troubleshooting

### Orders not printing automatically
- Check Supabase credentials in Settings
- Verify `orders` table exists and has the correct schema
- Check console output for errors
- Ensure PDFs are being saved to the `order_pdfs` directory

### Inventory not syncing
- Verify watch folder path is correct
- Check file pattern matches your Excel files
- Ensure Excel files are not open/locked
- Check Supabase credentials

### Application not starting
- Verify all dependencies are installed: `pip install -r requirements.txt`
- Check that Python 3.8+ is installed
- Look for error messages in the console

## Configuration File Location

When running as a script:
- `config.json` is stored in the same directory as the script

When running as a compiled exe:
- `%LOCALAPPDATA%\InventorySync\config.json`
- PDFs saved to `%LOCALAPPDATA%\InventorySync\order_pdfs`

## Support

For issues or questions, check the console output for detailed error messages.
