"""
Quick script to reset orders to 'not printed' so they can be regenerated with new PDF layout
"""
import json
from supabase import create_client

# Load config
with open('config.json', 'r') as f:
    config = json.load(f)

# Initialize Supabase
supabase = create_client(config['supabase_url'], config['supabase_key'])

# Reset all orders to not printed
result = supabase.table('orders').update({
    'printed': False,
    'printed_at': None,
    'pdf_path': None
}).neq('id', '00000000-0000-0000-0000-000000000000').execute()  # Update all rows

print(f"Reset {len(result.data) if result.data else 0} orders to 'not printed' status")
print("Orders will be regenerated with new PDF layout on next poll")
