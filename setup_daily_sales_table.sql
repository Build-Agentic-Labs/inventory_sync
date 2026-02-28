-- Create daily_sales table for storing daily sales totals
-- Run this in Supabase SQL Editor

CREATE TABLE IF NOT EXISTS daily_sales (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  store_name TEXT NOT NULL,
  report_date DATE NOT NULL,
  total_transactions INTEGER,
  total_qty_sold INTEGER,
  total_sales DECIMAL(10,2),
  total_cogs DECIMAL(10,2),
  total_gross_profit DECIMAL(10,2),
  avg_gross_margin DECIMAL(5,2),
  total_discounts DECIMAL(10,2),
  total_tax DECIMAL(10,2),
  total_receipts DECIMAL(10,2),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE(store_name, report_date)
);

-- Create index for faster queries by date
CREATE INDEX IF NOT EXISTS idx_daily_sales_date ON daily_sales(report_date DESC);

-- Create index for faster queries by store
CREATE INDEX IF NOT EXISTS idx_daily_sales_store ON daily_sales(store_name);
