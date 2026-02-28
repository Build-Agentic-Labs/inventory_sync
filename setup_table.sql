-- This script drops ALL existing tables and creates the inventory table fresh
-- Run this in your Supabase SQL Editor

-- Drop ALL existing tables in public schema
DO $$
DECLARE
    r RECORD;
BEGIN
    FOR r IN (SELECT tablename FROM pg_tables WHERE schemaname = 'public') LOOP
        EXECUTE 'DROP TABLE IF EXISTS public.' || quote_ident(r.tablename) || ' CASCADE';
    END LOOP;
END $$;

-- Create the inventory table
CREATE TABLE inventory (
    id BIGSERIAL PRIMARY KEY,
    sku TEXT UNIQUE NOT NULL,
    product_name TEXT,
    vendor TEXT,
    brand TEXT,
    price DECIMAL(10,2) DEFAULT 0,
    cost DECIMAL(10,2) DEFAULT 0,
    total_stock INTEGER DEFAULT 0,
    committed INTEGER DEFAULT 0,
    open_stock INTEGER DEFAULT 0,
    qty_on_order INTEGER DEFAULT 0,
    gross_margin DECIMAL(5,4) DEFAULT 0,
    total_retail DECIMAL(10,2) DEFAULT 0,
    total_cost DECIMAL(10,2) DEFAULT 0,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create index on SKU for faster lookups
CREATE INDEX idx_inventory_sku ON inventory(sku);
