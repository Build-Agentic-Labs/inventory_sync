-- Add pdf_path column to orders table
-- Run this in your Supabase SQL Editor to enable PDF path tracking

ALTER TABLE orders
ADD COLUMN IF NOT EXISTS pdf_path TEXT;

-- Add a comment to document the column
COMMENT ON COLUMN orders.pdf_path IS 'File path to the generated PDF order form';









