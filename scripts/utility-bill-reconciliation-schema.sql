-- ================================================================
-- UtilitySense Utility Bill Reconciliation Module Database Schema
-- Execute in Supabase SQL Editor
-- ================================================================

-- 1) Main Utility Bills Table
CREATE TABLE IF NOT EXISTS utility_bills (
    id TEXT PRIMARY KEY,
    location_id TEXT NOT NULL,
    plant_id TEXT NOT NULL,
    utility_id TEXT NOT NULL DEFAULT 'electricity',
    bill_month TEXT NOT NULL, -- Format: 'YYYY-MM' (e.g., '2026-07')
    bill_date DATE,
    consumer_number TEXT,
    meter_number TEXT,
    file_path TEXT,
    file_name TEXT,
    file_data TEXT,           -- Base64 or object URL preview
    status TEXT DEFAULT 'Matched', -- 'Matched', 'Variance', 'High Variance', 'Pending Review'
    notes TEXT,
    uploaded_by TEXT,
    uploaded_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL,
    last_modified_by TEXT,
    last_modified_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL,
    CONSTRAINT uq_utility_bill_loc_plant_util_month UNIQUE (location_id, plant_id, utility_id, bill_month)
);

CREATE INDEX IF NOT EXISTS idx_utility_bills_lookup ON utility_bills(location_id, plant_id, utility_id, bill_month);

-- 2) Bill Consumption Breakdown Table
CREATE TABLE IF NOT EXISTS bill_consumption (
    id TEXT PRIMARY KEY,
    bill_id TEXT NOT NULL REFERENCES utility_bills(id) ON DELETE CASCADE,
    parameter TEXT NOT NULL,
    value NUMERIC,
    unit TEXT,
    source TEXT DEFAULT 'Extracted',
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_bill_consumption_bill_id ON bill_consumption(bill_id);

-- 3) Bill Cost Breakdown Table
CREATE TABLE IF NOT EXISTS bill_cost_details (
    id TEXT PRIMARY KEY,
    bill_id TEXT NOT NULL REFERENCES utility_bills(id) ON DELETE CASCADE,
    charge_type TEXT NOT NULL,
    amount NUMERIC,
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_bill_cost_details_bill_id ON bill_cost_details(bill_id);

-- 4) Bill Solar Details Table
CREATE TABLE IF NOT EXISTS bill_solar_details (
    id TEXT PRIMARY KEY,
    bill_id TEXT NOT NULL REFERENCES utility_bills(id) ON DELETE CASCADE,
    parameter TEXT NOT NULL,
    value NUMERIC,
    unit TEXT,
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_bill_solar_details_bill_id ON bill_solar_details(bill_id);

-- 5) Reconciliation Comparison Results Table
CREATE TABLE IF NOT EXISTS reconciliation_results (
    id TEXT PRIMARY KEY,
    bill_id TEXT NOT NULL REFERENCES utility_bills(id) ON DELETE CASCADE,
    parameter TEXT NOT NULL,
    category TEXT, -- 'Consumption', 'Solar', 'Demand', 'Power Factor', 'Cost'
    bill_value NUMERIC,
    system_value NUMERIC,
    difference NUMERIC,
    difference_percentage NUMERIC,
    status TEXT, -- 'Matched', 'Variance', 'High Variance', 'N/A'
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_reconciliation_results_bill_id ON reconciliation_results(bill_id);

-- 6) Reconciliation Clarifications & Variance Analysis Table
CREATE TABLE IF NOT EXISTS reconciliation_clarifications (
    id TEXT PRIMARY KEY,
    bill_id TEXT NOT NULL REFERENCES utility_bills(id) ON DELETE CASCADE,
    reconciliation_id TEXT,
    parameter TEXT,
    reason TEXT NOT NULL,
    comment TEXT,
    status TEXT DEFAULT 'Open', -- 'Open', 'Under Review', 'Resolved'
    created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL,
    resolved_by TEXT,
    resolved_at TIMESTAMPTZ
);

CREATE INDEX IF NOT EXISTS idx_reconciliation_clarifications_bill_id ON reconciliation_clarifications(bill_id);
