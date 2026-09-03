-- Run once in Supabase Dashboard → SQL Editor
-- Adds plant/location-scoped tariffs and multiply factors (MF)

-- 1) Tariffs: location + plant scope
ALTER TABLE tariffs ADD COLUMN IF NOT EXISTS location TEXT;
ALTER TABLE tariffs ADD COLUMN IF NOT EXISTS plant_code TEXT;

-- 2) Multiply factors table
CREATE TABLE IF NOT EXISTS multiply_factors (
  mf_id TEXT PRIMARY KEY,
  location TEXT NOT NULL,
  plant_code TEXT,
  factor NUMERIC NOT NULL,
  effective_date DATE NOT NULL,
  status TEXT NOT NULL DEFAULT 'Active'
);

-- 3) Seed location defaults (Pune ×40, Bhiwadi ×30)
INSERT INTO multiply_factors (mf_id, location, plant_code, factor, effective_date, status)
VALUES
  ('MF-PUNE', 'PUNE', NULL, 40, '2026-01-01', 'Active'),
  ('MF-BHIWADI', 'BHIWADI', NULL, 30, '2026-01-01', 'Active')
ON CONFLICT (mf_id) DO NOTHING;

-- 4) Copy existing global tariffs to PUNE location (keeps old rows as global fallback)
UPDATE tariffs SET location = 'PUNE' WHERE location IS NULL AND plant_code IS NULL;

-- 5) Bhiwadi plant-specific tariffs (adjust rates as needed)
INSERT INTO tariffs (tariff_id, type, rate, location, plant_code, effective_date, status)
VALUES
  ('TF-BHI-E', 'electricity', 10.893945, 'BHIWADI', '4020', '2026-05-01', 'Active'),
  ('TF-BHI-S', 'solar', 10.893945, 'BHIWADI', '4020', '2026-05-01', 'Active')
ON CONFLICT (tariff_id) DO NOTHING;

-- 6) Pune plant tariffs (one row per plant — edit rates in Master after deploy)
INSERT INTO tariffs (tariff_id, type, rate, location, plant_code, effective_date, status)
VALUES
  ('TF-P-4010-E', 'electricity', 10.893945, 'PUNE', '4010', '2026-05-01', 'Active'),
  ('TF-P-4010-S', 'solar', 10.893945, 'PUNE', '4010', '2026-05-01', 'Active'),
  ('TF-P-2020-E', 'electricity', 10.893945, 'PUNE', '2020', '2026-05-01', 'Active'),
  ('TF-P-2020-S', 'solar', 10.893945, 'PUNE', '2020', '2026-05-01', 'Active'),
  ('TF-P-1040-E', 'electricity', 10.893945, 'PUNE', '1040', '2026-05-01', 'Active'),
  ('TF-P-1040-S', 'solar', 10.893945, 'PUNE', '1040', '2026-05-01', 'Active')
ON CONFLICT (tariff_id) DO NOTHING;
