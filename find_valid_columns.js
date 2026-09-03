import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6';

const supabase = createClient(supabaseUrl, supabaseKey);

const allColumns = [
  "date", "plant", "location", "department", "shift", "operator_name",
  "electricity_opening", "electricity_closing", "electricity_consumption", "electricity_cost",
  "solar_generated", "solar_utilized", "solar_cost", "solar_utilization_pct",
  "diesel_used", "diesel_cost", "total_cost", "odu", "idu",
  "production_set", "production_qty", "production_unit", "cost_per_set", "sec",
  "waste_hazardous", "waste_non_hazardous", "waste_recycled", "remarks"
];

async function run() {
  console.log("Checking columns one by one...");
  const valid = [];
  const invalid = [];

  for (const col of allColumns) {
    // Try to insert a dummy object with only this column set
    const dummy = { [col]: col === "date" ? "2026-07-01" : (col === "plant" || col === "location" || col === "department" || col === "shift" || col === "operator_name" || col === "production_unit" || col === "remarks" ? "test" : 0) };
    
    // We append date and plant as they might be required (non-null constraints)
    if (col !== "date") dummy.date = "2026-07-01";
    if (col !== "plant") dummy.plant = "NGM";

    const { error } = await supabase.from('daily_entries').insert(dummy).select();
    
    if (error) {
      if (error.message.includes(`column "${col}" of relation "daily_entries" does not exist`) ||
          error.message.includes(`column "${col}" does not exist`) ||
          error.message.includes(`Could not find the '${col}' column`)) {
        invalid.push(col);
      } else {
        // Other errors (like constraint violations) mean the column exists!
        valid.push(col);
      }
    } else {
      valid.push(col);
    }
  }

  console.log("VALID COLUMNS:", valid);
  console.log("INVALID COLUMNS:", invalid);
}

run();
