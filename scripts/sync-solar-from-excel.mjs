import XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const envPath = path.join(__dirname, "..", ".env");
const env = Object.fromEntries(
  fs
    .readFileSync(envPath, "utf8")
    .split(/\r?\n/)
    .filter((l) => l && !l.startsWith("#") && l.includes("="))
    .map((l) => {
      const i = l.indexOf("=");
      return [l.slice(0, i).trim(), l.slice(i + 1).trim()];
    })
);

const excelPath =
  "D:/OneDrive - PG Electroplast Ltd/Desktop/MSEB SOLAR ELECTRICITY CONSUMPTION NEW.xlsx";
const wb = XLSX.readFile(excelPath);
const sb = createClient(env.VITE_SUPABASE_URL, env.VITE_SUPABASE_ANON_KEY);

function excelDateToISO(v) {
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    const mm = String(d.m).padStart(2, "0");
    const dd = String(d.d).padStart(2, "0");
    return `${d.y}-${mm}-${dd}`;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return null;
}

function parseMonthSheet(sheetName, plantCode) {
  const sheet = wb.Sheets[sheetName];
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  let headerIdx = -1;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const c0 = String(rows[i][0] ?? "").trim().toLowerCase();
    if (c0 === "date") {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) return [];

  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const date = excelDateToISO(r[0]);
    if (!date) continue;
    const solar = Number(r[4]);
    const dieselRaw = Number(r[6]);
    const diesel_l = Number.isFinite(dieselRaw) ? dieselRaw : 0;
    if (!Number.isFinite(solar)) continue;
    out.push({ date, plant: plantCode, solar, diesel_l });
  }
  return out;
}

const excelRows = [
  ...parseMonthSheet("ELECTRICITY KWH May 26", "2020"),
  ...parseMonthSheet("ELECTRICITY KWH JUN 26", "2020"),
];

const { data: tariffs } = await sb.from("tariffs").select("type,rate,status");
const electRate =
  Number((tariffs || []).find((t) => t.type === "electricity" && t.status === "Active")?.rate) ||
  10.893945;
const dieselRate =
  Number((tariffs || []).find((t) => t.type === "diesel" && t.status === "Active")?.rate) || 90.62;

const { data: entries, error } = await sb
  .from("daily_entries")
  .select("id,date,plant,solar_generated,solar_cost,electricity_cost,diesel_used,diesel_cost,total_cost")
  .eq("plant", "2020")
  .order("date");

if (error) {
  console.error("DB fetch error:", error.message);
  process.exit(1);
}

const byKey = new Map(excelRows.map((r) => [`${r.date}|${r.plant}`, r]));
let updated = 0;
let skipped = 0;
const report = [];

for (const e of entries || []) {
  const x = byKey.get(`${e.date}|${e.plant}`);
  if (!x) {
    skipped += 1;
    report.push({ date: e.date, status: "no_excel_row", oldSolar: e.solar_generated });
    continue;
  }

  const solarUnits = Math.max(0, Number(x.solar) || 0);
  const dieselFinal = Number(e.diesel_used) || 0;
  const solarCost = solarUnits * electRate;
  const dieselCost = dieselFinal * dieselRate;
  const electCost = Number(e.electricity_cost) || 0;
  const totalCost = electCost + solarCost + dieselCost;

  const { error: upErr } = await sb
    .from("daily_entries")
    .update({
      solar_generated: solarUnits,
      solar_utilized: solarUnits,
      solar_cost: solarCost,
      diesel_cost: dieselCost,
      total_cost: totalCost,
    })
    .eq("id", e.id);

  if (upErr) {
    report.push({ date: e.date, status: "fail", error: upErr.message });
  } else {
    updated += 1;
    report.push({
      date: e.date,
      status: "ok",
      oldSolar: e.solar_generated,
      newSolar: solarUnits,
      newSolarCost: Math.round(solarCost * 100) / 100,
      diesel: dieselFinal,
    });
  }
}

console.log(JSON.stringify({ electRate, dieselRate, excelRows: excelRows.length, dbPgtl: (entries || []).length, updated, skipped, report }, null, 2));
