import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const envPath = path.join(__dirname, '..', '.env');

// Read .env
const env = Object.fromEntries(
  fs.readFileSync(envPath, 'utf8')
    .split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#') && l.includes('='))
    .map((l) => {
      const i = l.indexOf('=');
      return [l.slice(0, i).trim(), l.slice(i + 1).trim()];
    })
);

const supabaseUrl = env.VITE_SUPABASE_URL || 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6';
const supabase = createClient(supabaseUrl, supabaseKey);

function convertExcelDate(excelDate) {
  if (typeof excelDate === 'number') {
    const d = XLSX.SSF.parse_date_code(excelDate);
    if (!d) return null;
    const mm = String(d.m).padStart(2, '0');
    const dd = String(d.d).padStart(2, '0');
    return `${d.y}-${mm}-${dd}`;
  }
  if (typeof excelDate === 'string') {
    const s = excelDate.trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      const pad = (n) => String(n).padStart(2, '0');
      return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
    }
  }
  return null;
}

export async function syncProductionFile(filePath, defaultPlant = 'NGM', defaultLocation = 'PUNE') {
  console.log(`\n--- Scanning Production File: ${filePath} ---`);
  if (!fs.existsSync(filePath)) {
    console.warn(`File does not exist: ${filePath}`);
    return { success: false, reason: 'File not found' };
  }

  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName =
      workbook.SheetNames.find((s) => s.toLowerCase().includes('daily basis') || s.toLowerCase().includes('daily')) ||
      workbook.SheetNames[0];

    const sheet = workbook.Sheets[sheetName];
    console.log(`Using Sheet: [${sheetName}]`);
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const mapByKey = new Map();
    let skippedBadDate = 0;
    let skippedNonProd = 0;

    for (const row of rawRows) {
      const rawDate = row['DATE'] || row['Date'] || row['date'] || row['Dated'] || row['DATED'];
      const dateStr = convertExcelDate(rawDate);
      if (!dateStr) {
        skippedBadDate++;
        continue;
      }

      const plant = String(row['Plant'] || row['PLANT'] || defaultPlant).trim();
      const location = plant === 'NGM' || plant === 'PGTL' ? 'PUNE' : String(row['Location'] || row['LOCATION'] || defaultLocation).trim();
      const product = String(row['Product'] || row['PRODUCT'] || row['Product Main'] || '').trim().toUpperCase();
      const line = String(row['Line'] || row['LINE'] || 'Main').trim();
      const finalFg = Number(row['FINAL FG'] || row['Final FG'] || row['FG'] || row['FG QTY'] || 0) || 0;

      if (product !== 'ODU' && product !== 'IDU') {
        skippedNonProd++;
        continue;
      }
      if (finalFg <= 0) continue;

      const key = `${dateStr}_${location}_${plant}_${product}_${line}`;

      if (mapByKey.has(key)) {
        const existing = mapByKey.get(key);
        existing.final_fg += finalFg;
      } else {
        mapByKey.set(key, {
          date: dateStr,
          location,
          plant,
          product,
          line,
          final_fg: finalFg,
          sheet_name: `${path.basename(filePath)}:${sheetName}`,
        });
      }
    }

    const recordsToUpsert = Array.from(mapByKey.values());
    console.log(`Parsed ${rawRows.length} rows -> ${recordsToUpsert.length} aggregated records (skipped ${skippedBadDate} bad-date, ${skippedNonProd} non-ODU/IDU)`);

    if (recordsToUpsert.length === 0) {
      return { success: true, count: 0 };
    }

    // Upsert in batches of 100
    const batchSize = 100;
    let syncedCount = 0;

    for (let i = 0; i < recordsToUpsert.length; i += batchSize) {
      const batch = recordsToUpsert.slice(i, i + batchSize);
      const { error } = await supabase
        .from('production_raw')
        .upsert(batch, { onConflict: 'date,location,plant,product,line' });

      if (!error) {
        syncedCount += batch.length;
      } else {
        console.error('Upsert error:', error.message);
      }
    }

    console.log(`Synced ${syncedCount}/${recordsToUpsert.length} records to Supabase`);
    return { success: true, count: syncedCount };
  } catch (err) {
    console.error(`File sync error: ${err.message}`);
    return { success: false, error: err.message };
  }
}

// Default run if called directly
if (process.argv[1] && process.argv[1].endsWith('sync_production_excel.mjs')) {
  const defaultFiles = [
    { path: 'D:/OneDrive - PG Electroplast Ltd/Desktop/NGM_MIS 2025.xlsx', plant: 'NGM' },
    { path: 'D:/OneDrive - PG Electroplast Ltd/Desktop/Backup PGTL Daily Basis data.xlsx', plant: 'PGTL' },
  ];

  for (const item of defaultFiles) {
    await syncProductionFile(item.path, item.plant, 'PUNE');
  }
}
