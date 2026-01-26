'use strict';

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const pg = require('pg');
const dotenv = require('dotenv');

dotenv.config();

function safeName(s) {
  return String(s ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '')
    .replace(/^(\d)/, '_$1')
    .slice(0, 60);
}

const pool = new pg.Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.PGSSL_DISABLE === '1' ? undefined : { rejectUnauthorized: false },
});

async function ensureTable(tableName, columns) {
  const colDefs = columns.map(c => `"${c}" TEXT`).join(', ');
  const sql = `CREATE TABLE IF NOT EXISTS "${tableName}" (_id BIGSERIAL PRIMARY KEY, ${colDefs});`;
  await pool.query(sql);
}

async function replaceTable(tableName) {
  await pool.query(`DROP TABLE IF EXISTS "${tableName}"`);
}

async function insertRows(tableName, columns, rows, chunkSize = 500) {
  let inserted = 0;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);
    const values = [];
    const placeholders = chunk.map((row, idx) => {
      const base = idx * columns.length;
      columns.forEach((c) => values.push(row[c] ?? null));
      const ps = columns.map((_, j) => `$${base + j + 1}`).join(', ');
      return `(${ps})`;
    }).join(', ');

    const colsSql = columns.map(c => `"${c}"`).join(', ');
    const sql = `INSERT INTO "${tableName}" (${colsSql}) VALUES ${placeholders};`;
    await pool.query(sql, values);
    inserted += chunk.length;
  }
  return inserted;
}

function sheetToRows(ws) {
  const json = XLSX.utils.sheet_to_json(ws, { defval: null });
  if (!json.length) return { columns: [], rows: [] };

  const rawCols = Object.keys(json[0]);
  const seen = new Map();
  const columns = rawCols.map((c, i) => {
    let name = safeName(c) || `col_${i + 1}`;
    const n = (seen.get(name) || 0) + 1;
    seen.set(name, n);
    if (n > 1) name = `${name}_${n}`;
    return name;
  });

  const rows = json.map(obj => {
    const out = {};
    rawCols.forEach((k, i) => { out[columns[i]] = obj[k]; });
    return out;
  });

  return { columns, rows };
}

async function importExcel(excelPath, prefix, mode = 'replace') {
  const buf = fs.readFileSync(excelPath);
  const wb = XLSX.read(buf, { type: 'buffer' });

  console.log(`\n==> Importing: ${excelPath}`);
  console.log(`    prefix=${prefix} mode=${mode} sheets=${wb.SheetNames.length}`);

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const { columns, rows } = sheetToRows(ws);
    const table = `${safeName(prefix)}__${safeName(sheetName) || 'sheet'}`;

    if (!columns.length) {
      console.log(`  - ${sheetName}: empty -> ${table}`);
      continue;
    }

    if (mode === 'replace') await replaceTable(table);
    await ensureTable(table, columns);
    const inserted = await insertRows(table, columns, rows);
    console.log(`  - ${sheetName}: ${inserted} rows -> ${table}`);
  }
}

(async () => {
  try {
    if (!process.env.DATABASE_URL) {
      throw new Error('Missing DATABASE_URL in .env');
    }

    const dataDir = path.join(__dirname, '..', 'data');
    await importExcel(path.join(dataDir, 'ars_technician.xlsx'), 'ars_technician', 'replace');
    await importExcel(path.join(dataDir, 'oncall_rotation.xlsx'), 'oncall_rotation', 'replace');
    await importExcel(path.join(dataDir, 'uszips.xlsx'), 'uszips', 'replace');

    console.log('\nDone.');
    process.exit(0);
  } catch (e) {
    console.error(e);
    process.exit(1);
  } finally {
    await pool.end().catch(() => {});
  }
})();
