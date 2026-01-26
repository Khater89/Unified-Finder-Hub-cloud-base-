'use strict';

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const pg = require('pg');
const dotenv = require('dotenv');

dotenv.config();

const app = express();
app.use(express.json({ limit: '2mb' }));

const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
app.use(cors({ origin: allowedOrigin }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: Number(process.env.MAX_UPLOAD_MB || 25) * 1024 * 1024 }
});

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
  if (!columns.length) throw new Error('No columns detected.');
  // Always include an auto id so we can insert even if Excel headers duplicate
  const colDefs = columns.map(c => `"${c}" TEXT`).join(', ');
  const sql = `CREATE TABLE IF NOT EXISTS "${tableName}" (_id BIGSERIAL PRIMARY KEY, ${colDefs});`;
  await pool.query(sql);
}

async function replaceTable(tableName) {
  const sql = `DROP TABLE IF EXISTS "${tableName}"`;
  await pool.query(sql);
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
  // Use header row as column names (keeps your Excel structure)
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
    rawCols.forEach((k, i) => {
      out[columns[i]] = obj[k];
    });
    return out;
  });

  return { columns, rows };
}

app.get('/health', (req, res) => res.json({ ok: true }));

/**
 * POST /api/upload
 * Form-Data: file=<xlsx>
 * Query:
 *  - prefix: string (e.g. ars_technician, oncall_rotation, uszips)
 *  - mode: replace | append (default replace)
 *
 * Behavior: each sheet -> table: <prefix>__<sheetNameSanitized>
 */
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, error: 'No file uploaded.' });

    const prefix = safeName(req.query.prefix || req.body?.prefix || 'dataset');
    const mode = String(req.query.mode || req.body?.mode || 'replace').toLowerCase();

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const results = [];

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const { columns, rows } = sheetToRows(ws);

      const table = `${prefix}__${safeName(sheetName) || 'sheet'}`;

      if (!columns.length) {
        results.push({ sheet: sheetName, table, inserted: 0, note: 'empty sheet' });
        continue;
      }

      if (mode === 'replace') {
        await replaceTable(table);
      }

      await ensureTable(table, columns);
      const inserted = await insertRows(table, columns, rows);
      results.push({ sheet: sheetName, table, inserted, columns: columns.length });
    }

    res.json({ ok: true, file: req.file.originalname, prefix, mode, results });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

/**
 * GET /api/export/:table
 * Query:
 *  - format=aoa (array-of-arrays) | json (default)
 *  - limit (default 200000 for aoa, 500 for json)
 */
app.get('/api/export/:table', async (req, res) => {
  try {
    const table = safeName(req.params.table);
    const format = String(req.query.format || 'json').toLowerCase();
    const limitDefault = format === 'aoa' ? 200000 : 500;
    const limit = Math.max(1, Math.min(Number(req.query.limit || limitDefault), 300000));

    const r = await pool.query(`SELECT * FROM "${table}" LIMIT $1`, [limit]);

    if (format === 'aoa') {
      // Remove internal _id and return [ [col1, col2,...], ... ]
      const cols = r.fields.map(f => f.name).filter(n => n !== '_id');
      const aoa = r.rows.map(row => cols.map(c => row[c]));
      return res.json({ ok: true, table, columns: cols, rows: aoa });
    }

    res.json({ ok: true, table, rows: r.rows });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

// Convenience endpoints for your current webapps
app.get('/api/oncall/techdb', async (req, res) => {
  // expected by oncall_webapp_v8: window.TECH_DB_ROWS = [ [Tech ID, First, Last, Region, Zone, Type, City, State, Zip], ... ]
  try {
    const table = safeName(req.query.table || process.env.TECH_TABLE || 'ars_technician__01_21_2026_ars_technician');
    const limit = Math.max(1, Math.min(Number(req.query.limit || 200000), 300000));

    const r = await pool.query(`SELECT * FROM "${table}" LIMIT $1`, [limit]);
    const rows = r.rows.map(x => ([
      String(x.tech_id ?? x.techid ?? x.techid_2 ?? x.tech ?? x.col_1 ?? ''),
      String(x.first_name ?? x.firstname ?? x.col_2 ?? ''),
      String(x.last_name ?? x.lastname ?? x.col_3 ?? ''),
      String(x.region ?? x.col_4 ?? ''),
      String(x.zone ?? x.col_5 ?? ''),
      String(x.type ?? x.col_6 ?? ''),
      String(x.city ?? x.col_7 ?? ''),
      String(x.state ?? x.col_8 ?? ''),
      String(x.zip ?? x.zip_code ?? x.col_9 ?? ''),
    ]));
    res.json({ ok: true, table, rows });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

app.get('/api/oncall/uszips', async (req, res) => {
  // expected by oncall_webapp_v8: window.ZIP_DB_ROWS = [ [zip, lat, lon, cityLower, state], ... ]
  try {
    const table = safeName(req.query.table || process.env.ZIP_TABLE || 'uszips__sheet1');
    const limit = Math.max(1, Math.min(Number(req.query.limit || 200000), 300000));
    const r = await pool.query(`SELECT * FROM "${table}" LIMIT $1`, [limit]);

    const rows = r.rows.map(x => ([
      String(x.zip ?? x.zip_code ?? x.col_1 ?? ''),
      Number(x.lat ?? x.latitude ?? x.col_2 ?? null),
      Number(x.lon ?? x.longitude ?? x.col_3 ?? null),
      String(x.city ?? x.col_4 ?? '').trim().toLowerCase(),
      String(x.state ?? x.col_5 ?? '').trim().toUpperCase(),
    ]));

    res.json({ ok: true, table, rows });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

const port = Number(process.env.PORT || 8080);
app.listen(port, () => {
  console.log(`API running on :${port}`);
});
