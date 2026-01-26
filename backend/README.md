# Unified Finder Hub Backend (Cloud DB)

This backend lets you:
- Upload an Excel file
- Convert **each sheet** into a **separate PostgreSQL table**
- Export data for the existing webapps (On‑Call lookup) via JSON endpoints

## 1) Environment variables

Create `backend/.env`:

```
DATABASE_URL=postgresql://USER:PASSWORD@HOST:PORT/postgres
ALLOWED_ORIGIN=https://YOUR-FRONTEND.vercel.app
PORT=8080
# Optional:
TECH_TABLE=ars_technician__01_21_2026_ars_technician
ZIP_TABLE=uszips__sheet1
MAX_UPLOAD_MB=25
```

> For Supabase: copy the Postgres connection string from **Project Settings → Database → Connection string**.

## 2) Run locally

```
cd backend
npm i
npm start
```

Health check:
- `GET /health`

## 3) Seed the included Excel files (optional)

This project includes your Excel files in `backend/data/`.

```
cd backend
npm run seed:local
```

It will create tables like:
- `ars_technician__01_21_2026_ars_technician`
- `ars_technician__usa__w2`
- `oncall_rotation__on_call_rotation`
- `uszips__sheet1`

## 4) Upload from the web (API)

`POST /api/upload`

- Form-data: `file=<xlsx>`
- Query:
  - `prefix=ars_technician` (or any name you want)
  - `mode=replace` (default) or `append`

Example:
- `/api/upload?prefix=ars_technician&mode=replace`

## 5) Export tables

- `GET /api/export/<table>?format=aoa&limit=200000`

Convenience endpoints used by oncall webapp:
- `GET /api/oncall/techdb`  (returns rows in the same AOA format used by your current app)
- `GET /api/oncall/uszips`  (returns ZIP DB rows in AOA format)

