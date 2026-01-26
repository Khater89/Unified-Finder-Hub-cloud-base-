# UFH v9 — Cloud DB Upgrade (tables-per-sheet)

## What changed
- Added `backend/` (Node.js API) to store Excel sheets into PostgreSQL (Supabase).
- Updated:
  - `oncall_webapp_v8` to load **Tech DB** + **US ZIP DB** from the API when `API_BASE` is set.
  - `flex_webapp_v8` to load **Tier 2 Flex** Tech DB from the API when `API_BASE` is set.
- Added `admin_upload.html` (simple upload page).

## 1) Deploy Backend (Render / Railway)
1. Create a new Supabase project (Postgres).
2. In `backend/.env` (or Render env vars) set:
   - `DATABASE_URL`
   - `ALLOWED_ORIGIN` (your frontend domain)
   - `PORT` (optional)

Run locally:
```
cd backend
npm i
npm start
```

## 2) Load your Excel files into DB
Option A (recommended): use the Admin page
- Open `admin_upload.html`
- Set API Base URL
- Upload:
  - `01-21-2026 ARS Technician.xlsx` with prefix `ars_technician`
  - `uszips.xlsx` with prefix `uszips`
  - `On Call Rotation ... .xlsx` with prefix `oncall_rotation`

Option B: seed locally
```
cd backend
npm run seed:local
```

## 3) Point the webapps to the API
Edit:
- `oncall_webapp_v8/config.js`
- `flex_webapp_v8/config.js`

Set:
```
window.API_BASE = "https://YOUR-BACKEND-DOMAIN";
```

Then open the apps. If the API is reachable, the status will show “cloud”; otherwise it falls back to the bundled JS databases.

## Table names produced by the seed script
ARS Technician:
- `ars_technician__01_21_2026_ars_technician`
- `ars_technician__usa__w2`
- `ars_technician__usa_tier_2_flex_tech`
- `ars_technician__canada__w2`

On Call Rotation:
- `oncall_rotation__on_call_rotation`
- `oncall_rotation__extended_travel__coverage`
- `oncall_rotation__daily_tech_availability`

US Zips:
- `uszips__sheet1`
