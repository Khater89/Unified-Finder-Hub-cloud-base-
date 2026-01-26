
function escapeHtml(s){
  return String(s ?? '').replace(/[&<>"']/g, (c)=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function normState(s){ return String(s||'').trim().toUpperCase(); }

// Market-to-State mapping (lightweight, no ZIP database needed for this step)
// We use it ONLY to restrict candidate markets before choosing the nearest one.
const MARKET_STATE_GROUPS = [
  { re: /Seattle/i, states: ['WA'] },
  { re: /San\s*Francisco/i, states: ['CA'] },
  { re: /Los\s*Angeles/i, states: ['CA'] },
  { re: /SE\s*Cali|San\s*Diego/i, states: ['CA'] },
  { re: /Phoenix/i, states: ['AZ'] },
  { re: /Las\s*Vegas/i, states: ['NV','UT'] },
  { re: /Kansas\s*City/i, states: ['MO','KS'] },
  { re: /Denver/i, states: ['CO','UT'] },
  { re: /Dallas/i, states: ['TX'] },
  { re: /Houston/i, states: ['TX'] },
  { re: /South\s*Florida/i, states: ['FL'] },
  { re: /Central\s*Florida/i, states: ['FL'] },
  { re: /Chicago/i, states: ['IL'] },
  { re: /Detroit/i, states: ['MI'] },
  { re: /Minnesota/i, states: [ 'MN','ND','SD' ] },
  { re: /Northern\s*Ohio/i, states: ['OH'] },
  { re: /St\.\s*Louis|St\s*Louis/i, states: ['MO'] },
  { re: /New\s*York\s*\/\s*New\s*Jersey|New\s*York|New\s*Jersey/i, states: ['NY','NJ'] },
  { re: /Connecticut/i, states: ['CT'] },
  { re: /Boston/i, states: ['MA'] },
  { re: /Philadelphia/i, states: ['PA'] },
  { re: /Nashville/i, states: ['TN'] },
  { re: /Atlanta/i, states: ['GA'] },
  { re: /Charlotte/i, states: ['NC'] },
  { re: /Raleigh|Wilmington|Fayetteville/i, states: ['NC'] },
  { re: /Richmond/i, states: ['VA'] },
  { re: /Washington\s*DC|Washington\s*D\.C\./i, states: ['DC','MD','VA'] },
];




function kmToMarketHours(distKm){
  // Convert straight-line distance to a more realistic "market distance"
  // by applying an uplift (roads are not straight), then convert to hours using an average speed.
  const uplift = 1.25;     // +25% to approximate driving vs straight-line
  const avgKmh = 95;       // average effective driving speed (highway + local)
  return (distKm * uplift) / avgKmh;
}

function kmToMiles(km){
  return km * 0.621371;
}

function formatKmAndMiles(km, decimals=0){
  if(km==null || !Number.isFinite(km)) return 'N/A';
  const kmVal = Number(km);
  const miVal = kmToMiles(kmVal);
  const kmTxt = (decimals===0 ? String(Math.round(kmVal)) : kmVal.toFixed(decimals));
  const miTxt = (decimals===0 ? String(Math.round(miVal)) : miVal.toFixed(decimals));
  return `${kmTxt} km (${miTxt} mi)`;
}


function formatHoursHM(hours){
  if(hours==null || !Number.isFinite(hours)) return '';
  const totalM = Math.round(hours*60);
  const h = Math.floor(totalM/60);
  const m = totalM % 60;
  return `${h}h ${String(m).padStart(2,'0')}m`;
}


function confidenceNoticeHtml(conf, distKm, deltaKm){
  const c = String(conf||'').toLowerCase();
  if(!c) return '';

  const d = (distKm!=null && Number.isFinite(distKm)) ? distKm : null;
  const delta = (deltaKm!=null && Number.isFinite(deltaKm)) ? deltaKm : null;

  // Convert to "market ETA" (approx driving time) instead of straight-line distance meaning
  const etaH = (d!=null) ? kmToMarketHours(d) : null;

  // Coverage policy (time-based):
  // - <= 3:00 hours => Supported
  // - 3:01 to 3:30 => Check (borderline)
  // - > 3:30 => Coverage? (possible unsupported)
  let coverage = 'Supported';
  let coverageMsg = 'ETA is within coverage. You can usually proceed.';

  if(etaH != null){
    if(etaH > 3.5){
      coverage = 'Possible unsupported';
      coverageMsg = 'ETA exceeds 3h 30m. This area may be outside on-call coverage. Verify ticket details; if correct, escalate/manual selection.';
    }else if(etaH > 3.0){
      coverage = 'Verify';
      coverageMsg = 'ETA is between 3h and 3h 30m (borderline). Please verify City/State before confirming.';
    }else{
      coverage = 'Supported';
      coverageMsg = 'ETA is within 3 hours coverage. The result can often be adopted.';
    }
  }

  let msg = '';
  if(c === 'high'){
    msg = 'High: clear congruence (the closest area is sufficiently far away from the alternatives).';
  }else if(c === 'medium'){
    msg = 'Medium: good match, but a close alternative exists.';
  }else{
    msg = 'Low: multiple markets are very close—likely a boundary/ambiguous area.';
  }

  const extra = (delta!=null) ? ` (Next variant diff: ${delta.toFixed(0)} km)` : '';
  const etaText = (etaH!=null) ? ` ETA: ${formatHoursHM(etaH)}` : '';
  const badge = coverage === 'Supported' ? 'conf-ok' : (coverage === 'Verify' ? 'conf-warn' : 'conf-bad');
  const covTitle = coverage === 'Supported' ? 'Supported' : (coverage === 'Verify' ? 'Check' : 'Coverage?');
  return `
    <div class="conf-note ${badge}">
      <div><b>Confidence Index:</b> ${escapeHtml(msg)}${escapeHtml(extra)}${escapeHtml(etaText)}</div>
      <div style="margin-top:6px;"><b>User explanation:</b> ${escapeHtml(coverageMsg)} <span class="pill">${escapeHtml(covTitle)}</span></div>
    </div>`;
}



function marketAcceptsStateForMarket(market, st){
  const state = normState(st);
  if(!state) return true;
  if(market && market.stateHint){
    return String(market.stateHint).toUpperCase() === state;
  }
  return marketAcceptsState(market?.displayName, state);
}

function marketAcceptsState(displayName, st){
  const state = normState(st);
  if(!state) return true;
  const name = String(displayName||'');
  for(const g of MARKET_STATE_GROUPS){
    if(g.re.test(name)){
      return g.states.includes(state);
    }
  }
  // If unknown market name, don't block it.
  return true;
}
// On‑Call Lookup (v4)
// Replace ZIP_RANGES approach with local ZIP database (zip_code_database.csv) bundled with the project.
// Input: ZIP OR City+State. We resolve to a ZIP, then select nearest Market ZIP by geographic distance.
// End date rule: if chosen date equals an End Date AND it's Friday -> time required (AM stays, PM moves to next week).

const $ = (id) => document.getElementById(id);

function apiBase(){
  // Set window.API_BASE in config.js (recommended). Example: https://your-backend.onrender.com
  // If empty, the app falls back to the bundled JS databases.
  return String(window.API_BASE || '').replace(/\/+$/,'');
}
async function fetchJson(url){
  const r = await fetch(url, { headers: { 'Accept': 'application/json' } });
  if(!r.ok) throw new Error(`API error ${r.status}: ${await r.text()}`);
  return await r.json();
}

let oncallSheetAOA = null;
let cleanedOncallBlob = null;
let oncallMeta = null;
let dailyAvailByDate = new Map(); // dateStr -> [{name,state}]
let techMap = new Map();

// Built-in Tech DB (bundled in techdb.js). This lets the app work immediately after extract (file://).
function loadBuiltInTechDb(){
  if(!window.TECH_DB_ROWS) return {count:0, ok:false};
  techMap.clear();
  for(const row of window.TECH_DB_ROWS){
    const techId = String(row[0]||'').trim();
    if(!techId) continue;
    techMap.set(techId, {
      techId,
      firstName: String(row[1]||'').trim(),
      lastName: String(row[2]||'').trim(),
      region: String(row[3]||'').trim(),
      zone: String(row[4]||'').trim(),
      type: String(row[5]||'').trim(),
      city: String(row[6]||'').trim(),
      state: String(row[7]||'').trim(),
      zip: normalizeZip(row[8]||'')
    });
  }
  return {count: techMap.size, ok:true};
}

let zipDB = new Map(); // zip -> {lat, lon, city, state}

function getSelectedAmPm(){
  const el = document.querySelector('input[name="ampm"]:checked');
  return el ? el.value : '';
}
function clearAmPm(){
  document.querySelectorAll('input[name="ampm"]').forEach(x => x.checked = false);
}
function setAmPmVisible(visible){
  const wrap = document.getElementById('ampmWrap');
  if(!wrap) return;
  wrap.style.display = visible ? '' : 'none';
  if(!visible) clearAmPm();
}
 // zip -> {lat, lon, city, state}

function setStatus(msg, isError=false){
  const el = $('status');
  el.textContent = msg;
  el.style.color = isError ? '#ffb4b4' : '#cbd5e1';
}
function setProg(id, pct){ $(id).style.width = Math.max(0, Math.min(100, pct)) + '%'; }
function normalizeZip(z){
  if(z == null || z === '') return null;
  const s = String(z).trim();

  // Prefer any explicit 5-digit match anywhere in the string
  const m = s.match(/\d{5}/);
  if(m) return m[0];

  // If Excel dropped leading zeros (e.g., 7803 => 07803), pad digits to 5
  const digits = s.replace(/\D/g,'');
  if(!digits) return null;
  if(digits.length >= 5) return digits.slice(0,5);
  return digits.padStart(5,'0');
}

function parseExcelDate(value){
  if(value == null || value === '') return null;

  let d = null;

  // Excel serial number
  if(typeof value === 'number' && Number.isFinite(value)){
    let serial = value;
    if(serial >= 60) serial -= 1; // Excel 1900 leap bug
    const epoch = Date.UTC(1899, 11, 31, 12, 0, 0);
    d = new Date(epoch + serial * 86400000);
  }else if(value instanceof Date && !isNaN(value.getTime())){
    // Use UTC fields, anchor at noon
    d = new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate(), 12, 0, 0));
  }else if(typeof value === 'string'){
    const s = value.trim();

    // Explicit M/D/YY or M/D/YYYY
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if(m){
      const mm = parseInt(m[1],10);
      const dd = parseInt(m[2],10);
      let yy = parseInt(m[3],10);
      if(yy < 100) yy = 2000 + yy;
      if(mm>=1 && mm<=12 && dd>=1 && dd<=31){
        d = new Date(Date.UTC(yy, mm-1, dd, 12, 0, 0));
      }
    }

    if(!d){
      const dd2 = new Date(s);
      if(!isNaN(dd2.getTime())){
        d = new Date(Date.UTC(dd2.getUTCFullYear(), dd2.getUTCMonth(), dd2.getUTCDate(), 12, 0, 0));
      }
    }
  }

  if(!d || isNaN(d.getTime())) return null;

  // Add +1 day (requested) to fix observed -1 shift
  return new Date(d.getTime() + 86400000);
}

function parseExcelDateRaw(value){
  // Same as parseExcelDate but WITHOUT the +1 day adjustment.
  if(value == null || value === '') return null;

  let d = null;

  if(typeof value === 'number' && Number.isFinite(value)){
    let serial = value;
    if(serial >= 60) serial -= 1; // Excel 1900 leap bug
    const epoch = Date.UTC(1899, 11, 31, 12, 0, 0);
    d = new Date(epoch + serial * 86400000);
  }else if(value instanceof Date && !isNaN(value.getTime())){
    d = new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate(), 12, 0, 0));
  }else if(typeof value === 'string'){
    const s = value.trim();
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if(m){
      const mm = parseInt(m[1],10);
      const dd = parseInt(m[2],10);
      let yy = parseInt(m[3],10);
      if(yy < 100) yy = 2000 + yy;
      if(mm>=1 && mm<=12 && dd>=1 && dd<=31){
        d = new Date(Date.UTC(yy, mm-1, dd, 12, 0, 0));
      }
    }
    if(!d){
      const dd2 = new Date(s);
      if(!isNaN(dd2.getTime())){
        d = new Date(Date.UTC(dd2.getUTCFullYear(), dd2.getUTCMonth(), dd2.getUTCDate(), 12, 0, 0));
      }
    }
  }

  if(!d || isNaN(d.getTime())) return null;
  return d;
}

function monthNameToIndex(name){
  const n = String(name||'').trim().toLowerCase();
  const map = {january:0,february:1,march:2,april:3,may:4,june:5,july:6,august:7,september:8,october:9,november:10,december:11};
  return (n in map) ? map[n] : null;
}


function parseDailyAvailabilityWorksheet(ws){
  // Robust parser for "Daily Tech Availability" (Non-Availability)
  // Sheet pattern (as used in the 2024/2025 rotation workbook):
  // - Weekly header rows contain day numbers across columns 1,3,5,...,13
  // - Data rows below contain [STATE] in those same columns and [TECH NAME] in the adjacent column.
  // - Sometimes the header row contains a real date (YYYY-MM-DD) in one of the day columns.
  //   When it does, we back-calculate the Sunday of that week.
  // - Otherwise, we infer the next week's Sunday as (previousSunday + 7 days).
  // Output: Map(dateStr 'YYYY-MM-DD' -> [{state,name}])

  const out = new Map();
  try{
    if(!ws) return out;

    // AOA keeps Date objects when workbook was read with cellDates:true
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, raw:true});
    if(!aoa || !aoa.length) return out;

    const stateCols = [1,3,5,7,9,11,13];
    const techCols  = [2,4,6,8,10,12,14];
    const offsets = {1:0, 3:1, 5:2, 7:3, 9:4, 11:5, 13:6}; // Sunday..Saturday

    const isRealDate = (v)=>{
      const d = parseExcelDateRaw(v);
      return (d instanceof Date) && !isNaN(d.getTime()) && d.getUTCFullYear() >= 2020;
    };

    const dayIndicator = (v)=>{
      if(v == null || v === '') return null;
      if(typeof v === 'number' && Number.isFinite(v)){
        const n = Math.trunc(v);
        return (n>=1 && n<=31) ? n : null;
      }
      const d = parseExcelDateRaw(v);
      if(d && d instanceof Date && !isNaN(d.getTime())){
        const n = d.getUTCDate();
        return (n>=1 && n<=31) ? n : null;
      }
      if(typeof v === 'string'){
        const m = String(v).match(/(\d{1,2})/);
        if(m){
          const n = parseInt(m[1],10);
          if(n>=1 && n<=31) return n;
        }
      }
      return null;
    };

    const parseState = (v)=>{
      if(typeof v !== 'string') return null;
      const s = normState(v);
      return /^[A-Z]{2}$/.test(s) ? s : null;
    };

    const parseTech = (v)=>{
      if(typeof v !== 'string') return null;
      const t = String(v).replace(/\s+/g,' ').trim();
      return t ? t : null;
    };

    const isHeaderRow = (row, idx)=>{
      if(!row) return false;
      if(idx === 0) return false; // top weekday row

      let dayCount = 0;
      let techNonEmpty = 0;
      let stateCodeCount = 0;

      for(const c of stateCols){
        if(dayIndicator(row[c]) != null) dayCount++;
        if(parseState(row[c])) stateCodeCount++;
      }
      for(const c of techCols){
        const v = row[c];
        if(typeof v === 'string' && v.trim()) techNonEmpty++;
      }

      return (dayCount >= 4) && (techNonEmpty <= 1) && (stateCodeCount <= 2);
    };

    const headers = [];
    for(let r=0; r<aoa.length; r++){
      if(isHeaderRow(aoa[r], r)) headers.push(r);
    }
    if(!headers.length) return out;

    let prevSunday = null; // Date at UTC noon

    for(let hi=0; hi<headers.length; hi++){
      const hr = headers[hi];
      const row = aoa[hr] || [];

      // Find an anchor date in the header (preferred)
      let anchor = null;
      let anchorOffset = 0;
      for(const c of stateCols){
        if(isRealDate(row[c])){
          anchor = parseExcelDateRaw(row[c]);
          anchorOffset = offsets[c] || 0;
          break;
        }
      }

      let sunday = null;
      if(anchor){
        sunday = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), anchor.getUTCDate(), 12,0,0));
        sunday = new Date(sunday.getTime() - anchorOffset * 86400000);
      }else if(prevSunday){
        sunday = new Date(prevSunday.getTime() + 7 * 86400000);
      }else{
        // No anchor and no previous week; skip.
        continue;
      }

      prevSunday = sunday;

      const next = (hi + 1 < headers.length) ? headers[hi+1] : aoa.length;

      for(let r = hr + 1; r < next; r++){
        const rr = aoa[r] || [];
        for(let i=0; i<stateCols.length; i++){
          const sc = stateCols[i];
          const tc = techCols[i];

          const st = parseState(rr[sc]);
          const nm = parseTech(rr[tc]);
          if(!st || !nm) continue;

          const dt = new Date(sunday.getTime() + (offsets[sc] || 0) * 86400000);
          const dateStr = ymd(dt);

          const arr = out.get(dateStr) || [];
          arr.push({state: st, name: nm});
          out.set(dateStr, arr);
        }
      }
    }

    // Deduplicate per date: STATE + NAME
    for(const [dateStr, arr] of out.entries()){
      const seen = new Set();
      const dedup = [];
      for(const it of (arr || [])){
        const st = normState(it.state);
        const nm = String(it.name||'').replace(/\s+/g,' ').trim();
        if(!st || !nm) continue;
        const key = st + '|' + nm.toLowerCase();
        if(seen.has(key)) continue;
        seen.add(key);
        dedup.push({state: st, name: nm});
      }
      out.set(dateStr, dedup);
    }

    return out;
  }catch(e){
    console.warn('Daily Tech Availability parse failed:', e);
    return new Map();
  }
}




function parseDailyAvailabilityAOA(aoa){
  // Daily Tech Availability format:
  // - A month header row like "December 2024"
  // - Dates across columns 2,4,6... and each date uses two columns: [State] then [Tech Name]
  const out = new Map(); // dateStr -> [{name,state}]
  let curYear = null;
  let curMonth = null;
  let dayCols = [];

  for(let r=0;r<aoa.length;r++){
    const row = aoa[r] || [];
    const first = row[0];

    if(typeof first === 'string'){
      const mm = first.trim().match(/^([A-Za-z]+)\s+(\d{4})/);
      if(mm){
        curMonth = monthNameToIndex(mm[1]);
        curYear = parseInt(mm[2],10);
        dayCols = [];

        for(let c=1;c<row.length;c+=2){
          const v = row[c];
          if(v===null || v===undefined || v==='') continue;

          const d0 = parseExcelDateRaw(v);
          if(!d0 || isNaN(d0.getTime())) continue;

          // Many header cells store only the day number as a 1900-based date; take UTC date part only.
          const day = d0.getUTCDate();
          if(curMonth===null || !curYear) continue;

          const real = new Date(Date.UTC(curYear, curMonth, day, 12, 0, 0));
          const dateStr = ymd(real);
          dayCols.push({dateStr, stateCol:c, nameCol:c+1});
        }
        continue;
      }
    }

    if(dayCols.length){
      for(const col of dayCols){
        const name = row[col.nameCol];
        if(name===null || name===undefined || String(name).trim()==='') continue;

        const state = row[col.stateCol];
        const rec = {
          name: String(name).trim(),
          state: (state===null||state===undefined) ? '' : String(state).trim()
        };

        const arr = out.get(col.dateStr) || [];
        arr.push(rec);
        out.set(col.dateStr, arr);
      }
    }
  }
  return out;
}

function matchTechByInitialAndLast(shortName){
  // Examples: "F. Oshinowo" , "K Gibson"
  const s = String(shortName||'').trim().replace(/\s+/g,' ');
  const m = s.match(/^([A-Za-z])\.?\s*([A-Za-z'\-]+)$/);
  if(!m) return null;
  const ini = m[1].toUpperCase();
  const last = m[2].toUpperCase();

  const candidates = [];
  for(const t of techMap.values()){
    const ln = String(t.lastName||'').toUpperCase();
    const fn = String(t.firstName||'');
    if(ln===last && fn && fn[0].toUpperCase()===ini){
      candidates.push(t);
    }
  }
  return candidates.length===1 ? candidates[0] : null;
}
function ymd(d){
  // Use UTC fields to avoid off-by-one day issues across timezones.
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2,'0');
  const da = String(d.getUTCDate()).padStart(2,'0');
  return `${y}-${m}-${da}`;
}

function todayYmd(){
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth()+1).padStart(2,'0');
  const d = String(now.getDate()).padStart(2,'0');
  return `${y}-${m}-${d}`;
}


function isUtcMinus5(){
  // getTimezoneOffset() returns minutes *behind* UTC.
  // UTC-5 => 300
  return (new Date().getTimezoneOffset() === 300);
}

function shiftYmd(ymdStr, deltaDays){
  // ymdStr must be 'YYYY-MM-DD'
  if(!ymdStr) return ymdStr;
  const parts = String(ymdStr).split('-');
  if(parts.length !== 3) return ymdStr;
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);
  if(!y || !m || !d) return ymdStr;

  // Use UTC so we don't get DST edge issues.
  const dt = new Date(Date.UTC(y, m-1, d));
  if(isNaN(dt.getTime())) return ymdStr;
  dt.setUTCDate(dt.getUTCDate() + Number(deltaDays || 0));

  const yy = dt.getUTCFullYear();
  const mm = String(dt.getUTCMonth()+1).padStart(2,'0');
  const dd = String(dt.getUTCDate()).padStart(2,'0');
  return `${yy}-${mm}-${dd}`;
}

function effectiveInputDateYmd(enteredYmd){
  return isUtcMinus5() ? shiftYmd(enteredYmd, +1) : enteredYmd;
}

function effectiveTodayYmd(){
  const t = todayYmd();
  return isUtcMinus5() ? shiftYmd(t, +1) : t;
}


function normalizeDate(d){
  if(!(d instanceof Date) || isNaN(d.getTime())) return d;
  // Anchor at UTC noon to avoid timezone shifting across date boundaries.
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 12, 0, 0));
}
function haversineKm(lat1, lon1, lat2, lon2){
  const toRad = (x) => x * Math.PI / 180;
  const R = 6371;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat/2)**2 +
            Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLon/2)**2;
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

// show file names
['oncallFile','techFile'].forEach((id)=>{
  $(id).addEventListener('change', ()=>{
    const f = $(id).files?.[0];
    $(id==='oncallFile'?'oncallName':'techName').textContent = f ? f.name : 'No file selected';
  });
});

// Load bundled ZIP DB (US only)
async function loadBundledZipDb(){
  // Preferred (cloud): load from Backend API if configured.
  const base = apiBase();
  if(base){
    try{
      const j = await fetchJson(`${base}/api/oncall/uszips`);
      if(j && j.ok && Array.isArray(j.rows) && j.rows.length){
        window.ZIP_DB_ROWS = j.rows;
      }
    }catch(e){
      // silent fallback to bundled zipdb.js
      console.warn('ZIP API load failed; using bundled zipdb.js', e);
    }
  }
  // Fallback (offline): Data is bundled in zipdb.js as window.ZIP_DB_ROWS.
  if(!window.ZIP_DB_ROWS) throw new Error('ZIP DB file zipdb.js is not loaded. Please ensure it exists in the same folder.');
  zipDB.clear();
  for(const row of window.ZIP_DB_ROWS){
    // NOTE: When ZIPs come from Supabase they might be numbers, not strings.
    // Always normalize to a 5-digit string key so lookups work reliably.
    const z = normalizeZip(row[0]);
    const lat = Number(row[1]);
    const lon = Number(row[2]);
    const city = String(row[3]||'').trim().toLowerCase();
    const state = String(row[4]||'').trim().toUpperCase();
    if(!z || !Number.isFinite(lat) || !Number.isFinite(lon)) continue;
    zipDB.set(z, {lat, lon, city, state});
  }
  return zipDB.size;
}

function findZipByCityState(city, state){
  const c = (city||'').trim().toLowerCase();
  const s = (state||'').trim().toUpperCase();
  if(!c || !s) return null;

  const matches = [];
  for(const [z, info] of zipDB.entries()){
    if(info.city === c && info.state === s) matches.push(z);
  }
  if(!matches.length) return null;
  matches.sort((a,b)=>parseInt(a,10)-parseInt(b,10));
  return matches[Math.floor(matches.length/2)];
}

// Tech DB (first sheet only; ignore non-US)
async function loadTechDb(file){
  setProg('techProg', 5);
  const buf = await file.arrayBuffer();
  setProg('techProg', 25);
  const wb = XLSX.read(buf, {type:'array'});
  setProg('techProg', 45);

  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];
  const json = XLSX.utils.sheet_to_json(ws, {defval:''});
  setProg('techProg', 70);

  techMap.clear();
  for(const r of json){
    const techId = String(r['Tech ID'] ?? '').trim();
    if(!techId) continue;
    const c = String(r['Country'] ?? '').trim().toUpperCase();
    if(c && c !== 'US') continue;
    techMap.set(techId, {
      techId,
      firstName: String(r['First Name'] ?? '').trim(),
      lastName: String(r['Last Name'] ?? '').trim(),
      region: String(r['Region'] ?? '').trim(),
      zone: String(r['Zone'] ?? '').trim(),
      type: String(r['Type'] ?? '').trim(),
      city: String(r['City'] ?? '').trim(),
      state: String(r['State'] ?? '').trim(),
      zip: normalizeZip(r['Zip']),
    });
  }
  setProg('techProg', 100);
  return {count: techMap.size, sheet: firstSheetName};
}

// OnCall: first sheet + replace market names to ZIP
const MARKET_NAME_TO_ZIP = new Map([
  ['Seattle', '98101'], ['San Francisco', '94102'], ['Los Angeles', '90001'],
  ['SE Cali (LA south/San Diego)', '92101'], ['Phoenix', '85001'], ['Las Vegas', '89101'],
  ['Kansas City', '64106'], ['Denver', '80202'], ['Dallas', '75201'], ['Houston', '77002'],
  ['South Florida', '33101'], ['Central Florida', '32801'], ['Chicago', '60601'], ['Detroit', '48226'],
  ['Minnesota', '55401'], ['Northern Ohio', '44114'], ['St. Louis', '63101'],
  ['New York/New Jersey', '10001'], ['Connecticut', '06103'], ['Boston', '02108'], ['Philadelphia', '19102'],
  ['Nashville, TN', '37219'], ['Atlanta', '30303'], ['Charlotte, NC', '28202'], ['Raleigh', '27601'],
  ['Wilmington', '28401'], ['Fayetteville', '28301'], ['Richmond, VA', '23219'], ['Washington DC', '20001'],
]);
const MARKET_ZIP_TO_NAME = (() => {
  const m = new Map();
  for (const [name, zip] of MARKET_NAME_TO_ZIP.entries()) m.set(zip, name);
  return m;
})();

function replaceMarketNamesToZips(text){
  if(typeof text !== 'string') return text;
  let out = text;
  for(const [name, zip] of MARKET_NAME_TO_ZIP.entries()){
    const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*');
    out = out.replace(new RegExp(esc, 'ig'), zip);
  }
  return out;
}
function extractMarketInfoFromRow(rowArr, centerZip, firstDateCol){
  // Goal: show the descriptive sentence beside the market (often stored in the SAME cell as the ZIP, with new lines).
  // Example: "90001\nSchedule On Call Tech for All After Hours Work"
  // We scan the entire row and pick the best "note-like" text. If a cell contains the ZIP, we still extract text after removing ZIP.
  let best = '';
  let bestScore = -1;

  const NOTE_RE = /(schedule|effective|after hours|on\s*call|rotation|reserve)/i;

  function extractNoteFromCell(cellText){
    if(typeof cellText !== 'string') return '';
    let s = cellText.replace(/\r/g,'').trim();
    if(!s) return '';

    // If the cell contains the ZIP, remove ONLY the zip digits and common separators,
    // but keep the rest (this is where your note often lives).
    if(centerZip && s.includes(centerZip)){
      s = s.replace(new RegExp(centerZip, 'g'), '').replace(/[()\/\-\s]+$/g,'').trim();
    }

    // If multi-line, prefer lines that look like notes.
    const lines = s.split('\n').map(x=>x.trim()).filter(Boolean);
    const noteLines = lines.filter(x => NOTE_RE.test(x));

    if(noteLines.length){
      return noteLines.join(' | ');
    }

    // If not note-like lines, but still has words, return as-is.
    return lines.join(' | ');
  }

  for(let i=0;i<rowArr.length;i++){
    const raw = rowArr[i];
    if(typeof raw !== 'string') continue;

    const cand = extractNoteFromCell(raw);
    if(!cand) continue;

    if(/^\d{5}$/.test(cand)) continue;

    const hasNoteWord = NOTE_RE.test(cand) ? 1500 : 0;
    const letters = (cand.match(/[A-Za-z]/g) || []).length;
    const lengthBonus = Math.min(cand.length, 260);

    const hasTechId = /\b\d{4,6}\b/.test(cand);
    const looksLikeAssignment = hasTechId && !NOTE_RE.test(cand);

    let score = hasNoteWord + letters + lengthBonus;
    if(looksLikeAssignment) score -= 900;
    if(cand.length < 12) score -= 250;

    if(score > bestScore){
      bestScore = score;
      best = cand;
    }
  }

  // Fallback: scan only before date cols
  if(!best){
    const maxCol = (Number.isFinite(firstDateCol) && firstDateCol > 0) ? firstDateCol : rowArr.length;
    for(let i=0;i<Math.min(rowArr.length, maxCol);i++){
      const v=rowArr[i];
      if(typeof v !== 'string') continue;
      const s=v.trim();
      if(!s) continue;
      if(/^\d{5}$/.test(s)) continue;
      if(s.length > best.length) best=s;
    }
  }

  return best;
}
function findRowIndexContains(aoa, needle){
  const n = needle.toLowerCase();
  for(let r=0;r<aoa.length;r++){
    for(let c=0;c<aoa[r].length;c++){
      const v = aoa[r][c];
      if(typeof v === 'string' && v.toLowerCase().includes(n)) return r;
    }
  }
  return -1;
}
function parseOncallAOA(aoa){
  const startRow = findRowIndexContains(aoa, 'Start');
  const endRow = findRowIndexContains(aoa, 'End');
  if(startRow < 0 || endRow < 0) throw new Error('Could not find Start/End rows.');

  let firstDateCol = -1;
  for(let c=0;c<aoa[startRow].length;c++){
    if(parseExcelDate(aoa[startRow][c])){ firstDateCol=c; break; }
  }
  if(firstDateCol < 0) throw new Error('Could not find date columns.');

  const weeks=[];
  for(let c=firstDateCol;c<aoa[startRow].length;c++){
    const s=parseExcelDate(aoa[startRow][c]);
    const e=parseExcelDate(aoa[endRow][c]);
    if(s && e) weeks.push({col:c,start:normalizeDate(s),end:normalizeDate(e)});
  }
  if(!weeks.length) throw new Error('No weeks found.');
  // --- Robust end-date derivation ---
  // Excel End row sometimes shifts by 1 day. To make End dates stable:
  // For each week i, set end = (nextStart - 1 day) when nextStart exists.
  // This matches real "weekly column" boundaries and prevents the AM/PM prompt from appearing a day early.
  for(let i=0;i<weeks.length-1;i++){
    const ns = weeks[i+1].start;
    const nsUtc = Date.UTC(ns.getUTCFullYear(), ns.getUTCMonth(), ns.getUTCDate(), 12, 0, 0);
    weeks[i].end = new Date(nsUtc - 86400000); // day before next start (UTC-safe)
  }
  // For the last week: if End is missing/invalid, default to start + 6 days.
  const last = weeks[weeks.length-1];
  if(!(last.end instanceof Date) || isNaN(last.end.getTime())){
    last.end = new Date(last.start.getTime() + 6*86400000);
  }else{
    // normalize last end to date-only
    last.end = new Date(last.end.getFullYear(), last.end.getMonth(), last.end.getDate());
  }
  // --- End robust derivation ---
  const markets=[];
  let lastMarket = null;

  for(let r=0;r<aoa.length;r++){
    const a=normalizeZip(aoa[r][0]);
    const b=normalizeZip(aoa[r][1]);
    const centerZip=a||b;
    const sample=aoa[r][weeks[0].col];
    const hasSample = !(sample===''||sample==null);

    if(centerZip){
      if(!hasSample) { lastMarket = null; continue; }

      const displayName = MARKET_ZIP_TO_NAME.get(centerZip) || `Market ${centerZip}`;
      const info = extractMarketInfoFromRow(aoa[r], centerZip, firstDateCol);

      // Special case: merged market row "New York/New Jersey" has TWO tech rows:
      // top row = NY, next row = NJ (the market name cell is usually merged so the NJ row has empty market cell).
      if(/New\s*York\s*\/\s*New\s*Jersey/i.test(displayName)){
        markets.push({row:r, centerZip, displayName, info, stateHint:'NY', subIndex:0});
        lastMarket = {row:r, centerZip, displayName, info, subIndex:0};
      }else{
        markets.push({row:r, centerZip, displayName, info});
        lastMarket = {row:r, centerZip, displayName, info};
      }

      continue;
    }

    // Handle the merged second row for New York/New Jersey (NJ row)
    if(lastMarket && (r === lastMarket.row + 1) && hasSample){
      if(/New\s*York\s*\/\s*New\s*Jersey/i.test(lastMarket.displayName)){
        markets.push({row:r, centerZip:lastMarket.centerZip, displayName:lastMarket.displayName, info:lastMarket.info, stateHint:'NJ', subIndex:1});
      }
    }
  }
  if(!markets.length) throw new Error('No markets found.');
  return {weeks, markets};
}
async function loadOncall(file){
  setProg('oncallProg', 5);
  const buf = await file.arrayBuffer();
  setProg('oncallProg', 25);
  const wb = XLSX.read(buf, {type:'array', cellDates:true});
  // Read Daily Tech Availability (optional)
  try{
    const dailyName =
      wb.SheetNames.find(n=>String(n||'').trim().toLowerCase()==='daily tech availability') ||
      wb.SheetNames.find(n=>String(n||'').toLowerCase().includes('daily tech'));
    const dailyWs = dailyName ? wb.Sheets[dailyName] : null;
    if(dailyWs){
      dailyAvailByDate = parseDailyAvailabilityWorksheet(dailyWs);
    }else{
      dailyAvailByDate = new Map();
    }
  }catch(_){
    dailyAvailByDate = new Map();
  }
  setProg('oncallProg', 45);

  const ws = wb.Sheets[wb.SheetNames[0]];
  let aoa = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  setProg('oncallProg', 60);

  aoa = aoa.map(row => row.map(cell => replaceMarketNamesToZips(cell)));
  oncallSheetAOA = aoa;

  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, XLSX.utils.aoa_to_sheet(oncallSheetAOA), 'On Call Rotation (Cleaned)');
  const wbout = XLSX.write(newWb, {bookType:'xlsx', type:'array'});
  cleanedOncallBlob = new Blob([wbout], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  $('btnDownload').disabled = false;

  oncallMeta = parseOncallAOA(oncallSheetAOA);
  setProg('oncallProg', 100);
  return {meta:oncallMeta, sheet: wb.SheetNames[0]};
}

// Market detection by geo distance

function detectMarketForZip(inputZip, selectedState, week, dateStr){
  const z = normalizeZip(inputZip);
  if(!z) throw new Error('Enter a valid ZIP.');
  const zInfo = zipDB.get(z);
  if(!zInfo) throw new Error('ZIP not found in the local database.');

  if(!week || week.col==null) throw new Error('Week column is not available.');

  // IMPORTANT CHANGE:
  // Instead of comparing ticket ZIP to "market center ZIP",
  // we compare it to the ZIP of the *On-Call tech* for each market in the selected week.
  // This makes the "distance" represent proximity to the actual selected tech (approx).

  const isToday = (dateStr === effectiveTodayYmd());

  const warnings = [];

  // 1) Filter markets by State coverage mapping (State is a hint)
  let candidateMarkets = oncallMeta.markets.filter(m => marketAcceptsStateForMarket(m, selectedState));
  let usedStateFilter = true;

  if(!candidateMarkets.length){
    // If state isn't covered by the OnCall markets list, fall back to ZIP-only matching across all markets.
    warnings.push(`State ${selectedState.toUpperCase()} is not covered in the OnCall markets list. Using ZIP-only matching across all markets.`);
    candidateMarkets = oncallMeta.markets.slice();
    usedStateFilter = false;
  }

  const candidates = [];
  let reservedSkipped = 0;
  const reservedCandidates = [];
  let missingTechInDB = 0;
  let missingTechZip = 0;
  let missingTechZipInZipDB = 0;
  let missingTechIdInOncall = 0;
  for(const market of candidateMarkets){
    const rawCell = oncallSheetAOA[market.row]?.[week.col];
    const techCell = parseTechCell(rawCell);
    if(!techCell?.techId){ missingTechIdInOncall++; continue; }

    const tech = techMap.get(techCell.techId);
    if(!tech){ missingTechInDB++; continue; }

    const techZip = normalizeZip(tech.zip || tech.Zip || tech.ZIP);
    let techZipInfo = null;
    let distKm = null;

    if(!techZip){
      missingTechZip++;
    }else{
      techZipInfo = zipDB.get(techZip);
      if(!techZipInfo){
        missingTechZipInZipDB++;
      }else{
        distKm = haversineKm(zInfo.lat, zInfo.lon, techZipInfo.lat, techZipInfo.lon);
      }
    }

    const marketInfoText = String(market?.info || market?.note || '');
    const isReserved = !!(techCell?.reserved) || /\breserv/i.test(marketInfoText);

    const cand = {
      market,
      distKm,
      techId: techCell.techId,
      techZip,
      techName: tech.firstName ? `${tech.firstName} ${tech.lastName||''}`.trim() : (tech.name||tech.fullName||null)
    };

    // Reserved handling (like before):
    // If Reserved and NOT today -> keep candidate but mark it as reserved-only pool.
    // Prefer non-reserved when possible.
    if(isReserved && !isToday){
      reservedSkipped++;
      reservedCandidates.push(cand);
      continue;
    }

    candidates.push(cand);
}

  const pool = candidates.length ? candidates : reservedCandidates;
  const allReserved = (!candidates.length && reservedCandidates.length);


  if(!pool.length){
    // Auto week fallback: sometimes the selected week column is empty (especially around boundary dates).
    // Try adjacent weeks (next, then previous) before failing.
    const weekIdx = oncallMeta.weeks.findIndex(w => w.col === week.col);
    const tryWeeks = [];
    if(weekIdx >= 0){
      if(oncallMeta.weeks[weekIdx+1]) tryWeeks.push(oncallMeta.weeks[weekIdx+1]);
      if(oncallMeta.weeks[weekIdx-1]) tryWeeks.push(oncallMeta.weeks[weekIdx-1]);
    }

    function collectForWeek(wk){
      const cand = [];
      const resCand = [];
      for(const market of candidateMarkets){
        const rawCell = oncallSheetAOA[market.row]?.[wk.col];
        const tc = parseTechCell(rawCell);
        if(!tc?.techId) continue;
        const t = techMap.get(tc.techId);
        if(!t) continue;
        const tZip = normalizeZip(t.zip || t.Zip || t.ZIP);
        if(!tZip) continue;
        const tZipInfo = zipDB.get(tZip);
        if(!tZipInfo) continue;

        const mi = String(market?.info || market?.note || '');
        const isRes = !!(tc?.reserved) || /\breserv/i.test(mi);
        const dkm = haversineKm(zInfo.lat, zInfo.lon, tZipInfo.lat, tZipInfo.lon);

        const c = { market, distKm: dkm, techId: tc.techId, techZip: tZip,
          techName: t.firstName ? `${t.firstName} ${t.lastName||''}`.trim() : (t.name||t.fullName||null) };

        if(isRes && !isToday){ resCand.push(c); continue; }
        cand.push(c);
      }
      const p = cand.length ? cand : resCand;
      return {cand, resCand, pool:p};
    }

    for(const wk of tryWeeks){
      const alt = collectForWeek(wk);
      if(alt.pool.length){
        warnings.push(`No candidates found in the selected week column. Auto-switched to nearest available week: ${wk.startStr} → ${wk.endStr}.`);
        // replace current selection
        candidates.length = 0; reservedCandidates.length = 0;
        for(const c of alt.cand) candidates.push(c);
        for(const c of alt.resCand) reservedCandidates.push(c);
        week = wk;
        break;
      }
    }

    // recompute after fallback
    const pool2 = candidates.length ? candidates : reservedCandidates;
    if(!pool2.length){
      // Second-pass fallback: if we used a State filter but got 0 candidates, retry ZIP-only across ALL markets.
      if(usedStateFilter){
        warnings.push(`No candidates found after applying State filter (${selectedState.toUpperCase()}). Retrying ZIP-only across all markets...`);
        candidateMarkets = oncallMeta.markets.slice();
        usedStateFilter = false;

        // Try again on current week (and adjacent weeks) using all markets
        const altNow = collectForWeek(week);
        if(altNow.pool.length){
          candidates.length = 0; reservedCandidates.length = 0;
          for(const c of altNow.cand) candidates.push(c);
          for(const c of altNow.resCand) reservedCandidates.push(c);
        }else{
          for(const wk of tryWeeks){
            const alt = collectForWeek(wk);
            if(alt.pool.length){
              warnings.push(`Auto-switched to nearest available week: ${wk.startStr} → ${wk.endStr}.`);
              candidates.length = 0; reservedCandidates.length = 0;
              for(const c of alt.cand) candidates.push(c);
              for(const c of alt.resCand) reservedCandidates.push(c);
              week = wk;
              break;
            }
          }
        }

        const pool3 = candidates.length ? candidates : reservedCandidates;
        if(!pool3.length){
          throw new Error('No On-Call tech candidates found for the selected week/state.');
        }
      } else {
        throw new Error('No On-Call tech candidates found for the selected week/state.');
      }
    }
    // continue using pool2 below
  }

  const poolFinal = candidates.length ? candidates : reservedCandidates;
  const allReservedFinal = (!candidates.length && reservedCandidates.length);


  if(allReservedFinal){
    warnings.push('All candidate markets are Reserved for the selected date. Showing the closest reserved market (tech details will be blocked unless date is today).');
  }

  poolFinal.sort((a,b)=>( (a.distKm ?? Infinity) - (b.distKm ?? Infinity) ));
  const best = poolFinal[0];
  const second = poolFinal.length>1 ? poolFinal[1] : null;

  const deltaKm = (second && Number.isFinite(best.distKm) && Number.isFinite(second.distKm)) ? (second.distKm - best.distKm) : null;

  // Confidence based on separation between best and second best.
  let confidence = 'Low';
  if(deltaKm == null){
    confidence = 'High';
  }else if(deltaKm >= 150){
    confidence = 'High';
  }else if(deltaKm >= 75){
    confidence = 'Medium';
  }else{
    confidence = 'Low';
  }

  // Cap confidence so High only happens when coverage is Supported (ETA <= 3h)
  const etaH = kmToMarketHours(best.distKm);
  if(etaH > 3.5){
    confidence = 'Low';
  }else if(etaH > 3.0){
    if(String(confidence).toLowerCase() === 'high') confidence = 'Medium';
  }


  // Build warnings about skipped markets (helps when weekly OnCall sheet / Tech DB changes)
  if(reservedSkipped>0) warnings.push(`Reserved markets excluded for selected date: ${reservedSkipped}`);
  if(missingTechIdInOncall>0) warnings.push(`OnCall sheet has empty/unknown tech entries (skipped): ${missingTechIdInOncall}`);
  if(missingTechInDB>0) warnings.push(`Tech ID not found in Tech DB (skipped): ${missingTechInDB} — update Tech DB if OnCall changed`);
  if(missingTechZip>0) warnings.push(`Tech ZIP missing in Tech DB (skipped): ${missingTechZip}`);
  if(missingTechZipInZipDB>0) warnings.push(`Tech ZIP not found in ZIP DB (skipped): ${missingTechZipInZipDB}`);
  return {
    zip: z,
    market: best.market,
    distKm: best.distKm,
    zipInfo: zInfo,
    confidence,
    deltaKm,
    candidatesCount: poolFinal.length,
    allReserved: !!allReservedFinal,
    secondMarket: second ? second.market : null,
    secondDistKm: second ? second.distKm : null,
    alternatives: [ {market: best.market, distKm: best.distKm, techId: best.techId, techZip: best.techZip, techName: best.techName}, ...(second ? [{market: second.market, distKm: second.distKm, techId: second.techId, techZip: second.techZip, techName: second.techName}] : []) ],
    techZip: best.techZip,
    techIdHint: best.techId,
    techNameHint: best.techName,
    warnings
  };
}


// Date logic (End Date time only if Friday)
function pickWeekIndexForDate(dateStr, ampm){
  if(!dateStr) throw new Error('Choose a date.');

  // Build ordered list of start dates (YYYY-MM-DD)
  const starts = oncallMeta.weeks.map(w => ymd(w.start));

  // If the selected date is a shared boundary (it equals the Start of a week AFTER the first),
  // then prompt AM/PM:
  // AM  => previous week (same as "stay on the previous column")
  // PM  => current week (next column, whose Start equals that date)
  const boundaryIndex = starts.indexOf(dateStr);
  if(boundaryIndex > 0){
    if(!ampm) throw new Error('This date is shared between two weeks. Please choose ticket time: Morning or Evening.');
    return (ampm === 'AM') ? (boundaryIndex - 1) : boundaryIndex;
  }

  // Normal case: choose the week where start_i <= date < start_{i+1}
  for(let i=starts.length-1;i>=0;i--){
    if(dateStr >= starts[i]){
      // If it's the last week or before next start, it's in this week.
      if(i === starts.length-1) return i;
      if(dateStr < starts[i+1]) return i;
    }
  }

  throw new Error('Date is outside Start/End ranges.');
}

function parseTechCell(cellValue){
  const s=String(cellValue??'').trim();
  if(!s) return null;
  const reserved=/reserve/i.test(s);
  const idMatch=s.match(/\b(\d{4,6})\b/);
  return {raw:s, techId:idMatch?idMatch[1]:null, reserved};
}

function renderResult({inputZip,dateStr,enteredDateStr,ampm,marketDetected,week,techCell,tech,selectedState='',reservedBlocked=false, techNotFoundReason}){
  // `requireChoice` must be available before the HTML template is built.
  // It was previously declared later with `const`, which triggered a TDZ error:
  // "Cannot access 'requireChoice' before initialization".
  const cLower = String(marketDetected?.confidence || '').toLowerCase();
  const requireChoice = (!marketDetected?.userSelected) && !!marketDetected?.secondMarket && (cLower === 'low' || cLower === 'medium');

  const reservedBadge = techCell?.reserved ? '<span class="badge">RESERVED</span>' : '';
  const techFoundBadge = reservedBlocked ? '' : (tech ? '' : '<span class="badge">NOT FOUND IN TECH DB</span>');
  const reservedBlockHtml = reservedBlocked ? `<div class="warn-card">
          <div class="warn-title">🚫 OnCall Reserved</div>
          <div class="warn-body">This On-Call is Reserved. Please choose another tech.</div>
        </div>` : '';
  // Non-Availability: for this section only, use (entered date - 1 day).
  // We intentionally do NOT apply the UTC-5 +1 adjustment here.
  const __enteredForNA = (enteredDateStr || dateStr);
  const nonAvailDateStr = __enteredForNA ? shiftYmd(__enteredForNA, -1) : __enteredForNA;
  const rawNonAvail = (dailyAvailByDate && dailyAvailByDate.get(nonAvailDateStr)) ? dailyAvailByDate.get(nonAvailDateStr) : [];
  const stFilter = normState(selectedState);
  const nonAvailList = stFilter ? rawNonAvail.filter(x => normState(x.state) === stFilter) : [];


  const nonAvailHtml = nonAvailList.length
    ? `<div class="na-card">
         <div class="na-title">Non Available Tech <span class="na-count">(${nonAvailList.length})</span></div>
         <div class="na-sub">Date: <b>${escapeHtml(nonAvailDateStr)}</b> — State: <b>${escapeHtml(stFilter)}</b></div>
         <div class="na-list">
           ${nonAvailList.map(x=>{
              const st = x.state ? ` <span class="na-state">${escapeHtml(x.state)}</span>` : '';
              return `<div class="na-item">${escapeHtml(x.name)}${st}</div>`;
            }).join('')}
         </div>
       </div>`
    : `<div class="na-card">
         <div class="na-title">Non Available Tech</div>
         ${stFilter
           ? `<div class="na-empty">No non-available tech for <b>${escapeHtml(stFilter)}</b> on <b>${escapeHtml(nonAvailDateStr)}</b>.</div>`
           : `<div class="na-empty">Enter State (2-letter) to show Non Available Tech for the same date/state.</div>`}
       </div>`;

$('result').innerHTML = `
    <div class="row" style="justify-content:space-between">
      <div><b>Market</b>: ${marketDetected.market.displayName}${marketDetected.userSelected ? ' <span class="badge">User selected</span>' : ''} <span class="badge">${marketDetected.market.centerZip}</span>${marketDetected.market.info ? ` <span class="muted">— ${marketDetected.market.info}</span>` : ''}</div>
      <div class="muted">Distance (to Tech ZIP): ${formatKmAndMiles(marketDetected.distKm)}${(marketDetected.distKm!=null && Number.isFinite(marketDetected.distKm)) ? ` • ETA: ${formatHoursHM(kmToMarketHours(marketDetected.distKm))}` : ``} ${marketDetected.confidence ? ` • Confidence: <b>${marketDetected.confidence}</b>` : ``} ${(marketDetected.deltaKm!=null && Number.isFinite(marketDetected.deltaKm)) ? ` • Next Δ: ${formatKmAndMiles(marketDetected.deltaKm)}` : ``}</div>
      ${confidenceNoticeHtml(marketDetected.confidence, marketDetected.distKm, marketDetected.deltaKm)}
      ${(() => {
        const ws = marketDetected?.warnings || [];
        if(!ws.length) return '';
        return `<div class="warn-box"><b>Warnings:</b><ul>${ws.map(w=>`<li>${escapeHtml(w)}</li>`).join('')}</ul></div>`;
      })()}

      ${(() => {
         const c = String(marketDetected.confidence||'').toLowerCase();
         const show = (c==='high' || c==='medium' || c==='low') && marketDetected.secondMarket && marketDetected.secondDistKm!=null;
         if(!show) return '';
         return `
           <div class="market-choice">
             <div class="muted" style="margin-bottom:6px"><b>Choose Tech:</b> Select one of the two closest options (Top 2):</div>
             <div class="choice-grid">
               <button class="btn" id="btnUseBest">Use: ${marketDetected.market.displayName}${marketDetected.techNameHint ? ` — ${marketDetected.techNameHint}` : ``} — ${formatKmAndMiles(marketDetected.distKm)}</button>
               <button class="btn" id="btnUseSecond">Use: ${marketDetected.secondMarket.displayName}${(marketDetected.alternatives && marketDetected.alternatives[1] && marketDetected.alternatives[1].techName) ? ` — ${marketDetected.alternatives[1].techName}` : ``} — ${formatKmAndMiles(marketDetected.secondDistKm)}</button>
             </div>
             <div class="muted" style="margin-top:6px">Tip: If the ticket is near a boundary, City/State usually confirms the correct market.</div>
           </div>
         `;
      })()}

    </div>
    <div class="row" style="margin-top:8px">
      <div class="badge">Input ZIP: ${inputZip}</div>
      <div class="badge">City/State: ${marketDetected.zipInfo.city.toUpperCase()} / ${marketDetected.zipInfo.state}</div>
      <div class="badge">Date: ${dateStr}${ampm ? ' ' + ampm : ''}</div>
      <div class="badge">Range: ${ymd(week.start)} → ${ymd(week.end)}</div>
    </div>
    <hr style="border:0;border-top:1px solid rgba(148,163,184,.18);margin:12px 0">
    ${requireChoice ? `<div class="warn-card"><div class="warn-title">⚠️ Confirmation required</div><div class="warn-body">Two options were found     On‑Call. Please choose one from the <b>Top 2</b> options below.</div></div>` : `<div><b>On‑Call Cell</b>: ${techCell ? (techCell.raw+' '+reservedBadge) : '<span class="muted">Empty</span>'}</div>`}
    ${requireChoice ? '' : `<div style="margin-top:10px">
      <b>Tech Details</b> ${techFoundBadge}
      ${reservedBlocked ? `
        <div class="warn-card">
          <div class="warn-title">⚠️ RESERVED</div>
          <div class="warn-body">This region/week is marked Reserved.</div>
          <div class="warn-body">Because the chosen date <b>${escapeHtml(dateStr)}</b> is not today's date, tech details will not be displayed.</div>
          <div class="warn-body">Please choose another tech (not Reserved).</div>
        </div>
      ` : `
      <div class="kv" style="margin-top:8px">
        <div class="key">Tech ID</div><div>${techCell?.techId ?? '<span class="muted">N/A</span>'}</div>
        <div class="key">Name</div><div>${techNotFoundReason ? `<div class="warn-box warn-red"><b>Tech lookup issue:</b> ${escapeHtml(techNotFoundReason)}<div class="muted" style="margin-top:6px">Result is hidden to avoid dispatching the wrong person. Upload updated Tech DB.</div></div>` : ``}
      ${(!techNotFoundReason && tech) ? `${tech.firstName} ${tech.lastName}` : '<span class="muted">Not found</span>'}</div>
        <div class="key">City/State</div><div>${tech ? `${tech.city}, ${tech.state}` : '<span class="muted">-</span>'}</div>
        <div class="key">ZIP</div><div>${tech ? (tech.zip ?? '-') : '<span class="muted">-</span>'}</div>
        <div class="key">Region / Zone</div><div>${tech ? `${tech.region} / ${tech.zone}` : '<span class="muted">-</span>'}</div>
        <div class="key">Type</div><div>${tech ? tech.type : '<span class="muted">-</span>'}</div>
      </div>
      `}
      </div>`}

    ${nonAvailHtml}
  `;

  // Market chooser (for Medium/Low confidence): allow dispatcher to pick between top 2 markets
  try{
    const c = String(marketDetected.confidence||'').toLowerCase();
    const show = (c==='high' || c==='medium' || c==='low') && marketDetected?.alternatives?.length >= 2;
    if(show){
      const btnBest = document.getElementById('btnUseBest');
      const btnSecond = document.getElementById('btnUseSecond');
      if(btnBest){
        btnBest.addEventListener('click', ()=>{
          lookupWithChosenMarket(marketDetected.alternatives[0]);
        });
      }
      if(btnSecond){
        btnSecond.addEventListener('click', ()=>{
          lookupWithChosenMarket(marketDetected.alternatives[1]);
        });
      }
    }
  }catch(_){}
}


let __lastLookup = null;

function lookupWithChosenMarket(choice){
  if(!__lastLookup) return;
  const {inputZip, dateStr, enteredDateStr, ampm, selectedState, baseDetected} = __lastLookup;
  const weekIdx = pickWeekIndexForDate(dateStr, ampm);
  const week = oncallMeta.weeks[weekIdx];

  const chosenMarket = choice.market;
  const techCell = parseTechCell(oncallSheetAOA[chosenMarket.row][week.col]);
  const tech = techCell?.techId ? (techMap.get(techCell.techId)||null) : null;
  let techNotFoundReason = '';
  if(techCell?.techId && !tech){ techNotFoundReason = `Tech ID ${techCell.techId} not found in Tech DB. Please update Tech DB.`; }
  if(tech && !normalizeZip(tech.zip || tech.Zip || tech.ZIP)){ techNotFoundReason = 'Tech ZIP missing in Tech DB. Please update Tech DB.'; }

  const techZipHint = tech ? normalizeZip(tech.zip || tech.Zip || tech.ZIP) : null;

  const zipNorm = normalizeZip(inputZip);
  const zipInfo = zipDB.get(zipNorm);
  const marketDetected = {
    zip: zipNorm,
    zipInfo,
    market: chosenMarket,
    distKm: choice.distKm,
    techZip: techZipHint,
    techNameHint: tech ? `${tech.firstName||''} ${tech.lastName||''}`.trim() : null,
    confidence: baseDetected?.confidence || 'Low',
    deltaKm: baseDetected?.deltaKm ?? null,
    candidatesCount: baseDetected?.candidatesCount ?? null,
    secondMarket: baseDetected?.secondMarket ?? null,
    secondDistKm: baseDetected?.secondDistKm ?? null,
    alternatives: baseDetected?.alternatives ?? [choice],
    userSelected: true
  };

  const marketInfoText = String(marketDetected?.market?.info || marketDetected?.market?.note || '');
  const isReserved = !!(techCell?.reserved) || /\breserv/i.test(marketInfoText);
  const isToday = (dateStr === effectiveTodayYmd());

  if(isReserved && !isToday){
    renderResult({inputZip,dateStr,enteredDateStr,ampm,marketDetected,week,techCell,tech:null,selectedState,reservedBlocked:true,techNotFoundReason});
    return;
  }

  renderResult({inputZip,dateStr,enteredDateStr,ampm,marketDetected,week,techCell,tech,selectedState,techNotFoundReason});
}

// UI buttons
$('btnLoadTech').addEventListener('click', async ()=>{
  try{
    const f = $('techFile').files?.[0];

    // If user uploaded a file (for updates) -> load from file.
    if(f){
      setStatus('Loading Tech DB from file...');
      const {count, sheet} = await loadTechDb(f);
      setStatus(`Tech DB Ready (from file): ${count.toLocaleString()} — Sheet: ${sheet}`);
      return;
    }

    // Otherwise try Cloud API first (if configured).
    const base = apiBase();
    if(base){
      try{
        setStatus('Loading Tech DB from cloud API...');
        const j = await fetchJson(`${base}/api/oncall/techdb`);
        if(j && j.ok && Array.isArray(j.rows) && j.rows.length){
          window.TECH_DB_ROWS = j.rows;
          const r = loadBuiltInTechDb();
          setProg('techProg', 100);
          setStatus(`Tech DB Ready (cloud): ${r.count.toLocaleString()}`);
          return;
        }
      }catch(e){
        console.warn('Tech API load failed; using bundled techdb.js', e);
      }
    }

    // Otherwise use bundled Tech DB.
    setStatus('Loading built-in Tech DB...');
    const r = loadBuiltInTechDb();
    if(!r.ok || r.count===0) throw new Error('Built-in TECH DB is not available. (Ensure techdb.js exists)');
    setProg('techProg', 100);
    setStatus(`Tech DB Ready (built-in): ${r.count.toLocaleString()}`);
  }catch(e){
    setProg('techProg', 0);
    setStatus(String(e.message||e), true);
  }
});

$('btnLoadOncall').addEventListener('click', async ()=>{
  try{
    const f=$('oncallFile').files?.[0];
    if(!f) throw new Error('Choose the OnCall file.');

    if(zipDB.size===0){
      setStatus('Loading local ZIP DB...');
      await loadBundledZipDb();
    }

    setStatus('Loading and processing OnCall...');
    const {meta,sheet}=await loadOncall(f);
    setStatus(`OnCall Ready — Sheet: ${sheet} | Markets: ${meta.markets.length} | Weeks: ${meta.weeks.length}`);
    try{ $('dateInput').dispatchEvent(new Event('change')); }catch(_){ }
  }catch(e){
    setProg('oncallProg', 0);
    setStatus(String(e.message||e), true);
  }
});

$('btnDownload').addEventListener('click', ()=>{
  if(!cleanedOncallBlob) return;
  const a=document.createElement('a');
  a.href=URL.createObjectURL(cleanedOncallBlob);
  a.download='OnCall_FirstSheet_Cleaned_MarketsToZips.xlsx';
  a.click();
  URL.revokeObjectURL(a.href);
});

$('btnLookup').addEventListener('click', async ()=>{
  try{
    if(zipDB.size===0){
      setStatus('Loading local ZIP DB...');
      await loadBundledZipDb();
    }
    if(!oncallMeta) throw new Error('Load OnCall first.');
    if(techMap.size===0) throw new Error('Load Tech DB first.');

    let inputZip = normalizeZip($('zipInput').value);
    if(!inputZip){
      const z = findZipByCityState($('cityInput').value, $('stateInput').value);
      if(!z) throw new Error('Could not convert City/State to ZIP. Please enter a valid City+State or enter a ZIP.');
      inputZip = z;
    }

    const enteredDateStr = String($('dateInput').value||'').trim();
    const dateStr = effectiveInputDateYmd(enteredDateStr);
    const ampm = getSelectedAmPm();
    const selectedState = normState($('stateInput').value);

    if(!selectedState || selectedState.length !== 2){
      throw new Error('State is required (2 letters like TX, CA).');
    }


    if(!selectedState || selectedState.length !== 2){
      throw new Error('State is required (2 letters like TX, CA).');
    }


    const weekIdx=pickWeekIndexForDate(dateStr, ampm);
    const week=oncallMeta.weeks[weekIdx];

    const marketDetected=detectMarketForZip(inputZip, selectedState, week, dateStr);
    __lastLookup = { inputZip, dateStr, enteredDateStr, ampm, selectedState, baseDetected: marketDetected };

    const techCell=parseTechCell(oncallSheetAOA[marketDetected.market.row][week.col]);
    const tech = techCell?.techId ? (techMap.get(techCell.techId)||null) : null;
    let techNotFoundReason = '';
    if(techCell?.techId && !tech){ techNotFoundReason = `Tech ID ${techCell.techId} not found in Tech DB. Please upload the updated Tech DB.`; }
    const __techZip = tech ? normalizeZip(tech.zip || tech.Zip || tech.ZIP) : '';
    if(tech && !__techZip){ techNotFoundReason = techNotFoundReason || 'Tech ZIP is missing in Tech DB. Please upload the updated Tech DB.'; }
    // Reserved-today rule:
    // If this market/week is marked RESERVED, only allow showing tech details when date == today.
    // Otherwise, block tech display and tell dispatcher to choose another tech.
    const marketInfoText = String(marketDetected?.market?.info || marketDetected?.market?.note || '');
    const isReserved = !!(techCell?.reserved) || /\breserv/i.test(marketInfoText);
    const isToday = (dateStr === effectiveTodayYmd());

    if(isReserved && !isToday){
      renderResult({inputZip,dateStr,enteredDateStr,ampm,marketDetected,week,techCell,tech:null,selectedState,reservedBlocked:true,techNotFoundReason});
      return;
    }


    renderResult({inputZip,dateStr,enteredDateStr,ampm,marketDetected,week,techCell,tech,selectedState,techNotFoundReason});
    setStatus('Done.');
  }catch(e){
    setStatus(String(e.message||e), true);
  }
});

// Show AM/PM chooser only when selected date equals an End Date in the loaded OnCall sheet.
$('dateInput').addEventListener('change', ()=>{
  try{
    const dateStr = effectiveInputDateYmd($('dateInput').value);
    if(!oncallMeta || !dateStr){
      setAmPmVisible(false);
      return;
    }
    const d = new Date(dateStr + 'T00:00:00');
    if(isNaN(d.getTime())) { setAmPmVisible(false); return; }
    const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());

    const starts = oncallMeta.weeks.map(w => ymd(w.start));
    const isBoundary = (starts.indexOf(dateStr) > 0);
    setAmPmVisible(isBoundary);
  }catch(_){
    setAmPmVisible(false);
  }
});