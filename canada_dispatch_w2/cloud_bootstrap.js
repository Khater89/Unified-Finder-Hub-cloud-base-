// Canada Dispatch W2 — Cloud bootstrap
// Loads CA W2 tech list from Supabase Edge Function before loading app.js
(async function(){
  const API_BASE = String(window.API_BASE || "").replace(/\/$/, "");

  // Defaults (fallback)
  window.CA_W2_TECHS = Array.isArray(window.CA_W2_TECHS) ? window.CA_W2_TECHS : [];
  window.CA_POSTAL_PROV = (window.CA_POSTAL_PROV && typeof window.CA_POSTAL_PROV === 'object') ? window.CA_POSTAL_PROV : {};

  async function getJson(url){
    const r = await fetch(url, { method: 'GET' });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return await r.json();
  }

  // Try cloud tech list (small + fast). If it fails, we keep whatever is already embedded.
  if (API_BASE) {
    try {
      const j = await getJson(`${API_BASE}/api/canada/w2techdb`);
      if (j && j.ok && Array.isArray(j.techs) && j.techs.length) {
        window.CA_W2_TECHS = j.techs;
      }
    } catch (e) {
      // ignore — offline fallback
      console.warn('Canada cloud techdb load failed:', e);
    }

    // Optional: cloud postal->province mapping (can be large)
    // Enable only if you configured CA_POSTAL_TABLE and want better precision.
    // try {
    //   const j2 = await getJson(`${API_BASE}/api/canada/postalprov`);
    //   if (j2 && j2.ok && j2.mapping && typeof j2.mapping === 'object') {
    //     window.CA_POSTAL_PROV = j2.mapping;
    //   }
    // } catch (e) { /* ignore */ }
  }

  // Now load the main app
  const s = document.createElement('script');
  s.src = 'app.js';
  document.body.appendChild(s);
})();
