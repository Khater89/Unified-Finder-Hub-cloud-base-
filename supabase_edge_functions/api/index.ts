// Supabase Edge Function: api
// Put this file at: supabase/functions/api/index.ts (using Supabase CLI)
// Then deploy: supabase functions deploy api

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const TECH_TABLE = Deno.env.get("TECH_TABLE") || "";
const ZIP_TABLE  = Deno.env.get("ZIP_TABLE")  || "";
const FLEX_TABLE = Deno.env.get("FLEX_TABLE") || "";

// Canada (optional)
const CA_W2_TABLE = Deno.env.get("CA_W2_TABLE") || "";
const CA_POSTAL_TABLE = Deno.env.get("CA_POSTAL_TABLE") || ""; // optional mapping postal->province

const ALLOWED_TABLES = new Set([TECH_TABLE, ZIP_TABLE, FLEX_TABLE, CA_W2_TABLE, CA_POSTAL_TABLE].filter(Boolean));

const supabase = createClient(SUPABASE_URL, SERVICE_ROLE, {
  auth: { persistSession: false },
});

function cors(origin: string | null) {
  return {
    "Access-Control-Allow-Origin": origin ?? "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "content-type, authorization, apikey",
  };
}

function pick(obj: any, keys: string[], fallback = "") {
  for (const k of keys) {
    if (obj && obj[k] != null && String(obj[k]).trim() !== "") return obj[k];
  }
  return fallback;
}

Deno.serve(async (req) => {
  const url = new URL(req.url);
  const origin = req.headers.get("origin");

  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: cors(origin) });
  }

  let p = url.pathname;
  p = p.replace(/^\/functions\/v1\/api/, "");
  if (!p.startsWith("/")) p = "/" + p;

  try {
    // /oncall/techdb
    if (p === "/oncall/techdb") {
      const table = url.searchParams.get("table") || TECH_TABLE;
      const limit = Math.min(Math.max(Number(url.searchParams.get("limit") || 200000), 1), 200000);

      const { data, error } = await supabase.from(table).select("*").range(0, limit - 1);
      if (error) throw error;

      const rows = (data || []).map((x: any) => ([
        String(pick(x, ["tech_id","techid","tech","id","col_1"], "")),
        String(pick(x, ["first_name","firstname","col_2"], "")),
        String(pick(x, ["last_name","lastname","col_3"], "")),
        String(pick(x, ["region","col_4"], "")),
        String(pick(x, ["zone","col_5"], "")),
        String(pick(x, ["type","col_6"], "")),
        String(pick(x, ["city","col_7"], "")),
        String(pick(x, ["state","col_8"], "")),
        String(pick(x, ["zip","zip_code","col_9"], "")),
      ]));

      return new Response(JSON.stringify({ ok: true, table, rows }), {
        headers: { ...cors(origin), "Content-Type": "application/json" },
      });
    }

    // /oncall/uszips
    if (p === "/oncall/uszips") {
      const table = url.searchParams.get("table") || ZIP_TABLE;
      const limit = Math.min(Math.max(Number(url.searchParams.get("limit") || 200000), 1), 200000);

      const { data, error } = await supabase.from(table).select("*").range(0, limit - 1);
      if (error) throw error;

      const rows = (data || []).map((x: any) => ([
        String(pick(x, ["zip","zip_code","col_1"], "")),
        Number(pick(x, ["lat","latitude","col_2"], null)),
        Number(pick(x, ["lon","longitude","col_3"], null)),
        String(pick(x, ["city","col_4"], "")).trim().toLowerCase(),
        String(pick(x, ["state","col_5"], "")).trim().toUpperCase(),
      ]));

      return new Response(JSON.stringify({ ok: true, table, rows }), {
        headers: { ...cors(origin), "Content-Type": "application/json" },
      });
    }

    // /canada/w2techdb  -> returns tech objects: {tech_id,name,city,province,postal}
    if (p === "/canada/w2techdb") {
      const table = url.searchParams.get("table") || CA_W2_TABLE;
      if (!table) {
        return new Response(JSON.stringify({ ok: false, error: "CA_W2_TABLE not set" }), {
          status: 400,
          headers: { ...cors(origin), "Content-Type": "application/json" },
        });
      }

      const limit = Math.min(Math.max(Number(url.searchParams.get("limit") || 5000), 1), 50000);
      const { data, error } = await supabase.from(table).select("*").range(0, limit - 1);
      if (error) throw error;

      const techs = (data || []).map((x: any) => ({
        tech_id: String(pick(x, ["tech_id", "techid", "id", "tech"], "")),
        name: String(pick(x, ["name", "full_name", "tech_name", "col_2", "col_1"], "")),
        city: String(pick(x, ["city", "col_3"], "")),
        province: String(pick(x, ["province", "prov", "state", "col_4"], "")).toUpperCase(),
        postal: String(pick(x, ["postal", "postal_code", "postalcode", "postcode", "col_5", "zip"], "")).replace(/\s+/g, "").toUpperCase(),
      }));

      return new Response(JSON.stringify({ ok: true, table, techs }), {
        headers: { ...cors(origin), "Content-Type": "application/json" },
      });
    }

    // /canada/postalprov  -> returns mapping object {POSTAL6: "PR"}
    // Optional. If not configured, returns empty mapping.
    if (p === "/canada/postalprov") {
      const table = url.searchParams.get("table") || CA_POSTAL_TABLE;
      if (!table) {
        return new Response(JSON.stringify({ ok: true, table: "", mapping: {}, disabled: true }), {
          headers: { ...cors(origin), "Content-Type": "application/json" },
        });
      }

      const limit = Math.min(Math.max(Number(url.searchParams.get("limit") || 50000), 1), 200000);
      const { data, error } = await supabase.from(table).select("*").range(0, limit - 1);
      if (error) throw error;

      const mapping: Record<string, string> = {};
      for (const x of (data || []) as any[]) {
        const postal = String(pick(x, ["postal", "postal_code", "postalcode", "col_1"], "")).replace(/\s+/g, "").toUpperCase();
        const prov = String(pick(x, ["province", "prov", "state", "col_2"], "")).toUpperCase();
        if (postal && prov) mapping[postal] = prov;
      }

      return new Response(JSON.stringify({ ok: true, table, mapping }), {
        headers: { ...cors(origin), "Content-Type": "application/json" },
      });
    }

    // /export/<table>?format=aoa&limit=...
    if (p.startsWith("/export/")) {
      const table = decodeURIComponent(p.split("/")[2] || "");
      if (ALLOWED_TABLES.size && !ALLOWED_TABLES.has(table)) {
        return new Response(JSON.stringify({ ok: false, error: "Table not allowed" }), {
          status: 403,
          headers: { ...cors(origin), "Content-Type": "application/json" },
        });
      }

      const format = (url.searchParams.get("format") || "json").toLowerCase();
      const limit = Math.min(Math.max(Number(url.searchParams.get("limit") || (format === "aoa" ? 200000 : 500)), 1), 200000);

      const { data, error } = await supabase.from(table).select("*").range(0, limit - 1);
      if (error) throw error;

      if (format === "aoa") {
        const first = (data && data[0]) ? data[0] : {};
        const columns = Object.keys(first).filter((c) => c !== "_id");
        const rows = (data || []).map((r: any) => columns.map((c) => r[c] ?? null));
        return new Response(JSON.stringify({ ok: true, table, columns, rows }), {
          headers: { ...cors(origin), "Content-Type": "application/json" },
        });
      }

      return new Response(JSON.stringify({ ok: true, table, rows: data || [] }), {
        headers: { ...cors(origin), "Content-Type": "application/json" },
      });
    }

    return new Response(JSON.stringify({ ok: false, error: "Not found" }), {
      status: 404,
      headers: { ...cors(origin), "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(JSON.stringify({ ok: false, error: String((e as any)?.message || e) }), {
      status: 500,
      headers: { ...cors(origin), "Content-Type": "application/json" },
    });
  }
});
