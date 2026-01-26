Supabase Edge Functions helper
================================

This project ZIP was updated so the Frontend apps use:
  API_BASE = https://whocxxcqnjhvqmsldbkz.supabase.co/functions/v1

To make it work, you must deploy a Supabase Edge Function named: api

Quick steps (Supabase CLI):
1) Install CLI and login
2) supabase init
3) Put the function file at: supabase/functions/api/index.ts
   (You can copy from: supabase_edge_functions/api/index.ts in this ZIP)
4) Set secrets:
   supabase secrets set SUPABASE_SERVICE_ROLE_KEY=...
   supabase secrets set TECH_TABLE=...
   supabase secrets set ZIP_TABLE=...
   supabase secrets set FLEX_TABLE=...
   # Optional (Canada Dispatch W2)
   supabase secrets set CA_W2_TABLE=...
   # Optional (postal -> province mapping, can be large)
   supabase secrets set CA_POSTAL_TABLE=...
5) Deploy:
   supabase functions deploy api

Test:
  https://whocxxcqnjhvqmsldbkz.supabase.co/functions/v1/api/oncall/uszips
  https://whocxxcqnjhvqmsldbkz.supabase.co/functions/v1/api/oncall/techdb
  https://whocxxcqnjhvqmsldbkz.supabase.co/functions/v1/api/export/<your_table>?format=aoa&limit=5
