// Supabase project connection — the anon key is a public, RLS-scoped identifier,
// not a secret. Safe to ship in client code as long as every table has RLS
// policies scoping access to auth.uid() = user_id (see the SQL setup script).
const SUPABASE_URL = "https://dfyzhdakwievhvpuvtmz.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_KwMTgjU504mLxLJ-vd_XIw_f7RvAtwj";
