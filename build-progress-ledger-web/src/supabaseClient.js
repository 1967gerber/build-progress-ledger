import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // Shows up clearly in the browser console instead of a silent failure —
  // almost always means the Environment Variables weren't set in Render.
  console.error(
    "Missing Supabase environment variables. Set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY " +
    "(in a local .env file for development, or in Render's Environment tab for deployment)."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
