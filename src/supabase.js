import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL      = import.meta.env.VITE_SUPABASE_URL
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY

// persistSession: false — the login session lives only in memory for this
// browser tab. Closing the tab, closing the app, or refreshing the page all
// require logging in again. This is intentional: a shared/borrowed phone
// should never stay logged in for whoever opens the site next.
export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: false,
    autoRefreshToken: true, // still refreshes the token while actively in use, just doesn't save it anywhere
  },
})
export default supabase