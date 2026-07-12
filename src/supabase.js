// ─────────────────────────────────────────────────────────────────────────────
//  SAKER BAPTIST COLLEGE — Supabase Configuration
// ─────────────────────────────────────────────────────────────────────────────
//  HOW TO GET THESE VALUES:
//  1. Go to https://supabase.com and create a free account
//  2. Click "New project" → name it "saker-admin" → set a database password
//  3. Wait ~2 minutes for the project to be ready
//  4. Go to: Project Settings → API
//  5. Copy "Project URL"  → paste below as SUPABASE_URL
//  6. Copy "anon public" key → paste below as SUPABASE_ANON_KEY
// ─────────────────────────────────────────────────────────────────────────────

import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL      = "PASTE_YOUR_PROJECT_URL_HERE"
const SUPABASE_ANON_KEY = "PASTE_YOUR_ANON_KEY_HERE"

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
export default supabase
