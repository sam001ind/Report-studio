import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://dxnvsmlmocxuwclibphk.supabase.co'
const supabaseAnonKey = 'sb_publishable_qM48-2xT9rogd8JxaR8HLg_Hps-knW_'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)
