import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://dxnvsmlmocxuwclibphk.supabase.co';
const supabaseAnonKey = 'sb_publishable_qM48-2xT9rogd8JxaR8HLg_Hps-knW_';

const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function check() {
  const { data, error } = await supabase.from('datasets').select('*').limit(1);
  console.log("DATA:", data);
  console.log("ERROR:", error);
}

check();
