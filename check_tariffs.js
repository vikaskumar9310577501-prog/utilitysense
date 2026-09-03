import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6'; // Let's use the public key from check_tables.js

const supabase = createClient(supabaseUrl, supabaseKey);

async function run() {
  const { data, error } = await supabase.from('tariffs').select('*');
  if (error) {
    console.error("Error reading tariffs:", error);
  } else {
    console.log("Tariffs in DB:", JSON.stringify(data, null, 2));
  }
}

run();
