import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6';

const supabase = createClient(supabaseUrl, supabaseKey);

async function run() {
  const { data, error } = await supabase.from('daily_entries').select('*').limit(1);
  if (error) {
    console.error("Error fetching daily_entries:", error);
  } else {
    console.log("Columns of daily_entries:", data.length > 0 ? Object.keys(data[0]) : "No rows");
  }
}

run();
