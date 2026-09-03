import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6';

const supabase = createClient(supabaseUrl, supabaseKey);

async function run() {
  const { data, error } = await supabase
    .from('tariffs')
    .update({ type: 'electricity' })
    .eq('tariff_id', 'TF01');

  if (error) {
    console.error("Error updating tariff TF01:", error);
  } else {
    console.log("Tariff TF01 updated successfully to 'electricity'!");
  }
}

run();
