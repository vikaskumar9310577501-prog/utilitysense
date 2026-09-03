import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6';

const supabase = createClient(supabaseUrl, supabaseKey);

async function test() {
    console.log("Checking which tables exist in Supabase database...");
    const tables = [
        'plants', 
        'departments', 
        'meters', 
        'solar_meters', 
        'water_meters', 
        'air_meters', 
        'dg_sets', 
        'fuel_types', 
        'products', 
        'tariffs', 
        'target_values', 
        'users', 
        'otp_logs', 
        'daily_entries'
    ];
    
    for (const t of tables) {
        const { error } = await supabase.from(t).select('*').limit(1);
        if (error) {
            console.log(`Table "${t}": MISSING (${error.message})`);
        } else {
            console.log(`Table "${t}": EXISTS`);
        }
    }
}

test();
