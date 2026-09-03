import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://blxyhsggqgrwvhbqeqhz.supabase.co';
const supabaseKey = 'sb_publishable_AJSh01HH6mZziV92K3jo7g_fStCoum6';

const supabase = createClient(supabaseUrl, supabaseKey);

async function register() {
    console.log("Registering user software.2040@pgel.in in Supabase...");
    const { data, error } = await supabase
        .from('users')
        .upsert([
            {
                id: 'software_2040',
                name: 'Vikas Kumar',
                email: 'software.2040@pgel.in',
                role: 'IT_ADMIN',
                allowed_locations: 'all',
                allowed_plants: 'all',
                status: 'Active',
                created_date: new Date().toISOString().split('T')[0]
            }
        ]);

    if (error) {
        console.error("Error registering user:", error);
    } else {
        console.log("Successfully registered software.2040@pgel.in as IT_ADMIN!");
    }
}

register();
