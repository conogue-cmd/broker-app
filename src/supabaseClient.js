import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://hegyzgwvdinsrqylcova.supabase.co";
const supabaseKey = "sb_publishable_2NhDEya25Ac2pfYBU-8R8w_ckJfCuGi";

export const supabase = createClient(supabaseUrl, supabaseKey);
