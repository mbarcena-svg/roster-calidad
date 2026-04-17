/* eslint-disable no-console */
const fs = require("fs");
const path = require("path");
const { createClient } = require("@supabase/supabase-js");

const SUPABASE_TABLE = process.env.SUPABASE_TABLE || "roster_store";
const url = process.env.SUPABASE_URL;
const key =
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY || process.env.SUPABASE_KEY;

if (!url || !key) {
  console.error("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY/ANON_KEY.");
  process.exit(1);
}

const supabase = createClient(url, key, { auth: { persistSession: false } });

async function main() {
  const storePath = path.join(__dirname, "..", "data", "store.json");
  const raw = fs.readFileSync(storePath, "utf8");
  const data = JSON.parse(raw);

  const payload = { id: 1, data, updated_at: new Date().toISOString() };
  const { error } = await supabase.from(SUPABASE_TABLE).upsert(payload, { onConflict: "id" });
  if (error) throw error;

  console.log(`OK: seed completo en tabla '${SUPABASE_TABLE}' (id=1).`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

