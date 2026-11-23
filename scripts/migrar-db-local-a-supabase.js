"use strict";

const fs = require("fs");
const path = require("path");
const { SupabaseStore } = require("../supabase-store");

const COLLECTIONS = [
  "empresa",
  "usuarios",
  "areas",
  "horarios_especiales",
  "asistencia",
  "pendientes",
  "periodo_laboral",
  "vacaciones_solicitud",
  "sesiones",
  "superadmins",
  "superadmin_sesiones",
  "pagos_suscripcion",
  "push_suscripciones",
  "notificaciones_enviadas",
  "auditoria_superadmin"
];

async function main() {
  const args = process.argv.slice(2);
  const apply = args.includes("--apply");
  const dbFileArgument = args.find(arg => !arg.startsWith("--")) || "db.json";
  const dbFile = path.resolve(process.cwd(), dbFileArgument);

  if (!fs.existsSync(dbFile)) throw new Error("No existe el archivo local: " + dbFile);

  const payload = JSON.parse(fs.readFileSync(dbFile, "utf8"));
  const sourceDb = payload.db || payload;
  const db = {};
  let total = 0;

  for (const collection of COLLECTIONS) {
    const rows = sourceDb[collection] || [];
    if (!Array.isArray(rows)) throw new Error("La colección " + collection + " no es una lista.");
    const ids = new Set();
    db[collection] = rows.map(row => {
      const id = Number(row?.id);
      if (!Number.isSafeInteger(id) || id < 1) {
        throw new Error("Se encontró un id inválido en la colección " + collection + ".");
      }
      if (ids.has(id)) throw new Error("Hay ids duplicados en la colección " + collection + ".");
      ids.add(id);
      return { ...row, id };
    });
    total += rows.length;
  }

  const resumen = COLLECTIONS
    .filter(collection => db[collection].length)
    .map(collection => collection + ": " + db[collection].length)
    .join("\n");
  console.log("Archivo local: " + dbFile);
  console.log("Registros a migrar: " + total);
  console.log(resumen || "(sin registros)");

  if (!apply) {
    console.log("\nVista previa solamente. Para escribir en Supabase, vuelve a ejecutar con --apply.");
    return;
  }

  if (!process.env.SUPABASE_URL) throw new Error("Falta SUPABASE_URL.");
  const key = process.env.SUPABASE_SECRET_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error("Falta SUPABASE_SECRET_KEY (o SUPABASE_SERVICE_ROLE_KEY).");

  const store = new SupabaseStore({
    url: process.env.SUPABASE_URL,
    key,
    collections: COLLECTIONS
  });
  const existing = await store.load();
  const remoteCount = COLLECTIONS.reduce((sum, collection) => sum + existing.db[collection].length, 0);
  if (remoteCount > 0) {
    throw new Error(
      "Supabase ya tiene " + remoteCount + " registros. Se canceló para evitar sobrescribir datos."
    );
  }

  const emptyDb = Object.fromEntries(COLLECTIONS.map(collection => [collection, []]));
  const result = await store.persistChanges(db, emptyDb);
  console.log(
    "\nMigración terminada. Registros insertados: " + result.upserted +
    ". Conserva una copia segura del db.json hasta revisar la aplicación."
  );
}

main().catch(error => {
  console.error("No se completó la migración:", error.message);
  process.exitCode = 1;
});
