"use strict";

/**
 * Genera el usuario SUPERADMIN de JJM Clock.
 *
 * El backend guarda contraseñas con:
 *   scrypt:<salt base64>:<hash base64>   (64 bytes)
 * y los registros viven en public.jjm_records como JSONB,
 * colección "superadmins".
 *
 * Formas de uso:
 *   1) Solo imprimir el SQL (no necesita credenciales):
 *      node scripts/generar-superadmin.js
 *
 *   2) Subirlo directamente a Supabase ( credenciales SOLO en variables de entorno,
 *      nunca pegarlas en el chat ni en el repo ):
 *      $env:SUPABASE_URL="https://XXXX.supabase.co"
 *      $env:SUPABASE_SECRET_KEY="sb_secret_XXXX"
 *      node scripts/generar-superadmin.js --upload
 *
 *   Opcionalmente cambia los datos por entorno:
 *      $env:SUPERADMIN_NOMBRE="Tu nombre"
 *      $env:SUPERADMIN_CORREO="superadmin@tudominio.com"
 *      $env:SUPERADMIN_PASSWORD="una contraseña de al menos 12 caracteres"
 */

const crypto = require("crypto");
const { promisify } = require("util");
const scryptAsync = promisify(crypto.scrypt);

const NOMBRE = process.env.SUPERADMIN_NOMBRE || "JJM Solutions Super Admin";
const CORREO = String(process.env.SUPERADMIN_CORREO || "superadmin@gruponerba.com").trim().toLowerCase();
// La contraseña NUNCA se guarda en el repo: debe pasar por variable de entorno.
const PASSWORD = process.env.SUPERADMIN_PASSWORD || "";

async function guardarPassword(password) {
  const salt = crypto.randomBytes(16);
  const hash = await scryptAsync(password, salt, 64);
  return `scrypt:${salt.toString("base64")}:${hash.toString("base64")}`;
}

function escapaSQL(texto) {
  return String(texto).replace(/'/g, "''");
}

async function main() {
  if (PASSWORD.length < 12) {
    console.error("Define SUPERADMIN_PASSWORD con al menos 12 caracteres antes de ejecutar.");
    console.error("Ejemplo: $env:SUPERADMIN_PASSWORD='tu contraseña segura'");
    process.exit(1);
  }

  const passwordHash = await guardarPassword(PASSWORD);
  const id = 1;
  const registro = {
    id,
    nombre: NOMBRE,
    correo: CORREO,
    password_hash: passwordHash,
    activo: 1,
    creado_en: new Date().toISOString()
  };

  const jsonRecord = JSON.stringify(registro).replace(/'/g, "''");
  const sql = `-- JJM Clock · crear superadmin
-- Pega esto en Supabase → SQL Editor → New query → Run
INSERT INTO public.jjm_records (collection, id, record)
VALUES (
  'superadmins',
  ${id},
  '${jsonRecord}'::jsonb
)
ON CONFLICT (collection, id) DO UPDATE
SET record = EXCLUDED.record, updated_at = now();

-- Verificación:
-- SELECT collection, id, record->>'correo' AS correo, record->>'activo' AS activo
-- FROM public.jjm_records WHERE collection = 'superadmins';
`;

  console.log("=== DATOS DEL SUPERADMIN ===");
  console.log("Nombre:   " + NOMBRE);
  console.log("Correo:   " + CORREO);
  console.log("Password: (la definiste en SUPERADMIN_PASSWORD)");
  console.log("Hash:     " + passwordHash);
  console.log("");
  console.log("=== SQL PARA SUPABASE ===");
  console.log(sql);

  if (!process.argv.includes("--upload")) {
    console.log("Si quieres subirlo automáticamente, configura SUPABASE_URL y");
    console.log("SUPABASE_SECRET_KEY en el entorno y ejecuta con --upload.");
    return;
  }

  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SECRET_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) {
    console.error("Faltan SUPABASE_URL y SUPABASE_SECRET_KEY en el entorno.");
    process.exit(1);
  }

  const headers = {
    apikey: key,
    Accept: "application/json",
    "Content-Type": "application/json",
    Prefer: "resolution=merge-duplicates,return=minimal"
  };
  if (!key.startsWith("sb_secret_")) headers.Authorization = "Bearer " + key;

  const endpoint = url.replace(/\/+$/, "") +
    "/rest/v1/jjm_records?on_conflict=collection%2Cid";

  const response = await fetch(endpoint, {
    method: "POST",
    headers,
    body: JSON.stringify([{
      collection: "superadmins",
      id,
      record: registro,
      updated_at: new Date().toISOString()
    }])
  });

  const text = await response.text();
  if (!response.ok) {
    console.error("Supabase respondió " + response.status + ": " + text.slice(0, 800));
    process.exit(1);
  }
  console.log("✅ Superadmin creado/actualizado en Supabase.");
}

main().catch((error) => {
  console.error("Error:", error.message);
  process.exit(1);
});
