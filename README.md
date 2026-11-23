# JJM Clock API

Backend Node.js + Express de JJM Clock. Este repositorio contiene la API; la interfaz vive en `jjmclock-frontend`.

## Desarrollo local

1. Instala Node.js 24 o compatible.
2. Ejecuta `npm ci`.
3. Configura las variables de `.env.example` en el entorno de tu terminal. La app no lee `.env` automáticamente.
4. Ejecuta `npm start`.
5. Abre `http://localhost:5500` en la interfaz; la API debe estar en `http://localhost:3000`.

La API responde en `/api/*`; `/api/health` sirve para comprobar que está arriba. No sirve páginas HTML.

## Configuración web y sesión

- `FRONTEND_ORIGINS`: orígenes completos permitidos, separados por comas. En producción usa el dominio exacto de la interfaz, por ejemplo `https://asistencia.gruponerba.com`.
- `COOKIE_SAME_SITE`: `lax` para interfaz y API en subdominios HTTPS del mismo dominio registrable.
- `NODE_ENV=production`: activa la marca `Secure` de las cookies.
- `SUPERADMIN_SETUP_TOKEN`, `VAPID_PUBLIC_KEY`, `VAPID_PRIVATE_KEY` y `VAPID_SUBJECT`: secretos/configuración del servidor. No los subas al repositorio.

## Persistencia Supabase

La API puede seguir usando `db.json` en local o conectarse a Supabase cuando se configuren `SUPABASE_URL` y `SUPABASE_SECRET_KEY` (también acepta temporalmente `SUPABASE_SERVICE_ROLE_KEY`). Nunca agregues esa clave al frontend ni la subas a Git. El frontend continúa hablando con esta API por `api-client.js`.

La migración conserva el modelo actual y escribe cada elemento de las colecciones en `public.jjm_records` como JSONB. Es una capa de compatibilidad para llevar la app a una instancia remota y validar el flujo; la API carga las colecciones en memoria y está pensada para **una sola instancia**. No aumentes réplicas ni instancias hasta reemplazar este adaptador por consultas transaccionales y tablas normalizadas.

### Preparar Supabase

1. En Supabase abre **SQL Editor → New query**.
2. Copia y ejecuta `supabase/migrations/202611230001_create_jjm_records.sql`.
3. En la configuración del backend agrega `SUPABASE_URL` con la URL del proyecto y `SUPABASE_SECRET_KEY` con una clave secreta del backend. La clave antigua `service_role` también funciona, pero debe mantenerse privada.

### Migrar la base local

Hazlo solo si quieres copiar los datos del `db.json` local. En este equipo está en `../sistema-asistencia-web/db.json` respecto al repo del backend. Conserva un respaldo y primero ejecuta la vista previa; no necesita clave ni se conecta a Supabase:

```powershell
node scripts/migrar-db-local-a-supabase.js "..\sistema-asistencia-web\db.json"
```

La vista previa no escribe datos. Si el conteo y las colecciones son correctos, crea una clave secret en Supabase y corre el importador asistido desde la carpeta del backend. El comando pide la clave oculta, la usa solo durante la importación y la retira del entorno al terminar:

```powershell
.\scripts\importar-db-supabase.ps1
```

El script se detiene si Supabase ya contiene registros, para no sobrescribirlos accidentalmente. No pegues la clave en chats ni la guardes en el repositorio.

Cuando Supabase tenga los datos, inicia la API con esas variables configuradas. Si es una base nueva sin superadmin, define además `SUPERADMIN_SETUP_TOKEN` como secreto antes del primer arranque y usa ese token en la configuración inicial; elimínalo del host al terminar el bootstrap.

## Despliegue y límites actuales

Supabase aporta Postgres y APIs, pero no ejecuta este `server.js` como un proceso Express permanente. Este repo ya incluye el contenedor y el Worker de Cloudflare en `Dockerfile`, `src/index.mjs` y `wrangler.jsonc`.

### Publicar el backend en Cloudflare

Containers requiere el plan Workers Paid. En Cloudflare conecta este repo desde **Workers & Pages → Create application → Workers Builds** (no como Pages estático):

1. Repositorio: `ControldeHorasBackend`; root directory: `.`.
2. Build command: `npm ci`.
3. Deploy command: `npm run cf:deploy`.
4. Después del primer deploy, abre el Worker → **Settings → Variables and Secrets** y crea `SUPABASE_SECRET_KEY` como **secret**. `SUPABASE_URL` y `FRONTEND_ORIGINS` ya están declaradas en `wrangler.jsonc`; actualízalas ahí si cambia el dominio.
5. El proyecto enruta todo a una sola instancia de Container porque esta versión carga el estado en memoria. No aumentes `max_instances`.

El frontend ya apunta a `https://api.gruponerba.com`. Después del primer despliegue, asígnale ese custom domain al Worker y asigna `asistencia.gruponerba.com` a Pages. Mantener ambos bajo `gruponerba.com` permite que las cookies de sesión con `SameSite=Lax` funcionen como mismo sitio.

La imagen necesita salida a internet para conectar con Supabase. No subas `SUPABASE_SECRET_KEY` a `wrangler.jsonc`, GitHub, Pages ni al chat; el Worker la pasa al contenedor al iniciarlo.

Los archivos subidos y `data/horas-original.xlsx` todavía usan almacenamiento local; falta moverlos a Supabase Storage o a otro almacenamiento persistente antes de producción. La conexión de datos de esta fase no debe considerarse una arquitectura final para escalar comercialmente.
