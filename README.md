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

## Estado de persistencia

La versión separada aún conserva el almacenamiento local `db.json` y `data/` para desarrollo. Ambos están excluidos de Git. **No desplegar como producción hasta migrar la persistencia y los archivos a Supabase**; los discos de los contenedores pueden ser efímeros.
