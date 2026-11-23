-- Almacenamiento de compatibilidad para la API actual de JJM Clock.
-- La API conserva sus colecciones existentes, pero cada registro vive en Postgres.
-- Acceso exclusivo desde el backend; no exponer la clave secreta al frontend.

create table if not exists public.jjm_records (
  collection text not null check (
    collection in (
      'empresa',
      'usuarios',
      'areas',
      'horarios_especiales',
      'asistencia',
      'pendientes',
      'periodo_laboral',
      'vacaciones_solicitud',
      'sesiones',
      'superadmins',
      'superadmin_sesiones',
      'pagos_suscripcion',
      'push_suscripciones',
      'notificaciones_enviadas',
      'auditoria_superadmin'
    )
  ),
  id bigint not null check (id > 0),
  record jsonb not null check (
    jsonb_typeof(record) = 'object' and record ->> 'id' = id::text
  ),
  updated_at timestamptz not null default now(),
  primary key (collection, id)
);

create index if not exists jjm_records_company_idx
  on public.jjm_records (collection, (record ->> 'empresa_id'));

create index if not exists jjm_records_email_idx
  on public.jjm_records (collection, (lower(record ->> 'correo')))
  where record ? 'correo';

alter table public.jjm_records enable row level security;

-- Ningún cliente del navegador obtiene acceso directo a esta tabla.
revoke all on table public.jjm_records from public, anon, authenticated;
grant select, insert, update, delete on table public.jjm_records to service_role;

comment on table public.jjm_records is
  'Persistencia temporal compatible con el modelo de colecciones de JJM Clock; acceso solo desde el backend.';
