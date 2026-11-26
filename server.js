/*********************************************************************************
 * CONTROL DE ASISTENCIA — V8 WEB
 * Migración desde Google Apps Script a Node.js + Express + JSON (sin compilación)
 * - Login con código de empresa + correo + contraseña (sesión persistente en cookie)
 * - Entrada / Salida / Comida (entrada y salida)
 * - Horarios por área, sábado sin comida, domingo = incidencia
 * - Horarios especiales por persona y día
 * - Reportes semanales, pendientes y RH/Administrador
 * - Seguimiento de periodo laboral (días trabajados) para vacaciones
 * - Base de datos en archivo db.json (fácil de respaldar: copia el archivo)
 *********************************************************************************/

const express = require('express');
const cookieParser = require('cookie-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { promisify } = require('util');
const webpush = require('web-push');
const { SupabaseStore } = require('./supabase-store');
const scryptAsync = promisify(crypto.scrypt);

// =============================================
// CONFIGURACIÓN
// =============================================
const APP_NAME = "Control de Asistencia — Web";
const VERSION_CONTROL_ASISTENCIA = "V8 WEB";

const ORDEN_AREAS = ["Oficina", "Técnico", "Sin asignar"];

const HORARIOS = {
  Oficina: { inicio: "09:00", fin: "19:00" },
  Técnico: { inicio: "09:00", fin: "20:00" },
  "Sin asignar": { inicio: "09:00", fin: "19:00" }
};

const HORARIO_SABADO = { inicio: "09:00", fin: "15:00" };

const CORREO_BLANCA = "nerbapublicidad@gmail.com";
const HORARIO_BLANCA = { inicio: "08:30", fin: "19:00" };

const MINUTOS_COMIDA_POR_DEFECTO = 60;
const MINUTOS_MINIMOS_PARA_DESCONTAR_COMIDA = 300;
const SABADO_SIN_COMIDA = true;
const MINUTOS_MINIMOS_ENTRE_MARCAS = 5;
const HORA_MIN_REGISTRO = "06:00";
const HORA_MAX_REGISTRO = "23:30";

const TOTAL_COLUMNAS = 7;
const TOTAL_COLUMNAS_PERSONA = 10;

const NOMBRES_DIA = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const DIAS_LABORALES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

const NOMBRE_GRUPO = {
  Oficina: "OFICINA",
  Técnico: "TÉCNICOS",
  "Sin asignar": "SIN ÁREA ASIGNADA"
};

/*********************************************************************************
 * UBICACIÓN DE LA OFICINA (por empresa, opcional)
 * Cada empresa puede configurar su propia ubicación y radio permitido.
 * Si no hay configuración, no se valida ubicación al marcar.
 *********************************************************************************/

function obtenerOficinaConfig(empresaId) {
  const e = findOne("empresa", emp => emp.id === empresaId);
  if (!e) return null;
  if (e.habilitar_ubicacion === false) return null;
  if (e.lat === null || e.lat === undefined || e.lng === null || e.lng === undefined) return null;
  return {
    lat: Number(e.lat),
    lng: Number(e.lng),
    radioMetros: Number(e.radio_metros || 100),
    direccion: e.direccion || ""
  };
}

function distanciaHaversonMetros(lat1, lng1, lat2, lng2) {
  const R = 6371000; // radio de la Tierra en metros
  const rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad;
  const dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) *
    Math.sin(dLng / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

// =============================================
// PERSISTENCIA JSON (sin compilación nativa)
// =============================================
const DB_PATH = path.join(__dirname, "db.json");
const DIRECTORIO_DATOS = path.join(__dirname, "data");
const ARCHIVO_HORAS_ORIGINAL = path.join(DIRECTORIO_DATOS, "horas-original.xlsx");

const TABLAS = [
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

const SUPABASE_URL = String(process.env.SUPABASE_URL || "").trim();
const SUPABASE_KEY = String(
  process.env.SUPABASE_SECRET_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY || ""
).trim();
const USAR_SUPABASE = Boolean(SUPABASE_URL || SUPABASE_KEY);

if (USAR_SUPABASE && (!SUPABASE_URL || !SUPABASE_KEY)) {
  throw new Error("Para usar Supabase configura SUPABASE_URL y SUPABASE_SECRET_KEY.");
}

let db = {};
let nextIdCounter = {};
let supabaseStore = null;
let snapshotSupabase = null;
let colaPersistencia = Promise.resolve();
let ultimoErrorPersistencia = null;

function cargarDB() {
  if (fs.existsSync(DB_PATH)) {
    try {
      const contenido = fs.readFileSync(DB_PATH, "utf8");
      const datos = JSON.parse(contenido);
      db = datos.db || {};
      nextIdCounter = datos.nextId || {};
      // Asegurar que todas las tablas existan
      TABLAS.forEach(t => { if (!db[t]) db[t] = []; });
      return;
    } catch (e) {
      console.warn("No se pudo leer db.json, se inicia vacío:", e.message);
    }
  }
  resetDB();
}

function resetDB() {
  db = {};
  nextIdCounter = {};
  TABLAS.forEach(t => { db[t] = []; });
}

function copiarColecciones(origen = db) {
  return Object.fromEntries(TABLAS.map(tabla => [
    tabla,
    JSON.parse(JSON.stringify(origen[tabla] || []))
  ]));
}

function guardarDB() {
  if (USAR_SUPABASE) {
    if (!supabaseStore || !snapshotSupabase) {
      throw new Error("La conexión a Supabase todavía no terminó de inicializarse.");
    }
    colaPersistencia = colaPersistencia.then(async () => {
      const snapshot = copiarColecciones();
      await supabaseStore.persistChanges(snapshot, snapshotSupabase);
      snapshotSupabase = snapshot;
      ultimoErrorPersistencia = null;
    }).catch(error => {
      ultimoErrorPersistencia = error;
      console.error("No se pudo guardar en Supabase:", error.message);
    });
    return;
  }
  const payload = { db, nextId: nextIdCounter };
  fs.writeFileSync(DB_PATH, JSON.stringify(payload, null, 2), "utf8");
}

async function esperarPersistenciaDB() {
  if (!USAR_SUPABASE) return;
  await colaPersistencia;
  if (ultimoErrorPersistencia) {
    ultimoErrorPersistencia = null;
    guardarDB();
    await colaPersistencia;
    if (ultimoErrorPersistencia) {
      throw new Error("No se pudieron guardar los cambios en Supabase: " + ultimoErrorPersistencia.message);
    }
  }
}

function tieneDatosLocales() {
  if (!fs.existsSync(DB_PATH)) return false;
  try {
    const payload = JSON.parse(fs.readFileSync(DB_PATH, "utf8"));
    const datos = payload.db || payload;
    return TABLAS.some(tabla => Array.isArray(datos[tabla]) && datos[tabla].length > 0);
  } catch {
    return false;
  }
}

async function cargarDBSupabase() {
  supabaseStore = new SupabaseStore({
    url: SUPABASE_URL,
    key: SUPABASE_KEY,
    collections: TABLAS
  });
  const remoto = await supabaseStore.load();
  const tieneDatosRemotos = TABLAS.some(tabla => remoto.db[tabla].length > 0);
  if (!tieneDatosRemotos && tieneDatosLocales()) {
    throw new Error(
      "Supabase está vacío y db.json local contiene datos. Migra primero con " +
      "scripts/migrar-db-local-a-supabase.js --apply para evitar perderlos."
    );
  }
  db = remoto.db;
  nextIdCounter = remoto.nextId;
  snapshotSupabase = copiarColecciones(db);
  console.log("Persistencia Supabase cargada.");
}

function nextId(tabla) {
  if (USAR_SUPABASE) {
    let id = 0;
    do {
      id = crypto.randomBytes(6).readUIntBE(0, 6);
    } while (!id || (db[tabla] || []).some(fila => fila.id === id));
    return id;
  }
  if (!nextIdCounter[tabla]) nextIdCounter[tabla] = 1;
  return nextIdCounter[tabla]++;
}

function insert(tabla, obj) {
  obj.id = nextId(tabla);
  db[tabla].push(obj);
  guardarDB();
  return obj;
}

function update(tabla, id, cambios) {
  const fila = db[tabla].find(r => r.id === id);
  if (!fila) return false;
  Object.assign(fila, cambios);
  guardarDB();
  return true;
}

function remove(tabla, id) {
  const idx = db[tabla].findIndex(r => r.id === id);
  if (idx === -1) return false;
  db[tabla].splice(idx, 1);
  guardarDB();
  return true;
}

function find(tabla, filtroFn) {
  return db[tabla].filter(filtroFn);
}

function findOne(tabla, filtroFn) {
  return db[tabla].find(filtroFn);
}

resetDB();

// =============================================
// INICIALIZAR DATOS POR DEFECTO
// =============================================
function inicializarDatos() {
  let cambio = false;
  db.usuarios.forEach(u => {
    if (!u.rol) {
      u.rol = u.correo === "admin@empresa.local" ? "admin" : "empleado";
      cambio = true;
    }
  });
  if (cambio) guardarDB();
}

function crearAreasPredeterminadas(empresaId) {
  if (findOne("areas", a => a.empresa_id === empresaId)) return;
  Object.entries(HORARIOS).filter(([nombre]) => nombre !== "Sin asignar").forEach(([nombre, horario]) => {
    insert("areas", {
      empresa_id: empresaId, nombre, entrada: horario.inicio, salida: horario.fin,
      activo: 1, creado_en: new Date().toISOString()
    });
  });
}

function inicializarAreasExistentes() {
  let cambios = false;
  (db.empresa || []).forEach(empresa => {
    if (empresa.areas_configuradas === undefined) {
      empresa.areas_configuradas = true;
      cambios = true;
    }
    crearAreasPredeterminadas(empresa.id);
  });
  if (cambios) guardarDB();
}

function areasActivasEmpresa(empresaId) {
  return find("areas", a => a.empresa_id === empresaId && Number(a.activo) === 1);
}

function contarCuposOcupados(empresaId) {
  return find("usuarios", u => u.empresa_id === empresaId && u.rol !== "admin" && u.activo === 1).length;
}

function estadoSuscripcion(empresa) {
  if (!empresa) return "no_disponible";
  if (empresa.archivada_en || empresa.estado_suscripcion === "archivada") return "archivada";
  if (empresa.estado_suscripcion === "activa" && empresa.fecha_fin_suscripcion
      && empresa.fecha_fin_suscripcion < claveFecha(new Date())) return "vencida";
  return empresa.estado_suscripcion || (empresa.codigo ? "activa" : "pendiente");
}

function suscripcionActiva(empresa) {
  return estadoSuscripcion(empresa) === "activa" && Boolean(empresa?.codigo);
}

function inicializarSuscripcionesEquipos() {
  let cambios = false;
  (db.empresa || []).forEach(empresa => {
    if (!empresa.estado_suscripcion) {
      empresa.estado_suscripcion = empresa.codigo ? "activa" : "pendiente";
      cambios = true;
    }
    if (empresa.cupos_limite === undefined && empresa.codigo) {
      // Conserva el acceso de los equipos existentes y evita agregar miembros sin cupo asignado.
      empresa.cupos_limite = Math.max(10, contarCuposOcupados(empresa.id));
      cambios = true;
    }
    if (empresa.archivada === true && !empresa.archivada_en) {
      empresa.archivada_en = new Date().toISOString();
      empresa.estado_suscripcion = "archivada";
      empresa.codigo = null;
      cambios = true;
    }
  });
  if (cambios) guardarDB();
}

// Asociar la hoja de horas entregada con la empresa histórica EMPRESA01.
// El vínculo se guarda en el registro de esa empresa para no exponerla a equipos nuevos.
function asociarArchivoHorasOriginal() {
  if (!fs.existsSync(ARCHIVO_HORAS_ORIGINAL)) return;
  const empresaOriginal = findOne("empresa", e => e.codigo === "EMPRESA01");
  if (!empresaOriginal || empresaOriginal.archivo_horas_original) return;
  update("empresa", empresaOriginal.id, { archivo_horas_original: "data/horas-original.xlsx" });
  console.log(`Archivo de horas original vinculado al equipo histórico (${empresaOriginal.id}).`);
}

const CODIGO_EMPRESA_ALFABETO = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
const COOKIE_SESION = "asistencia_session";
const DIAS_SESION_MS = 30 * 24 * 60 * 60 * 1000;
const MIN_PASSWORD_LENGTH = 8;
const codigosConfiguracionAdmin = new Map();

function generarCodigoEmpresa() {
  const bytes = crypto.randomBytes(8);
  let codigo = "";
  for (const byte of bytes) codigo += CODIGO_EMPRESA_ALFABETO[byte % CODIGO_EMPRESA_ALFABETO.length];
  return `JJM-${codigo.slice(0, 4)}-${codigo.slice(4)}`;
}

function generarCodigoUnico() {
  let codigo;
  do { codigo = generarCodigoEmpresa(); }
  while (findOne("empresa", e => e.codigo === codigo));
  return codigo;
}

function normalizarCorreo(correo) {
  return String(correo || "").trim().toLowerCase();
}

function validarCuenta(nombre, correo, password) {
  if (!String(nombre || "").trim() || !normalizarCorreo(correo)) {
    return "Escribe el nombre y el correo.";
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizarCorreo(correo))) {
    return "Escribe un correo válido.";
  }
  if (typeof password !== "string" || password.length < MIN_PASSWORD_LENGTH) {
    return `La contraseña debe tener al menos ${MIN_PASSWORD_LENGTH} caracteres.`;
  }
  return null;
}

async function guardarPassword(password) {
  const salt = crypto.randomBytes(16);
  const hash = await scryptAsync(password, salt, 64);
  return `scrypt:${salt.toString("base64")}:${hash.toString("base64")}`;
}

async function verificarPassword(password, guardado) {
  const partes = String(guardado || "").split(":");
  if (partes.length !== 3 || partes[0] !== "scrypt") return false;
  try {
    const salt = Buffer.from(partes[1], "base64");
    const esperado = Buffer.from(partes[2], "base64");
    if (!salt.length || esperado.length !== 64) return false;
    const calculado = await scryptAsync(password, salt, esperado.length);
    return crypto.timingSafeEqual(esperado, calculado);
  } catch (e) {
    return false;
  }
}

function hashToken(token) {
  return crypto.createHash("sha256").update(token).digest("hex");
}

function prepararCodigosConfiguracionAdmin() {
  db.empresa.forEach(empresa => {
    const adminListo = findOne("usuarios", u =>
      u.empresa_id === empresa.id && u.rol === "admin" && Boolean(u.password_hash) && u.activo === 1
    );
    if (adminListo) return;
    const codigo = crypto.randomBytes(10).toString("hex").toUpperCase();
    codigosConfiguracionAdmin.set(empresa.id, codigo);
    console.log(`Código único para crear al administrador inicial de ${empresa.nombre}: ${codigo}`);
  });
}
// =============================================
// UTILIDADES (migradas de tu Apps Script)
// =============================================

function normalizar(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function aMinutos(textoHora) {
  if (textoHora === null || textoHora === undefined || textoHora === "") return null;
  if (typeof textoHora === "object" && textoHora instanceof Date) {
    const h = textoHora.getHours();
    const m = textoHora.getMinutes();
    return aMinutos(`${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`);
  }
  const partes = String(textoHora).trim().split(":");
  if (partes.length < 2) return null;
  const horas = Number(partes[0]);
  const minutos = Number(partes[1]);
  if (!Number.isFinite(horas) || !Number.isFinite(minutos) || horas < 0 || horas > 23 || minutos < 0 || minutos > 59) return null;
  return horas * 60 + minutos;
}

function aTextoHoras(minutos) {
  if (minutos === null || minutos === undefined || !isFinite(minutos)) return "0:00";
  const totalMinutos = Math.max(0, Math.round(Number(minutos)));
  const horas = Math.floor(totalMinutos / 60);
  const minutosEnteros = totalMinutos % 60;
  return horas + ":" + String(minutosEnteros).padStart(2, "0");
}

function obtenerZonaHoraria() {
  return process.env.TZ_ASISTENCIA || "America/Mexico_City";
}

function truncarSegundos(fecha) {
  const f = new Date(fecha.getTime());
  f.setSeconds(0, 0);
  return f;
}

function fechaDesdeEntrada(valor) {
  if (valor instanceof Date) return new Date(valor.getTime());
  const texto = String(valor || "").trim();
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(texto);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 12, 0, 0, 0);
  const visible = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(texto);
  if (visible) return new Date(Number(visible[3]), Number(visible[2]) - 1, Number(visible[1]), 12, 0, 0, 0);
  return new Date(texto);
}

function fechaISODesdeTexto(texto) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(texto || "").trim());
  if (!m) return null;
  const dia = Number(m[1]), mes = Number(m[2]), anio = Number(m[3]);
  const f = new Date(anio, mes - 1, dia, 12, 0, 0, 0);
  if (f.getFullYear() !== anio || f.getMonth() !== mes - 1 || f.getDate() !== dia) return null;
  return `${String(anio).padStart(4, "0")}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;
}

function normalizarFechaSoloDia(fecha) {
  if (!fecha) return null;
  const f = fechaDesdeEntrada(fecha);
  if (isNaN(f.getTime())) return null;
  const partes = new Intl.DateTimeFormat("en-US", {
    timeZone: obtenerZonaHoraria(), year: "numeric", month: "2-digit", day: "2-digit"
  }).formatToParts(f).reduce((salida, parte) => {
    if (parte.type !== "literal") salida[parte.type] = parte.value;
    return salida;
  }, {});
  return new Date(Number(partes.year), Number(partes.month) - 1, Number(partes.day), 12, 0, 0, 0);
}

function claveFecha(fecha) {
  if (typeof fecha === "string") {
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(fecha.trim());
    if (iso) {
      const validada = fechaDesdeEntrada(fecha);
      if (validada.getFullYear() === Number(iso[1]) && validada.getMonth() === Number(iso[2]) - 1 && validada.getDate() === Number(iso[3])) return fecha.trim();
      return "";
    }
  }
  const f = normalizarFechaSoloDia(fecha);
  if (!f) return "";
  return `${String(f.getFullYear()).padStart(4, "0")}-${String(f.getMonth() + 1).padStart(2, "0")}-${String(f.getDate()).padStart(2, "0")}`;
}

function fechaISOValida(valor) {
  const texto = String(valor || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(texto) && Boolean(claveFecha(texto));
}

function contarDiasSinDomingos(inicioISO, finISO) {
  if (!fechaISOValida(inicioISO) || !fechaISOValida(finISO) || inicioISO > finISO) return 0;
  const [ia, im, id] = inicioISO.split("-").map(Number);
  const [fa, fm, fd] = finISO.split("-").map(Number);
  const cursor = new Date(Date.UTC(ia, im - 1, id, 12));
  const fin = Date.UTC(fa, fm - 1, fd, 12);
  let total = 0;
  while (cursor.getTime() <= fin) {
    if (cursor.getUTCDay() !== 0) total++;
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return total;
}

function fechaTextoVisual(fecha) {
  const f = normalizarFechaSoloDia(fecha);
  if (!f) return "";
  return `${String(f.getDate()).padStart(2, "0")}/${String(f.getMonth() + 1).padStart(2, "0")}/${f.getFullYear()}`;
}

function repararFechasGuardadas() {
  let asistenciasCorregidas = 0;
  let pendientesCorregidos = 0;
  db.asistencia.forEach(fila => {
    const fechaCorrecta = fechaISODesdeTexto(fila.fecha_texto);
    if (!fechaCorrecta || fila.fecha === fechaCorrecta) return;
    fila.fecha = fechaCorrecta;
    fila.dia_semana = fechaDesdeEntrada(fechaCorrecta).getDay();
    asistenciasCorregidas++;
  });
  db.pendientes.forEach(fila => {
    const marca = String(fila.marca_texto || "").match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
    const titulo = String(fila.titulo_dia || "").match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
    const fechaTexto = marca ? marca[1] : (titulo ? titulo[1] : null);
    const fechaCorrecta = fechaISODesdeTexto(fechaTexto);
    if (!fechaCorrecta || fila.fecha === fechaCorrecta) return;
    fila.fecha = fechaCorrecta;
    if (fechaTexto) fila.titulo_dia = `ASISTENCIA — ${fechaTexto}`;
    pendientesCorregidos++;
  });
  if (asistenciasCorregidas || pendientesCorregidos) {
    guardarDB();
    console.log(`Fechas reparadas desde su texto original: ${asistenciasCorregidas} asistencias y ${pendientesCorregidos} pendientes.`);
  }
}

repararFechasGuardadas();

function claveUsuarioFecha(correo, fecha) {
  return String(correo || "").trim().toLowerCase() + "|" + claveFecha(fecha);
}

function esDomingoPorFecha(fecha) {
  if (!fecha) return false;
  if (fecha instanceof Date) return fecha.getDay() === 0;
  const f = normalizarFechaSoloDia(fecha);
  return f ? f.getDay() === 0 : false;
}

function esSabadoPorFecha(fecha) {
  if (!fecha) return false;
  if (fecha instanceof Date) return fecha.getDay() === 6;
  const f = normalizarFechaSoloDia(fecha);
  return f ? f.getDay() === 6 : false;
}

function normalizarMinutosComida(valor) {
  if (valor === "" || valor === null || valor === undefined) return MINUTOS_COMIDA_POR_DEFECTO;
  if (typeof valor === "number") {
    if (valor >= 0 && valor < 1) return valor * 24 * 60;
    return Math.max(0, valor);
  }
  const texto = String(valor).trim();
  const partes = texto.split(":");
  if (partes.length >= 2) {
    const h = Number(partes[0]);
    const m = Number(partes[1]);
    const ss = partes.length > 2 ? Number(partes[2]) : 0;
    if (!isNaN(h) && !isNaN(m) && !isNaN(ss)) return Math.max(0, h * 60 + m + ss / 60);
  }
  const n = Number(texto);
  return isNaN(n) ? MINUTOS_COMIDA_POR_DEFECTO : Math.max(0, n);
}

function obtenerMinutosComidaUsuario(usuario) {
  return usuario && usuario.minutos_comida !== undefined
    ? normalizarMinutosComida(usuario.minutos_comida)
    : MINUTOS_COMIDA_POR_DEFECTO;
}

function normalizarHoraManual(valor) {
  if (valor === null || valor === undefined || valor === "") return "";
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return `${String(valor.getHours()).padStart(2,'0')}:${String(valor.getMinutes()).padStart(2,'0')}`;
  }
  if (typeof valor === "number" && !isNaN(valor) && valor >= 0 && valor < 1) {
    const totalMinutos = Math.round(valor * 24 * 60);
    const horas = Math.floor(totalMinutos / 60) % 24;
    const minutos = totalMinutos % 60;
    return String(horas).padStart(2, '0') + ":" + String(minutos).padStart(2, '0');
  }
  const texto = String(valor).trim().toLowerCase();
  let m = /^(\d{1,2}):(\d{2})(?::\d{2})?$/.exec(texto);
  if (m) {
    const h = Number(m[1]);
    const min = Number(m[2]);
    if (h >= 0 && h <= 23 && min >= 0 && min <= 59) return String(h).padStart(2, '0') + ":" + String(min).padStart(2, '0');
    return null;
  }
  m = /^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$/.exec(texto);
  if (m) {
    let h = Number(m[1]);
    const min = Number(m[2] || 0);
    const ampm = m[3];
    if (h < 1 || h > 12 || min < 0 || min > 59) return null;
    if (ampm === "pm" && h !== 12) h += 12;
    if (ampm === "am" && h === 12) h = 0;
    return String(h).padStart(2, '0') + ":" + String(min).padStart(2, '0');
  }
  return null;
}

function horaPermitida(hora) {
  if (!HORA_MIN_REGISTRO || !HORA_MAX_REGISTRO) return true;
  const m = aMinutos(hora);
  return m !== null && m >= aMinutos(HORA_MIN_REGISTRO) && m <= aMinutos(HORA_MAX_REGISTRO);
}

/*********************************************************************************
 * HORARIOS ESPECIALES (migrado)
 *********************************************************************************/

let HORARIOS_ESP_CACHE = null;
let HORARIOS_ESP_CACHE_EMPRESA = null;

function limpiarCacheHorarios(empresaId) {
  if (empresaId !== undefined) {
    HORARIOS_ESP_CACHE = null;
    HORARIOS_ESP_CACHE_EMPRESA = empresaId;
  } else {
    HORARIOS_ESP_CACHE = null;
    HORARIOS_ESP_CACHE_EMPRESA = null;
  }
}

function cargarHorariosEspeciales(empresaId) {
  if (HORARIOS_ESP_CACHE !== null && HORARIOS_ESP_CACHE_EMPRESA === empresaId) {
    return HORARIOS_ESP_CACHE;
  }

  const reglas = find("horarios_especiales", r => r.empresa_id === empresaId && r.activo === 1);

  const resultado = [];
  reglas.forEach(r => {
    const entrada = normalizarHoraManual(r.entrada);
    const salida = normalizarHoraManual(r.salida);
    if (!entrada || !salida) return;
    resultado.push({
      correo: String(r.correo || "").trim().toLowerCase(),
      dia: normalizar(r.dia || ""),
      inicio: entrada,
      fin: salida,
      desde: r.vigente_desde || "0000-00-00",
      hasta: r.vigente_hasta || null
    });
  });

  HORARIOS_ESP_CACHE = resultado;
  HORARIOS_ESP_CACHE_EMPRESA = empresaId;
  return resultado;
}

function obtenerHorarioEspecial(correo, fecha, empresaId) {
  if (!(fecha instanceof Date)) fecha = normalizarFechaSoloDia(fecha);
  if (!fecha) return null;

  const clave = String(correo || "").trim().toLowerCase();
  const dia = normalizar(NOMBRES_DIA[fecha.getDay()]);
  const fk = claveFecha(fecha);

  let mejor = null;
  cargarHorariosEspeciales(empresaId).forEach(r => {
    if (r.correo !== clave) return;
    if (r.dia !== dia) return;
    if (fk < r.desde) return;
    if (r.hasta && fk > r.hasta) return;
    if (!mejor || r.desde > mejor.desde) mejor = r;
  });

  return mejor ? { inicio: mejor.inicio, fin: mejor.fin } : null;
}

/*********************************************************************************
 * HORARIO PARA FECHA
 * Prioridad:
 * 1) Horario especial vigente
 * 2) Sábado general (09:00 a 15:00)
 * 3) Excepción de Blanca
 * 4) Horario del área
 *********************************************************************************/

function obtenerHorarioParaFecha(area, fecha, correoUsuario, empresaId) {
  const correo = String(correoUsuario || "").trim().toLowerCase();

  const especial = obtenerHorarioEspecial(correo, fecha, empresaId);
  if (especial) return especial;

  if (esSabadoPorFecha(fecha)) {
    return HORARIO_SABADO;
  }

  if (normalizar(correo) === normalizar(CORREO_BLANCA)) {
    return HORARIO_BLANCA;
  }

  const areaConfigurada = findOne("areas", a => a.empresa_id === empresaId && normalizar(a.nombre) === normalizar(area || ""));
  if (areaConfigurada && areaConfigurada.entrada && areaConfigurada.salida) {
    return { inicio: areaConfigurada.entrada, fin: areaConfigurada.salida };
  }

  return HORARIOS[area] || HORARIOS["Sin asignar"];
}

function textoHorarioSabado(correo, fecha, empresaId) {
  const h = obtenerHorarioParaFecha(null, fecha, correo, empresaId);
  return h.inicio + " a " + h.fin;
}

/*********************************************************************************
 * CÁLCULO DE HORAS
 *********************************************************************************/

function calcularHoras(entrada, salida, area, fecha, minutosComidaUsuario, correoUsuario, empresaId) {
  const horario = obtenerHorarioParaFecha(area, fecha, correoUsuario, empresaId);

  const minutosEntrada = aMinutos(entrada);
  let minutosSalida = aMinutos(salida);

  if (minutosEntrada === null || minutosSalida === null) {
    return {
      brutas: "", comida: "", trabajadas: "", extra: "",
      minutosBrutas: 0, minutosComida: 0, minutosTrabajadas: 0, minutosExtra: 0
    };
  }

  if (minutosSalida < minutosEntrada) minutosSalida = minutosEntrada;

  const brutas = minutosSalida - minutosEntrada;
  const esSabado = esSabadoPorFecha(fecha);

  const comidaConfigurada = Math.max(0, normalizarMinutosComida(minutosComidaUsuario));
  const aplicaComida = !(esSabado && SABADO_SIN_COMIDA) && brutas >= MINUTOS_MINIMOS_PARA_DESCONTAR_COMIDA;
  const comida = aplicaComida ? Math.min(comidaConfigurada, brutas) : 0;
  const netas = Math.max(0, brutas - comida);

  const jornadaBruta = Math.max(0, aMinutos(horario.fin) - aMinutos(horario.inicio));
  const comidaJornada = aplicaComida ? Math.min(comidaConfigurada, jornadaBruta) : 0;
  const jornadaNeta = Math.max(0, jornadaBruta - comidaJornada);
  const extra = Math.max(0, netas - jornadaNeta);

  return {
    brutas: aTextoHoras(brutas),
    comida: aTextoHoras(comida),
    trabajadas: aTextoHoras(netas),
    extra: aTextoHoras(extra),
    minutosBrutas: brutas,
    minutosComida: comida,
    minutosTrabajadas: netas,
    minutosExtra: extra
  };
}

function registrarSalida(area, entrada, horaSalida, fechaTexto, minutosComidaUsuario, correoUsuario, empresaId) {
  const fecha = normalizarFechaSoloDia(fechaTexto);
  const calculoBase = calcularHoras(entrada, horaSalida, area, fecha, minutosComidaUsuario, correoUsuario, empresaId);
  const esDomingo = esDomingoPorFecha(fechaTexto);
  const esSabado = esSabadoPorFecha(fechaTexto);
  const calculo = esDomingo ? {
    ...calculoBase,
    brutas: "0:00", comida: "0:00", trabajadas: "0:00", extra: "0:00",
    minutosBrutas: 0, minutosComida: 0, minutosTrabajadas: 0, minutosExtra: 0
  } : calculoBase;

  let nota;
  if (esDomingo) {
    nota = "Registro realizado en domingo; no se contabiliza en horas ni horas extra.\nEste registro queda únicamente como incidencia y no afecta ningún total semanal.";
  } else {
    nota = "En bruto: " + calculo.brutas + "  ·  Comida descontada: " + calculo.comida;
    if (esSabado) {
      nota += "\nSábado: jornada de " + textoHorarioSabado(correoUsuario, fechaTexto, empresaId) + ".\nSábado: NO se descuenta hora de comida.";
      if (calculo.minutosExtra > 0) {
        nota += "\nHoras extra del sábado: " + calculo.extra + ". Se contabilizan para la semana siguiente.";
      }
    }
  }

  return {
    ...calculo,
    esDomingo,
    esSabado,
    estado: esDomingo ? "Incidencia: domingo" : (calculo.minutosExtra > 0 ? "Completado con horas extra" : "Completado"),
    nota
  };
}

/*********************************************************************************
 * FECHAS SEMANA (lunes a sábado)
// *********************************************************************************/

function lunesDeLaSemana(fecha) {
  const d = new Date(fecha.getTime());
  d.setHours(12, 0, 0, 0);
  const dia = d.getDay();
  const ajuste = (dia === 0) ? -6 : (1 - dia);
  d.setDate(d.getDate() + ajuste);
  return d;
}

function tituloSemana(fecha) {
  const lunes = lunesDeLaSemana(fecha);
  const sabado = new Date(lunes.getTime());
  sabado.setDate(sabado.getDate() + 5);
  return `SEMANA DEL ${claveFecha(lunes)} AL ${claveFecha(sabado)}`;
}

/*********************************************************************************
 * MIDDLEWARE
 *********************************************************************************/

const app = express();
function normalizarOrigenFrontend(origen) {
  try {
    return new URL(String(origen || '').trim()).origin;
  } catch {
    return '';
  }
}

const origenesFrontend = new Set(
  String(process.env.FRONTEND_ORIGINS || process.env.FRONTEND_ORIGIN || 'http://localhost:5500,http://127.0.0.1:5500')
    .split(',').map(normalizarOrigenFrontend).filter(Boolean)
);
// Dominio canónico donde se conserva correctamente la sesión del admin.
origenesFrontend.add('https://asistencia.gruponerba.com');
const cookieSameSite = String(process.env.COOKIE_SAME_SITE || 'lax').toLowerCase();
const corsOptions = {
  origin(origen, callback) {
    callback(null, !origen || origenesFrontend.has(origen));
  },
  credentials: true,
  methods: ['GET', 'HEAD', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  maxAge: 600
};
app.use((req, res, next) => {
  const origen = req.get('Origin');
  if (origen && !origenesFrontend.has(origen)) {
    return res.status(403).json({ ok: false, error: 'Origen no autorizado.' });
  }
  next();
});
app.use(cors(corsOptions));
app.use(express.json({ limit: "1mb" }));
app.use(cookieParser());
app.use((req, res, next) => {
  if (!USAR_SUPABASE) return next();
  const enviar = res.send.bind(res);
  let esperandoGuardado = false;
  res.send = function (...args) {
    if (esperandoGuardado) return res;
    esperandoGuardado = true;
    esperarPersistenciaDB().then(() => enviar(...args)).catch(error => {
      console.error("Se rechazó una respuesta porque no se guardaron los cambios:", error.message);
      if (res.headersSent) {
        res.destroy(error);
        return;
      }
      res.statusCode = 503;
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.setHeader("Cache-Control", "no-store");
      enviar(JSON.stringify({ ok: false, error: "No se pudo guardar la información. Inténtalo de nuevo." }));
    });
    return res;
  };
  next();
});
app.get("/", (req, res) => res.json({ ok: true, app: APP_NAME, api: "/api/health" }));

// Helper: obtener empresa y usuario desde cookie de sesión
// Cookie: token aleatorio; la sesión y el rol se resuelven desde el servidor.
function obtenerUsuarioSesion(req) {
  try {
    const token = req.cookies[COOKIE_SESION];
    if (!token) return null;
    const sesion = findOne("sesiones", s => s.token_hash === hashToken(token));
    if (!sesion || new Date(sesion.expira_en).getTime() <= Date.now()) return null;
    const usuario = findOne("usuarios", u => u.id === sesion.usuario_id && u.empresa_id === sesion.empresa_id && u.activo === 1);
    if (!usuario) return null;
    const empresa = findOne("empresa", e => e.id === sesion.empresa_id);
    if (!empresa || estadoSuscripcion(empresa) === "archivada") return null;
    return {
      id: usuario.id, empresa_id: usuario.empresa_id, correo: usuario.correo,
      nombre: usuario.nombre, area: usuario.area || "Sin asignar",
      minutos_comida: usuario.minutos_comida || MINUTOS_COMIDA_POR_DEFECTO,
      rol: usuario.rol || "empleado"
    };
  } catch (e) {
    return null;
  }
}

function requireAuth(req, res, next) {
  const u = obtenerUsuarioSesion(req);
  if (!u) return res.status(401).json({ ok: false, error: "No autenticado" });
  req.usuario = u;
  next();
}

function requireAdmin(req, res, next) {
  const u = obtenerUsuarioSesion(req);
  if (!u) return res.status(401).json({ ok: false, error: "No autenticado" });
  if (u.rol !== "admin") return res.status(403).json({ ok: false, error: "Acceso de administrador requerido" });
  req.usuario = u;
  next();
}

function crearSesion(res, usuario) {
  const token = crypto.randomBytes(32).toString("base64url");
  const expira = new Date(Date.now() + DIAS_SESION_MS);
  insert("sesiones", {
    token_hash: hashToken(token), usuario_id: usuario.id,
    empresa_id: usuario.empresa_id, expira_en: expira.toISOString()
  });
  res.cookie(COOKIE_SESION, token, {
    httpOnly: true, secure: process.env.NODE_ENV === "production",
    sameSite: cookieSameSite, path: "/", maxAge: DIAS_SESION_MS
  });
}

function usuarioPublico(usuario, empresa) {
  const ocupados = contarCuposOcupados(empresa.id);
  return {
    id: usuario.id, empresa_id: empresa.id,
    codigo_empresa: suscripcionActiva(empresa) ? empresa.codigo : null, empresa_nombre: empresa.nombre,
    correo: usuario.correo, nombre: usuario.nombre,
    area: usuario.area || "Sin asignar",
    minutos_comida: usuario.minutos_comida || MINUTOS_COMIDA_POR_DEFECTO,
    rol: usuario.rol || "empleado",
    areas_configuradas: empresa.areas_configuradas !== false,
    estado_suscripcion: estadoSuscripcion(empresa),
    fecha_fin_suscripcion: empresa.fecha_fin_suscripcion || null,
    cupos_limite: Number(empresa.cupos_limite || 0),
    cupos_ocupados: ocupados,
    monto_ciclo_mxn: empresa.monto_ciclo_mxn ?? null,
    periodicidad: empresa.periodicidad || "mensual"
  };
}

function usuarioSeUnioConCodigo(usuario) {
  if (!usuario || !usuario.password_hash) return false;
  if (usuario._registro_codigo === true) return true;
  if (usuario._alta_compartida === true) return true;
  // Compatibilidad con cuentas anteriores: los empleados ya activados se consideran miembros.
  return usuario.rol !== "admin";
}

function cerrarSesionActual(req, res) {
  const token = req.cookies[COOKIE_SESION];
  if (token) {
    const sesion = findOne("sesiones", s => s.token_hash === hashToken(token));
    if (sesion) remove("sesiones", sesion.id);
  }
  res.clearCookie(COOKIE_SESION, { path: "/", sameSite: cookieSameSite, secure: process.env.NODE_ENV === "production" });
  res.json({ ok: true });
}

// =============================================
// RUTAS PÚBLICAS
// =============================================

app.get("/api/health", (req, res) => {
  res.json({ ok: true, app: APP_NAME, version: VERSION_CONTROL_ASISTENCIA });
});

// Crear empresa y cuenta de administrador propietaria.
app.post("/api/registro/empresa", async (req, res) => {
  const { empresa: nombreEmpresa, nombre, correo, password } = req.body || {};
  const error = validarCuenta(nombre, correo, password);
  if (error) return res.status(400).json({ ok: false, error });
  if (!String(nombreEmpresa || "").trim()) return res.status(400).json({ ok: false, error: "Escribe el nombre de la empresa." });

  const empresa = insert("empresa", {
    codigo: null, nombre: String(nombreEmpresa).trim(),
    creado_en: new Date().toISOString(), lat: null, lng: null,
    radio_metros: null, direccion: null, habilitar_ubicacion: false,
    areas_configuradas: false,
    estado_suscripcion: "pendiente",
    solicitud_codigo_estado: "sin_solicitud",
    cupos_limite: 0,
    archivada_en: null
  });
  crearAreasPredeterminadas(empresa.id);
  const usuario = insert("usuarios", {
    empresa_id: empresa.id, correo: normalizarCorreo(correo), nombre: String(nombre).trim(),
    area: "Oficina", rol: "admin", activo: 1,
    minutos_comida: MINUTOS_COMIDA_POR_DEFECTO,
    fecha_ingreso: null, fecha_alta: new Date().toISOString(),
    password_hash: await guardarPassword(password)
  });
  crearSesion(res, usuario);
  notificarSuperadmins({
    clave: `equipo-nuevo-${empresa.id}`,
    titulo: "Nuevo equipo creado",
    cuerpo: `${empresa.nombre} ya aparece en la bandeja para revisar su solicitud.`,
    url: "/superadmin.html"
  });
  res.json({ ok: true, usuario: usuarioPublico(usuario, empresa) });
});

// Reclamar de forma segura una empresa de una versión anterior.
app.post("/api/registro/admin-inicial", async (req, res) => {
  const { codigo_empresa, codigo_configuracion, nombre, correo, password } = req.body || {};
  const error = validarCuenta(nombre, correo, password);
  if (error) return res.status(400).json({ ok: false, error });
  const empresa = findOne("empresa", e => e.codigo === String(codigo_empresa || "").trim().toUpperCase());
  if (!empresa) return res.status(401).json({ ok: false, error: "Código de empresa no válido." });
  const esperado = codigosConfiguracionAdmin.get(empresa.id);
  const recibido = String(codigo_configuracion || "").trim().toUpperCase();
  if (!esperado || recibido.length !== esperado.length || !crypto.timingSafeEqual(Buffer.from(recibido), Buffer.from(esperado))) {
    return res.status(401).json({ ok: false, error: "El código temporal de configuración no es válido. Revisa la consola del servidor." });
  }
  const correoNormalizado = normalizarCorreo(correo);
  const usuarioExistente = findOne("usuarios", u => u.empresa_id === empresa.id && u.correo === correoNormalizado);
  if (usuarioExistente && usuarioExistente.password_hash) {
    return res.status(400).json({ ok: false, error: "Ese correo ya está registrado en el equipo." });
  }

  const codigoAnterior = empresa.codigo;
  empresa.codigo = generarCodigoUnico();
  guardarDB();
  const credenciales = {
    nombre: String(nombre).trim(), area: "Oficina", rol: "admin", activo: 1,
    fecha_ingreso: usuarioExistente?.fecha_ingreso || null,
    fecha_alta: usuarioExistente?.fecha_alta || new Date().toISOString(),
    password_hash: await guardarPassword(password)
  };
  let usuario;
  if (usuarioExistente) {
    update("usuarios", usuarioExistente.id, credenciales);
    usuario = findOne("usuarios", u => u.id === usuarioExistente.id);
  } else {
    usuario = insert("usuarios", { empresa_id: empresa.id, correo: correoNormalizado, minutos_comida: MINUTOS_COMIDA_POR_DEFECTO, ...credenciales });
  }
  const administradorViejo = findOne("usuarios", u => u.empresa_id === empresa.id && u.correo === "admin@empresa.local" && !u.password_hash && u.id !== usuario.id);
  if (administradorViejo) update("usuarios", administradorViejo.id, { activo: 0, rol: "empleado" });
  codigosConfiguracionAdmin.delete(empresa.id);
  crearSesion(res, usuario);
  res.json({ ok: true, usuario: usuarioPublico(usuario, empresa), codigo_equipo: empresa.codigo, codigo_anterior: codigoAnterior });
});

// Alta de un empleado usando el código de equipo compartido por RH.
app.post("/api/registro/equipo", async (req, res) => {
  const { codigo_empresa, nombre, correo, password } = req.body || {};
  const error = validarCuenta(nombre, correo, password);
  if (error) return res.status(400).json({ ok: false, error });
  const empresa = findOne("empresa", e => e.codigo === String(codigo_empresa || "").trim().toUpperCase());
  if (!empresa) return res.status(401).json({ ok: false, error: "El código compartido no corresponde a un equipo." });
  if (!suscripcionActiva(empresa)) {
    return res.status(403).json({ ok: false, error: "No puedes ingresar al equipo. Acércate con tu encargad@ para solucionar eso." });
  }
  const correoNormalizado = normalizarCorreo(correo);
  const existente = findOne("usuarios", u => u.empresa_id === empresa.id && u.correo === correoNormalizado);
  if (existente && (existente.password_hash || (existente.rol === "admin" && !existente._alta_compartida))) {
    return res.status(409).json({ ok: false, error: "Ese correo ya tiene cuenta o requiere activación del administrador." });
  }
  const cupoYaOcupado = Boolean(existente && existente.activo === 1 && existente.rol !== "admin");
  if (!cupoYaOcupado && Number(empresa.cupos_limite || 0) <= contarCuposOcupados(empresa.id)) {
    return res.status(409).json({ ok: false, error: "No puedes ingresar al equipo, acércate con tu encargad@ para solucionar eso. El equipo llegó al límite de cupos." });
  }

  const password_hash = await guardarPassword(password);
  let usuario;
  if (existente) {
    const rolDestino = existente.rol === "admin" && existente._alta_compartida ? "admin" : "empleado";
    update("usuarios", existente.id, {
      nombre: String(nombre).trim(), password_hash, activo: 1, rol: rolDestino,
      ...(rolDestino === "admin" ? { area: "Oficina" } : {}),
      _alta_compartida: false,
      _registro_codigo: true,
      fecha_alta: existente.fecha_alta || new Date().toISOString()
    });
    usuario = findOne("usuarios", u => u.id === existente.id);
  } else {
    usuario = insert("usuarios", {
      empresa_id: empresa.id, correo: correoNormalizado, nombre: String(nombre).trim(),
      area: "Sin asignar", rol: "empleado", activo: 1,
      minutos_comida: MINUTOS_COMIDA_POR_DEFECTO,
      fecha_ingreso: null, fecha_alta: new Date().toISOString(), password_hash,
      _registro_codigo: true
    });
  }
  crearSesion(res, usuario);
  notificarAdministradoresCupo(empresa.id);
  res.json({ ok: true, usuario: usuarioPublico(usuario, empresa) });
});

// Iniciar sesión con código de equipo, correo y contraseña.
app.post("/api/login", async (req, res) => {
  const { codigo_empresa, correo, password } = req.body || {};
  const correoNormalizado = normalizarCorreo(correo);
  if (!correoNormalizado || !password) {
    return res.status(400).json({ ok: false, error: "Escribe tu correo y contraseña." });
  }

  const codigoNormalizado = String(codigo_empresa || "").trim().toUpperCase();
  let candidatos;
  if (codigoNormalizado) {
    const empresaCodigo = findOne("empresa", e => e.codigo === codigoNormalizado);
    if (!empresaCodigo) return res.status(401).json({ ok: false, error: "Código de equipo incorrecto." });
    candidatos = find("usuarios", u => u.empresa_id === empresaCodigo.id && u.correo === correoNormalizado && u.activo === 1);
  } else {
    candidatos = find("usuarios", u => u.correo === correoNormalizado && u.activo === 1);
  }

  const coincidencias = [];
  for (const candidato of candidatos) {
    if (candidato.password_hash && await verificarPassword(String(password), candidato.password_hash)) {
      coincidencias.push(candidato);
    }
  }
  if (coincidencias.length > 1) {
    return res.status(409).json({ ok: false, requiere_codigo: true, error: "Tu correo pertenece a más de un equipo. Ingresa el código de tu equipo." });
  }
  const usuario = coincidencias[0];
  if (!usuario) return res.status(401).json({ ok: false, error: "Correo o contraseña incorrectos. Si aún no tienes cuenta, únete con el código de tu equipo." });
  const empresa = findOne("empresa", e => e.id === usuario.empresa_id);
  if (!empresa) return res.status(401).json({ ok: false, error: "No se encontró el equipo de esta cuenta." });
  if (estadoSuscripcion(empresa) === "archivada") return res.status(403).json({ ok: false, error: "Este equipo ya no está disponible." });

  crearSesion(res, usuario);
  res.json({ ok: true, usuario: usuarioPublico(usuario, empresa) });
});

// Logout
app.post("/api/logout", cerrarSesionActual);

// Perfil actual
app.get("/api/perfil", requireAuth, (req, res) => {
  const u = req.usuario;
  const empresa = findOne("empresa", e => e.id === u.empresa_id);
  if (!empresa) return res.status(401).json({ ok: false, error: "Empresa no encontrada." });
  const usuario = findOne("usuarios", x => x.id === u.id);
  res.json({ ok: true, usuario: usuarioPublico(usuario, empresa) });
});

app.get("/api/asistencia/hoy", requireAuth, (req, res) => {
  const usuario = req.usuario;
  const fecha = claveFecha(new Date());
  const registro = findOne("asistencia", a =>
    a.empresa_id === usuario.empresa_id && a.correo === usuario.correo && a.fecha === fecha
  ) || null;
  const fechaDia = fechaDesdeEntrada(fecha);
  res.json({ ok: true, fecha, dia_semana: fechaDia.getDay(), registro });
});

// =============================================
// MARCAR ASISTENCIA (flujo web sin formulario)
// =============================================

function obtenerFechaHoyTexto() {
  return claveFecha(new Date());
}

function obtenerFechaHoyFecha() {
  return normalizarFechaSoloDia(new Date());
}

app.post("/api/asistencia/marcar", requireAuth, (req, res) => {
  const { tipo, lat, lng } = req.body; // entrada, salida, comida_entrada, comida_salida
  const usuario = req.usuario;
  const empresaId = usuario.empresa_id;
  const correo = usuario.correo;
  const nombre = usuario.nombre;
  const area = usuario.area;
  const empresaActual = findOne("empresa", e => e.id === empresaId);
  if (!suscripcionActiva(empresaActual)) {
    return res.status(403).json({ ok: false, error: "El equipo aún no está activo o su suscripción venció. Contacta al administrador." });
  }

  const ahora = new Date();
  const fecha = normalizarFechaSoloDia(ahora);
  const fechaTexto = claveFecha(fecha); // yyyy-MM-dd para BD
  const fechaTextoMostrar = fechaTextoVisual(ahora);
  const hora = new Intl.DateTimeFormat("en-GB", {
    timeZone: obtenerZonaHoraria(), hour: "2-digit", minute: "2-digit", hourCycle: "h23"
  }).format(ahora);
  const diaSemana = fecha.getDay();

  const latitud = lat !== null && lat !== undefined && lat !== "" ? Number(lat) : null;
  const longitud = lng !== null && lng !== undefined && lng !== "" ? Number(lng) : null;

  // El administrador puede marcar desde cualquier lugar (gestiona y supervisa);
  // los empleados sí deben estar dentro del radio autorizado si la empresa lo configuró.
  const esAdmin = String(usuario.rol || "") === "admin";
  const oficina = obtenerOficinaConfig(empresaId);
  if (oficina && !esAdmin) {
    if (!Number.isFinite(latitud) || !Number.isFinite(longitud) || latitud < -90 || latitud > 90 || longitud < -180 || longitud > 180) {
      return res.status(400).json({ ok: false, error: "Activa la ubicación para registrar asistencia." });
    }
    const distancia = distanciaHaversonMetros(latitud, longitud, oficina.lat, oficina.lng);
    if (distancia > oficina.radioMetros) {
      return res.status(403).json({ ok: false, error: `Estás fuera del rango autorizado (${oficina.radioMetros} m).` });
    }
  }

  if (!horaPermitida(hora)) {
    return res.status(400).json({ ok: false, error: `Hora fuera de rango permitido (${HORA_MIN_REGISTRO} - ${HORA_MAX_REGISTRO})` });
  }

  // El domingo queda como incidencia para empleados; el administrador puede registrar todas sus marcas.
  if (diaSemana === 0 && !esAdmin) {
    const existente = findOne(
      "asistencia",
      a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fechaTexto
    );

    if (existente) {
      return res.json({ ok: true, tipo: "domingo", mensaje: "Ya hay registro de domingo (incidencia). No se procesa cierre automático.", data: existente });
    }

    const info = {
      entrada: hora, salida: "",
      horas_brutas: "0:00", comida: "0:00", horas_trabajadas: "0:00", horas_extra: "0:00",
      estado: "Incidencia: domingo",
      observaciones: "Registro realizado en domingo; no se contabiliza en horas ni horas extra.",
      origen: "Web"
    };

    insert("asistencia", {
      empresa_id: empresaId, correo, nombre, area, fecha: fechaTexto, fecha_texto: fechaTextoMostrar,
      dia_semana: diaSemana, entrada: info.entrada, salida: info.salida,
      horas_brutas: info.horas_brutas, comida: info.comida, horas_trabajadas: info.horas_trabajadas,
      horas_extra: info.horas_extra, estado: info.estado, observaciones: info.observaciones, origen: info.origen,
      latitud, longitud
    });

    actualizarPendienteWeb(empresaId, correo, nombre, area, fechaTexto, fechaTextoMostrar, hora);
    actualizarPeriodoLaboral(empresaId, correo, fecha);

    return res.json({ ok: true, tipo: "domingo", mensaje: "Registro de domingo como incidencia (no suma horas).", data: info });
  }

  switch (tipo) {
    case "entrada": {
      const existente = findOne(
        "asistencia",
        a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fechaTexto
      );

      if (existente && existente.entrada) {
        return res.status(400).json({ ok: false, error: "Ya tienes una entrada registrada hoy." });
      }

      const esDomingo = esDomingoPorFecha(fechaTexto);
      const estado = esDomingo && !esAdmin ? "Incidencia: domingo" : "Entrada registrada — salida pendiente";

      if (existente) {
        const obs = (existente.observaciones || "") + " [Entrada registrada desde web]";
        update("asistencia", existente.id, {
          entrada: hora, salida: "", horas_brutas: "", comida: "", horas_trabajadas: "", horas_extra: "",
          comida_entrada: "", comida_salida: "", estado,
          observaciones: obs,
          fecha_texto: fechaTextoMostrar, dia_semana: diaSemana, latitud, longitud
        });
        actualizarPendienteWeb(empresaId, correo, nombre, area, fechaTexto, fechaTextoMostrar, hora, existente.id);
        actualizarPeriodoLaboral(empresaId, correo, fecha);
        return res.json({ ok: true, tipo: "entrada", mensaje: "Entrada actualizada.", data: { id: existente.id, entrada: hora } });
      }

      insert("asistencia", {
        empresa_id: empresaId, correo, nombre, area, fecha: fechaTexto, fecha_texto: fechaTextoMostrar,
        dia_semana: diaSemana, entrada: hora, salida: "",
        horas_brutas: "", comida: "", horas_trabajadas: "", horas_extra: "",
        estado, observaciones: "Entrada registrada desde web. Salida pendiente.", origen: "Web",
        latitud, longitud
      });

      const nueva = findOne(
        "asistencia",
        a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fechaTexto
      );

      actualizarPendienteWeb(empresaId, correo, nombre, area, fechaTexto, fechaTextoMostrar, hora, nueva.id);
      actualizarPeriodoLaboral(empresaId, correo, fecha);

      return res.json({ ok: true, tipo: "entrada", mensaje: "Entrada registrada.", data: { id: nueva.id, entrada: hora } });
    }

    case "salida": {
      const existente = findOne(
        "asistencia",
        a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fechaTexto
      );

      if (!existente || !existente.entrada) {
        return res.status(400).json({ ok: false, error: "No hay entrada registrada hoy para cerrar." });
      }
      if (existente.salida) {
        return res.status(400).json({ ok: false, error: "Ya tienes una salida registrada hoy." });
      }

      const minutosComida = obtenerMinutosComidaUsuario(usuario);
      const calculo = registrarSalida(area, existente.entrada, hora, fechaTextoMostrar, minutosComida, correo, empresaId);

      const obs = (existente.observaciones || "") + " [Salida capturada desde web]";
      update("asistencia", existente.id, {
        salida: hora,
        horas_brutas: calculo.brutas,
        comida: calculo.comida,
        horas_trabajadas: calculo.trabajadas,
        horas_extra: calculo.extra,
        estado: calculo.estado,
        observaciones: obs,
        ...(latitud !== null ? { latitud } : {}),
        ...(longitud !== null ? { longitud } : {})
      });

      eliminarPendienteWeb(empresaId, correo, fechaTexto, existente.id);
      actualizarPeriodoLaboral(empresaId, correo, fecha);

      return res.json({
        ok: true,
        tipo: "salida",
        mensaje: "Salida registrada y jornada calculada.",
        data: { id: existente.id, entrada: existente.entrada, salida: hora, ...calculo }
      });
    }

    case "comida_entrada": {
      const existente = findOne(
        "asistencia",
        a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fechaTexto
      );

      if (!existente || !existente.entrada) {
        return res.status(400).json({ ok: false, error: "Primero debes registrar tu entrada." });
      }
      if (existente.salida) {
        return res.status(400).json({ ok: false, error: "Ya registraste salida; no se puede marcar comida." });
      }
      if (existente.comida_entrada && !existente.comida_salida) {
        return res.status(400).json({ ok: false, error: "Ya tienes una salida a comida pendiente de regreso." });
      }

      const obs = (existente.observaciones || "") + `\nComida entrada: ${hora} (web)`;
      update("asistencia", existente.id, { observaciones: obs, comida_entrada: hora, comida_salida: "" });

      return res.json({ ok: true, tipo: "comida_entrada", mensaje: "Entrada de comida registrada.", data: { hora } });
    }

    case "comida_salida": {
      const existente = findOne(
        "asistencia",
        a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fechaTexto
      );

      if (!existente || !existente.entrada) {
        return res.status(400).json({ ok: false, error: "No hay entrada registrada hoy." });
      }
      if (!existente.comida_entrada || existente.comida_salida) {
        return res.status(400).json({ ok: false, error: "Primero registra tu salida a comida." });
      }

      const obs = (existente.observaciones || "") + `\nComida salida: ${hora} (web)`;
      update("asistencia", existente.id, { observaciones: obs, comida_salida: hora });

      return res.json({ ok: true, tipo: "comida_salida", mensaje: "Salida de comida registrada.", data: { hora } });
    }

    default:
      return res.status(400).json({ ok: false, error: "Tipo de marcación inválido." });
  }
});

// =============================================
// CONSULTAS DE USUARIO
// =============================================

app.get("/api/asistencia/semana", requireAuth, (req, res) => {
  const usuario = req.usuario;
  const empresaId = usuario.empresa_id;
  const correo = usuario.correo;
  const { fecha } = req.query;

  const baseFecha = fecha ? fechaDesdeEntrada(fecha) : new Date();
  const lunes = lunesDeLaSemana(baseFecha);
  const sabado = new Date(lunes.getTime()); sabado.setDate(sabado.getDate() + 5);

  const desde = claveFecha(lunes);
  const hasta = claveFecha(sabado);

  const filas = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === correo && a.fecha >= desde && a.fecha <= hasta
  ).sort((a, b) => a.fecha.localeCompare(b.fecha));

  const mapa = {};
  filas.forEach(f => { mapa[f.fecha] = f; });

  const dias = [];
  for (let i = 0; i < 6; i++) {
    const d = new Date(lunes.getTime()); d.setDate(d.getDate() + i);
    const clave = claveFecha(d);
    const exist = mapa[clave];
    if (exist) {
      dias.push(exist);
    } else {
      dias.push({
        fecha: clave,
        fecha_texto: new Intl.DateTimeFormat("es-MX", { day: "2-digit", month: "2-digit", year: "numeric" }).format(d),
        dia_semana: d.getDay(),
        nombre_dia: NOMBRES_DIA[d.getDay()],
        entrada: null, salida: null, estado: "Sin registro",
        horas_trabajadas: "0:00", horas_extra: "0:00"
      });
    }
  }

  res.json({ ok: true, semana: { lunes: desde, sabado: hasta }, dias });
});

app.get("/api/asistencia/totales-semana", requireAuth, (req, res) => {
  const usuario = req.usuario;
  const empresaId = usuario.empresa_id;
  const correo = usuario.correo;
  const { fecha } = req.query;
  const baseFecha = fecha ? fechaDesdeEntrada(fecha) : new Date();
  const lunes = lunesDeLaSemana(baseFecha);
  const sabado = new Date(lunes.getTime()); sabado.setDate(sabado.getDate() + 5);
  const desde = claveFecha(lunes);
  const hasta = claveFecha(sabado);

  const filas = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === correo && a.fecha >= desde && a.fecha <= hasta
  ).sort((a, b) => a.fecha.localeCompare(b.fecha));

  let brutas = 0, comida = 0, trabajadas = 0, extraPropia = 0, extraSabado = 0;
  let asistencias = 0, faltas = 0, abiertos = 0;
  const incidencias = [];

  filas.forEach(f => {
    const nombreDia = NOMBRES_DIA[f.dia_semana];
    if (nombreDia === "Domingo") {
      if (f.entrada || f.salida || f.observaciones) {
        incidencias.push("Domingo: " + (f.observaciones || "Asistencia registrada, no contabilizada."));
      }
      return;
    }
    if (f.estado === "Falta") {
      faltas++; incidencias.push(nombreDia + ": " + (f.observaciones || "No se registró entrada."));
      return;
    }
    if (!f.entrada) {
      if (f.observaciones) incidencias.push(nombreDia + ": " + f.observaciones);
      return;
    }
    if (!f.salida) {
      abiertos++; asistencias++;
      incidencias.push(nombreDia + ": Entrada registrada; salida pendiente.");
      return;
    }
    asistencias++;
    brutas += aMinutos(f.horas_brutas) || 0;
    comida += aMinutos(f.comida) || 0;
    trabajadas += aMinutos(f.horas_trabajadas) || 0;
    const minE = aMinutos(f.horas_extra) || 0;
    if (nombreDia === "Sábado") {
      extraSabado += minE;
    } else {
      extraPropia += minE;
    }
  });

  // Extra traída de la semana anterior (sábado)
  const lunesAnterior = new Date(lunes.getTime()); lunesAnterior.setDate(lunesAnterior.getDate() - 7);
  const desdeAnt = claveFecha(lunesAnterior);
  const hastaAnt = claveFecha(new Date(lunesAnterior.getTime() + 5 * 86400000));
  const filasAnt = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === correo && a.fecha >= desdeAnt && a.fecha <= hastaAnt
  );
  let extraTraida = 0;
  filasAnt.forEach(f => {
    if (NOMBRES_DIA[f.dia_semana] === "Sábado") extraTraida += aMinutos(f.horas_extra) || 0;
  });

  const extraTotal = extraPropia + extraTraida;

  res.json({
    ok: true,
    semana: { lunes: desde, sabado: hasta },
    brutas: aTextoHoras(brutas),
    comida: aTextoHoras(comida),
    trabajadas: aTextoHoras(trabajadas),
    extra: aTextoHoras(extraTotal),
    extraPropia: aTextoHoras(extraPropia),
    extraTraida: aTextoHoras(extraTraida),
    asistencias, faltas, abiertos,
    incidencias
  });
});

app.get("/api/asistencia/pendientes", requireAuth, (req, res) => {
  const usuario = req.usuario;
  const empresaId = usuario.empresa_id;
  const correo = usuario.correo;

  const filas = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === correo && a.entrada && a.entrada !== "" && (!a.salida || a.salida === "")
  ).sort((a, b) => b.fecha.localeCompare(a.fecha));

  res.json({ ok: true, pendientes: filas });
});

// =============================================
// ADMIN / RH
// =============================================

app.get("/api/admin/usuarios", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const filas = find("usuarios", u => u.empresa_id === empresaId).sort((a, b) => a.nombre.localeCompare(b.nombre));
  res.json({ ok: true, usuarios: filas.map(fila => {
    const { password_hash, _alta_compartida, _registro_codigo, _importado, ...usuario } = fila;
    return {
      ...usuario,
      cuenta_activada: Boolean(password_hash),
      se_unio_con_codigo: usuarioSeUnioConCodigo(fila)
    };
  }) });
});

app.get("/api/admin/areas", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  crearAreasPredeterminadas(empresaId);
  const empresa = findOne("empresa", e => e.id === empresaId);
  const areas = find("areas", a => a.empresa_id === empresaId)
    .sort((a, b) => Number(b.activo) - Number(a.activo) || a.nombre.localeCompare(b.nombre));
  res.json({ ok: true, areas, configuracion_inicial_pendiente: empresa?.areas_configuradas === false });
});

app.post("/api/admin/areas", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { id, nombre, entrada, salida, activo } = req.body || {};
  const nombreLimpio = String(nombre || "").trim();
  if (!nombreLimpio || nombreLimpio.length > 80) {
    return res.status(400).json({ ok: false, error: "Escribe un nombre de área de hasta 80 caracteres." });
  }
  const horaEntrada = normalizarHoraManual(entrada);
  const horaSalida = normalizarHoraManual(salida);
  if (!horaEntrada || !horaSalida || aMinutos(horaSalida) <= aMinutos(horaEntrada)) {
    return res.status(400).json({ ok: false, error: "Define un horario válido; la salida debe ser posterior a la entrada." });
  }
  const existente = id ? findOne("areas", a => a.id === Number(id) && a.empresa_id === empresaId) : null;
  if (id && !existente) return res.status(404).json({ ok: false, error: "Área no encontrada en este equipo." });
  const duplicada = findOne("areas", a => a.empresa_id === empresaId && a.id !== Number(id || 0) && normalizar(a.nombre) === normalizar(nombreLimpio));
  if (duplicada) return res.status(400).json({ ok: false, error: "Ya existe un área con ese nombre." });
  const cambios = { nombre: nombreLimpio, entrada: horaEntrada, salida: horaSalida, activo: activo === false ? 0 : 1 };
  if (existente && Number(existente.activo) === 1 && cambios.activo === 0 && findOne("usuarios", u => u.empresa_id === empresaId && u.activo === 1 && normalizar(u.area || "") === normalizar(existente.nombre))) {
    return res.status(409).json({ ok: false, error: "Reasigna primero a los colaboradores activos de esta área." });
  }
  if (existente && normalizar(existente.nombre) !== normalizar(nombreLimpio)) {
    db.usuarios.filter(u => u.empresa_id === empresaId && normalizar(u.area || "") === normalizar(existente.nombre))
      .forEach(u => { u.area = nombreLimpio; });
  }
  if (existente) update("areas", existente.id, cambios);
  else insert("areas", { empresa_id: empresaId, ...cambios, creado_en: new Date().toISOString() });
  guardarDB();
  limpiarCacheHorarios(empresaId);
  res.json({ ok: true });
});

app.post("/api/admin/areas/confirmar", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  if (!areasActivasEmpresa(empresaId).length) {
    return res.status(400).json({ ok: false, error: "Activa al menos un área antes de continuar." });
  }
  const empresa = findOne("empresa", e => e.id === empresaId);
  if (!empresa) return res.status(404).json({ ok: false, error: "Empresa no encontrada." });
  empresa.areas_configuradas = true;
  guardarDB();
  res.json({ ok: true });
});

app.delete("/api/admin/areas/:id", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const area = findOne("areas", a => a.id === Number(req.params.id) && a.empresa_id === empresaId);
  if (!area) return res.status(404).json({ ok: false, error: "Área no encontrada en este equipo." });
  if (findOne("usuarios", u => u.empresa_id === empresaId && u.activo === 1 && normalizar(u.area || "") === normalizar(area.nombre))) {
    return res.status(409).json({ ok: false, error: "Reasigna primero a los colaboradores activos de esta área." });
  }
  update("areas", area.id, { activo: 0 });
  limpiarCacheHorarios(empresaId);
  res.json({ ok: true });
});

function horaActualEnZona() {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: obtenerZonaHoraria(), hour: "2-digit", minute: "2-digit", hourCycle: "h23"
  }).format(new Date());
}

function usuarioDebeHaberTerminado(usuario, fecha, empresaId) {
  if (fecha.getDay() === 0) return false;
  const hoy = claveFecha(new Date());
  const clave = claveFecha(fecha);
  if (clave < hoy) return true;
  if (clave > hoy) return false;
  const horario = obtenerHorarioParaFecha(usuario.area, fecha, usuario.correo, empresaId);
  return aMinutos(horaActualEnZona()) >= aMinutos(horario.fin);
}

app.get("/api/admin/asistencia-dia", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const fecha = req.query.fecha ? fechaDesdeEntrada(req.query.fecha) : normalizarFechaSoloDia(new Date());
  if (!fecha || isNaN(fecha.getTime())) return res.status(400).json({ ok: false, error: "Fecha inválida." });
  const clave = claveFecha(fecha);
  const usuarios = find("usuarios", u => u.empresa_id === empresaId && u.activo === 1)
    .sort((a, b) => (a.area || "").localeCompare(b.area || "") || (a.nombre || "").localeCompare(b.nombre || ""));
  const registros = find("asistencia", a => a.empresa_id === empresaId && a.fecha === clave);
  const porCorreo = new Map(registros.map(fila => [fila.correo, fila]));

  const asistencias = usuarios.map(usuario => {
    const registro = porCorreo.get(usuario.correo) || null;
    const horario = obtenerHorarioParaFecha(usuario.area, fecha, usuario.correo, empresaId);
    const fechaIngreso = fechaISOValida(usuario.fecha_ingreso) ? usuario.fecha_ingreso : null;
    let estado = registro?.estado || "";
    if (!estado && fechaIngreso && clave < fechaIngreso) estado = "Antes del ingreso";
    if (!estado && fecha.getDay() === 0) estado = "Descanso";
    if (!estado && usuarioDebeHaberTerminado(usuario, fecha, empresaId)) estado = "Falta";
    if (!estado) estado = "Sin registro";
    return {
      nombre: usuario.nombre,
      correo: usuario.correo,
      area: usuario.area || "Sin asignar",
      horario: `${horario.inicio} – ${horario.fin}`,
      entrada: registro?.entrada || "",
      comida_entrada: registro?.comida_entrada || "",
      comida_salida: registro?.comida_salida || "",
      salida: registro?.salida || "",
      horas_trabajadas: registro?.horas_trabajadas || "",
      horas_extra: registro?.horas_extra || "",
      estado,
      observaciones: registro?.observaciones || "",
      origen: registro?.origen || "",
      latitud: registro?.latitud ?? null,
      longitud: registro?.longitud ?? null
    };
  });

  const resumen = {
    empleados: usuarios.length,
    asistencias: asistencias.filter(f => f.entrada).length,
    sinSalida: asistencias.filter(f => f.entrada && !f.salida).length,
    faltas: asistencias.filter(f => f.estado === "Falta").length,
    horas: aTextoHoras(asistencias.reduce((suma, f) => suma + (aMinutos(f.horas_trabajadas) || 0), 0))
  };
  res.json({ ok: true, fecha: clave, dia: NOMBRES_DIA[fecha.getDay()], resumen, asistencias });
});

app.get("/api/admin/resumen", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const fecha = req.query.fecha ? fechaDesdeEntrada(req.query.fecha) : normalizarFechaSoloDia(new Date());
  if (!fecha || isNaN(fecha.getTime())) return res.status(400).json({ ok: false, error: "Fecha inválida." });
  const clave = claveFecha(fecha);
  const usuarios = find("usuarios", u => u.empresa_id === empresaId && u.activo === 1);
  const filas = find("asistencia", a => a.empresa_id === empresaId && a.fecha === clave);
  const porCorreo = new Map(filas.map(f => [f.correo, f]));
  const minutosHoy = filas.reduce((suma, f) => suma + (aMinutos(f.horas_trabajadas) || 0), 0);
  const ausentes = usuarios.filter(u => {
    const fila = porCorreo.get(u.correo);
    if (fila?.estado === "Falta") return true;
    if (fila?.entrada || clave < (u.fecha_ingreso || "0000-00-00")) return false;
    return usuarioDebeHaberTerminado(u, fecha, empresaId);
  }).length;
  const abiertos = filas.filter(a => a.entrada && !a.salida).length;
  res.json({ ok: true, fecha: clave, empleados: usuarios.length,
    asistencias: filas.filter(f => f.entrada).length,
    enTurno: filas.filter(f => f.entrada && !f.salida).length,
    sinSalida: abiertos, faltas: ausentes, horas: aTextoHoras(minutosHoy) });
});

app.get("/api/admin/suscripcion", requireAdmin, (req, res) => {
  const empresa = findOne("empresa", e => e.id === req.usuario.empresa_id);
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  const pagos = find("pagos_suscripcion", p => p.empresa_id === empresa.id)
    .sort((a, b) => String(b.registrado_en).localeCompare(String(a.registrado_en)));
  res.json({
    ok: true,
    suscripcion: {
      estado: estadoSuscripcion(empresa),
      solicitud_estado: empresa.solicitud_codigo_estado || "sin_solicitud",
      solicitud_en: empresa.solicitud_codigo_en || null,
      codigo_equipo: suscripcionActiva(empresa) ? empresa.codigo : null,
      cupos_limite: Number(empresa.cupos_limite || 0),
      cupos_ocupados: contarCuposOcupados(empresa.id),
      monto_pagado_mxn: Number(empresa.monto_pagado_mxn || 0),
      monto_ciclo_mxn: empresa.monto_ciclo_mxn === undefined ? null : Number(empresa.monto_ciclo_mxn),
      periodicidad: empresa.periodicidad || "mensual",
      fecha_inicio_suscripcion: empresa.fecha_inicio_suscripcion || null,
      fecha_fin_suscripcion: empresa.fecha_fin_suscripcion || null,
      ultima_nota_pago: pagos[0] || null
    }
  });
});

app.post("/api/admin/solicitar-codigo", requireAdmin, (req, res) => {
  const empresa = findOne("empresa", e => e.id === req.usuario.empresa_id);
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  const cupos = Number(req.body?.cupos_solicitados);
  if (!Number.isInteger(cupos) || cupos < 1 || cupos > 10000) return res.status(400).json({ ok: false, error: "Indica un número de cupos válido." });
  const esAmpliacion = req.body?.tipo === "ampliacion";
  if (suscripcionActiva(empresa) && !esAmpliacion) return res.status(409).json({ ok: false, error: "Tu código ya está activo." });
  if (esAmpliacion && (!suscripcionActiva(empresa) || cupos <= Number(empresa.cupos_limite || 0))) {
    return res.status(400).json({ ok: false, error: "El rango solicitado debe superar los cupos actuales de tu equipo." });
  }
  const ahora = new Date().toISOString();
  update("empresa", empresa.id, {
    solicitud_codigo_estado: esAmpliacion ? "solicitada_ampliacion" : "solicitada",
    solicitud_codigo_en: ahora,
    cupos_solicitados: cupos,
    estado_suscripcion: !esAmpliacion && empresa.estado_suscripcion === "declinada" ? "pendiente" : empresa.estado_suscripcion
  });
  const nombreSolicitante = req.usuario.nombre;
  notificarSuperadmins({
    clave: `solicitud-codigo-${empresa.id}-${ahora.slice(0, 16)}`,
    titulo: esAmpliacion ? "Solicitud de ampliación" : "Solicitud de código",
    cuerpo: `${empresa.nombre} solicita ${esAmpliacion ? "ampliar a" : "activación para"} ${cupos === 31 ? "31 o más" : `hasta ${cupos}`} personas.`,
    url: "/superadmin.html"
  });
  res.json({ ok: true, mensaje: `${esAmpliacion ? "Solicitud de ampliación" : "Solicitud de código"} registrada para ${nombreSolicitante}. Continúa por WhatsApp para confirmar el pago.` });
});

app.post("/api/admin/usuarios", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { id, correo, nombre, area, activo, minutos_comida, fecha_ingreso, rol } = req.body || {};
  if (!correo || !nombre) return res.status(400).json({ ok: false, error: "Correo y nombre son obligatorios" });
  const correoNormalizado = normalizarCorreo(correo);
  const fechaIngresoNormalizada = fecha_ingreso ? String(fecha_ingreso).trim() : null;
  if (fechaIngresoNormalizada && (!fechaISOValida(fechaIngresoNormalizada) || fechaIngresoNormalizada > claveFecha(new Date()))) {
    return res.status(400).json({ ok: false, error: "La fecha de ingreso debe ser una fecha válida y no futura." });
  }
  const existente = id ? findOne("usuarios", u => u.id === Number(id) && u.empresa_id === empresaId) : null;
  if (id && !existente) return res.status(404).json({ ok: false, error: "Usuario no encontrado en este equipo." });
  const areaNormalizada = String(area || "Sin asignar").trim();
  if (normalizar(areaNormalizada) !== normalizar("Sin asignar") && !areasActivasEmpresa(empresaId).some(a => normalizar(a.nombre) === normalizar(areaNormalizada))) {
    return res.status(400).json({ ok: false, error: "Selecciona un área activa de este equipo." });
  }
  const duplicado = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === correoNormalizado && u.id !== Number(id || 0));
  if (duplicado) return res.status(400).json({ ok: false, error: "El correo ya existe en esta empresa." });

  const empresa = findOne("empresa", e => e.id === empresaId);
  const rolNormalizado = rol === "admin" ? "admin" : "empleado";
  const activoDestino = activo !== undefined ? (activo ? 1 : 0) : 1;
  const sumaCupo = rolNormalizado !== "admin" && activoDestino === 1
    && (!existente || existente.activo !== 1 || existente.rol === "admin");
  if (sumaCupo && Number(empresa?.cupos_limite || 0) <= contarCuposOcupados(empresaId)) {
    return res.status(409).json({ ok: false, error: "Este equipo llegó al límite de cupos. Solicita una ampliación al superadmin." });
  }

  const datos = {
    empresa_id: empresaId,
    correo: correoNormalizado,
    nombre: nombre.trim(),
    area: areaNormalizada,
    activo: activo !== undefined ? (activo ? 1 : 0) : 1,
    rol: rol === "admin" ? "admin" : "empleado",
    _alta_compartida: true,
    _importado: false,
    minutos_comida: minutos_comida ? normalizarMinutosComida(minutos_comida) : MINUTOS_COMIDA_POR_DEFECTO,
    fecha_ingreso: fechaIngresoNormalizada,
    fecha_alta: existente?.fecha_alta || new Date().toISOString()
  };

  if (id) {
    const quedaAdmin = existente.rol === "admin" && (datos.rol !== "admin" || datos.activo !== 1);
    const adminsActivos = find("usuarios", u => u.empresa_id === empresaId && u.rol === "admin" && u.activo === 1).length;
    if (quedaAdmin && adminsActivos <= 1) return res.status(400).json({ ok: false, error: "El equipo debe conservar al menos un administrador activo." });
    update("usuarios", Number(id), datos);
  } else {
    insert("usuarios", datos);
  }
  limpiarCacheHorarios(empresaId);
  notificarAdministradoresCupo(empresaId);
  res.json({ ok: true });
});

app.delete("/api/admin/usuarios/:id", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const usuario = findOne("usuarios", u => u.id === Number(req.params.id) && u.empresa_id === empresaId);
  if (!usuario) return res.status(404).json({ ok: false, error: "Usuario no encontrado en este equipo." });
  const adminsActivos = find("usuarios", u => u.empresa_id === empresaId && u.rol === "admin" && u.activo === 1).length;
  if (usuario.rol === "admin" && adminsActivos <= 1) return res.status(400).json({ ok: false, error: "El equipo debe conservar al menos un administrador activo." });
  update("usuarios", usuario.id, { activo: 0 });
  res.json({ ok: true });
});

app.post("/api/admin/codigo-equipo", requireAdmin, (req, res) => {
  res.status(403).json({ ok: false, error: "Los códigos de invitación los administra JJM Clock. Solicita ayuda desde Código y notificaciones." });
});

app.post("/api/admin/vaciar-equipo", requireAdmin, (req, res) => {
  if (String(req.body?.confirmacion || "").trim().toUpperCase() !== "VACIAR EQUIPO") {
    return res.status(400).json({ ok: false, error: "Escribe VACIAR EQUIPO para confirmar la acción." });
  }

  const empresaId = req.usuario.empresa_id;
  const usuariosEmpresa = find("usuarios", u => u.empresa_id === empresaId);
  const administradores = usuariosEmpresa.filter(u => u.rol === "admin");
  const idsColaboradores = new Set(usuariosEmpresa.filter(u => u.rol !== "admin").map(u => u.id));
  const eliminados = {
    colaboradores: idsColaboradores.size,
    asistencias: find("asistencia", a => a.empresa_id === empresaId).length,
    marcasPendientes: find("pendientes", p => p.empresa_id === empresaId).length,
    horariosEspeciales: find("horarios_especiales", h => h.empresa_id === empresaId).length,
    periodosLaborales: find("periodo_laboral", p => p.empresa_id === empresaId).length,
    solicitudesVacaciones: find("vacaciones_solicitud", v => v.empresa_id === empresaId).length
  };

  db.usuarios = db.usuarios.filter(u => u.empresa_id !== empresaId || u.rol === "admin");
  db.asistencia = db.asistencia.filter(a => a.empresa_id !== empresaId);
  db.pendientes = db.pendientes.filter(p => p.empresa_id !== empresaId);
  db.horarios_especiales = db.horarios_especiales.filter(h => h.empresa_id !== empresaId);
  db.periodo_laboral = db.periodo_laboral.filter(p => p.empresa_id !== empresaId);
  db.vacaciones_solicitud = db.vacaciones_solicitud.filter(v => v.empresa_id !== empresaId);
  db.sesiones = db.sesiones.filter(s => s.empresa_id !== empresaId || !idsColaboradores.has(s.usuario_id));
  guardarDB();
  limpiarCacheHorarios(empresaId);

  res.json({ ok: true, eliminados, administradoresConservados: administradores.length,
    excelOriginalConservado: Boolean(findOne("empresa", e => e.id === empresaId)?.archivo_horas_original) });
});

app.get("/api/admin/horarios-especiales", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const filas = find("horarios_especiales", h => h.empresa_id === empresaId)
    .sort((a, b) => (b.vigente_desde || "").localeCompare(a.vigente_desde || ""));
  res.json({ ok: true, horarios: filas });
});

app.post("/api/admin/horarios-especiales", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { id, correo, nombre_referencia, dia, entrada, salida, vigente_desde, vigente_hasta, activo } = req.body || {};
  if (!correo || !dia || !entrada || !salida) {
    return res.status(400).json({ ok: false, error: "Faltan datos: correo, día, entrada, salida" });
  }
  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === normalizarCorreo(correo));
  if (!usuarioSeUnioConCodigo(usuario)) return res.status(400).json({ ok: false, error: "Selecciona a una persona que ya se haya unido con el código del equipo." });

  const ent = normalizarHoraManual(entrada);
  const sal = normalizarHoraManual(salida);
  if (!ent || !sal) return res.status(400).json({ ok: false, error: "Horas inválidas (usa HH:mm)" });

  const datos = {
    empresa_id: empresaId,
    correo: correo.trim().toLowerCase(),
    nombre_referencia: nombre_referencia || "",
    dia: dia.trim(),
    entrada: ent,
    salida: sal,
    vigente_desde: vigente_desde || null,
    vigente_hasta: vigente_hasta || null,
    activo: activo !== undefined ? (activo ? 1 : 0) : 1,
    creado_en: new Date().toISOString()
  };

  if (id) {
    const existente = findOne("horarios_especiales", h => h.id === Number(id) && h.empresa_id === empresaId);
    if (!existente) return res.status(404).json({ ok: false, error: "Horario no encontrado en este equipo." });
    update("horarios_especiales", Number(id), datos);
  } else {
    insert("horarios_especiales", datos);
  }
  limpiarCacheHorarios(empresaId);
  res.json({ ok: true });
});

app.delete("/api/admin/horarios-especiales/:id", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const horario = findOne("horarios_especiales", h => h.id === Number(req.params.id) && h.empresa_id === empresaId);
  if (!horario) return res.status(404).json({ ok: false, error: "Horario no encontrado en este equipo." });
  remove("horarios_especiales", horario.id);
  limpiarCacheHorarios(empresaId);
  res.json({ ok: true });
});

app.get("/api/admin/reporte-semanal", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { fecha } = req.query;
  const baseFecha = fecha ? fechaDesdeEntrada(fecha) : new Date();
  const lunes = lunesDeLaSemana(baseFecha);
  const sabado = new Date(lunes.getTime()); sabado.setDate(sabado.getDate() + 5);
  const desde = claveFecha(lunes);
  const hasta = claveFecha(sabado);

  const usuarios = find("usuarios", u => u.empresa_id === empresaId && u.activo === 1).sort((a, b) => a.nombre.localeCompare(b.nombre));

  const filas = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.fecha >= desde && a.fecha <= hasta
  );

  const mapa = new Map();
  filas.forEach(f => {
    if (!mapa.has(f.correo)) mapa.set(f.correo, new Map());
    mapa.get(f.correo).set(f.fecha, f);
  });

  const primeraAsistenciaPorCorreo = new Map();
  find("asistencia", a => a.empresa_id === empresaId && a.entrada && a.fecha <= hasta).forEach(a => {
    const actual = primeraAsistenciaPorCorreo.get(a.correo);
    if (!actual || a.fecha < actual) primeraAsistenciaPorCorreo.set(a.correo, a.fecha);
  });

  let totalTrabajadas = 0, totalExtra = 0, totalAsistencias = 0, totalFaltas = 0;
  const reporte = usuarios.map(u => {
    const diasUsuario = mapa.get(u.correo) || new Map();
    // fecha_alta describe cuándo se creó/importó la cuenta, no cuándo inició el empleo.
    // Si no hay fecha laboral capturada, usamos la primera marca histórica para no
    // ocultar jornadas importadas anteriores a la cuenta; la antigüedad queda aparte.
    const fechaAlta = fechaISOValida(u.fecha_ingreso)
      ? u.fecha_ingreso
      : (primeraAsistenciaPorCorreo.get(u.correo) || desde);
    let trabajadas = 0, extra = 0, asistencias = 0, faltas = 0;
    const detalle = [];
    for (let i = 0; i < 6; i++) {
      const dia = new Date(lunes.getTime());
      dia.setDate(dia.getDate() + i);
      const claveDia = claveFecha(dia);
      if (claveDia < fechaAlta) {
        detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: "Antes del ingreso" });
        continue;
      }
      const fila = diasUsuario.get(claveDia);
      if (fila?.estado === "Falta") {
        faltas++;
        detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: "Falta", entrada: "", salida: "", horas: "0:00", extra: "0:00" });
        continue;
      }
      if (!fila?.entrada) {
        const falta = usuarioDebeHaberTerminado(u, dia, empresaId);
        if (falta) faltas++;
        detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: falta ? "Falta" : "Pendiente", entrada: "", salida: "", horas: "0:00", extra: "0:00" });
        continue;
      }
      asistencias++;
      trabajadas += aMinutos(fila.horas_trabajadas) || 0;
      const minutosExtraDia = aMinutos(fila.horas_extra) || 0;
      // La extra del sábado se acumula en la semana siguiente, como en V8.
      if (dia.getDay() !== 6) extra += minutosExtraDia;
      detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: fila.estado || "Asistencia",
        entrada: fila.entrada || "", comida_entrada: fila.comida_entrada || "", comida_salida: fila.comida_salida || "",
        salida: fila.salida || "", horas: fila.horas_trabajadas || "0:00", extra: aTextoHoras(minutosExtraDia) });
    }
    const lunesAnterior = new Date(lunes.getTime());
    lunesAnterior.setDate(lunesAnterior.getDate() - 7);
    const sabadoAnterior = new Date(lunesAnterior.getTime());
    sabadoAnterior.setDate(sabadoAnterior.getDate() + 5);
    const claveSabadoAnterior = claveFecha(sabadoAnterior);
    const asistenciaSabadoAnterior = findOne("asistencia", a => a.empresa_id === empresaId && a.correo === u.correo
      && a.fecha === claveSabadoAnterior && a.dia_semana === 6);
    if (asistenciaSabadoAnterior) extra += aMinutos(asistenciaSabadoAnterior.horas_extra) || 0;
    totalTrabajadas += trabajadas;
    totalExtra += extra;
    totalAsistencias += asistencias;
    totalFaltas += faltas;
    return { correo: u.correo, nombre: u.nombre, area: u.area,
      trabajadas: aTextoHoras(trabajadas), extra: aTextoHoras(extra), asistencias, faltas, detalle };
  });

  res.json({ ok: true, semana: { lunes: desde, sabado: hasta }, reporte,
    totales: { trabajadas: aTextoHoras(totalTrabajadas), extra: aTextoHoras(totalExtra), asistencias: totalAsistencias, faltas: totalFaltas } });
});

app.get("/api/admin/reporte-individual", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const correo = normalizarCorreo(req.query.correo || "");
  const desde = String(req.query.fecha_ini || "");
  const hasta = String(req.query.fecha_fin || "");
  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === correo);
  if (!usuarioSeUnioConCodigo(usuario)) return res.status(404).json({ ok: false, error: "Selecciona a una persona que ya se haya unido con el código del equipo." });
  if (!fechaISOValida(desde) || !fechaISOValida(hasta) || desde > hasta) {
    return res.status(400).json({ ok: false, error: "Revisa el rango de fechas del reporte." });
  }
  const diasCalendario = (fechaDesdeEntrada(hasta).getTime() - fechaDesdeEntrada(desde).getTime()) / 86400000;
  if (diasCalendario > 366) {
    return res.status(400).json({ ok: false, error: "El reporte individual admite rangos de hasta un año." });
  }

  const primeraAsistencia = find("asistencia", a => a.empresa_id === empresaId && a.correo === correo && a.entrada)
    .map(a => a.fecha).sort()[0];
  const fechaBase = fechaISOValida(usuario.fecha_ingreso) ? usuario.fecha_ingreso : (primeraAsistencia || desde);
  const registros = new Map(find("asistencia", a => a.empresa_id === empresaId && a.correo === correo
    && a.fecha >= desde && a.fecha <= hasta).map(a => [a.fecha, a]));
  const detalle = [];
  const cursor = fechaDesdeEntrada(desde);
  const fin = fechaDesdeEntrada(hasta);
  let asistencias = 0, faltas = 0, trabajadas = 0, extra = 0;
  while (cursor <= fin) {
    const fecha = claveFecha(cursor);
    if (cursor.getDay() !== 0) {
      if (fecha < fechaBase) {
        detalle.push({ fecha, dia: NOMBRES_DIA[cursor.getDay()], estado: "Antes del ingreso", entrada: "", comida_entrada: "", comida_salida: "", salida: "", horas: "0:00", extra: "0:00" });
      } else {
        const fila = registros.get(fecha);
        if (fila?.entrada) {
          asistencias++;
          const minutos = aMinutos(fila.horas_trabajadas) || 0;
          const minutosExtra = aMinutos(fila.horas_extra) || 0;
          trabajadas += minutos;
          extra += minutosExtra;
          detalle.push({ fecha, dia: NOMBRES_DIA[cursor.getDay()], estado: fila.estado || "Asistencia",
            entrada: fila.entrada || "", comida_entrada: fila.comida_entrada || "", comida_salida: fila.comida_salida || "",
            salida: fila.salida || "", horas: fila.horas_trabajadas || "0:00", extra: aTextoHoras(minutosExtra) });
        } else {
          const falta = usuarioDebeHaberTerminado(usuario, cursor, empresaId);
          if (falta) faltas++;
          detalle.push({ fecha, dia: NOMBRES_DIA[cursor.getDay()], estado: falta ? "Falta" : "Pendiente",
            entrada: "", comida_entrada: "", comida_salida: "", salida: "", horas: "0:00", extra: "0:00" });
        }
      }
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  res.json({ ok: true, trabajador: { correo, nombre: usuario.nombre, area: usuario.area || "Sin asignar", fecha_ingreso: usuario.fecha_ingreso || null },
    periodo: { desde, hasta }, totales: { asistencias, faltas, trabajadas: aTextoHoras(trabajadas), extra: aTextoHoras(extra) }, detalle });
});

app.post("/api/admin/asistencia-manual", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const correo = normalizarCorreo(req.body?.correo || "");
  const fecha = String(req.body?.fecha || "");
  const motivo = String(req.body?.motivo || "").trim().slice(0, 500);
  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === correo && u.activo === 1 && usuarioSeUnioConCodigo(u));
  if (!usuario) return res.status(404).json({ ok: false, error: "Selecciona a un trabajador activo que ya se haya unido con el código del equipo." });
  if (!fechaISOValida(fecha) || fecha > claveFecha(new Date())) {
    return res.status(400).json({ ok: false, error: "La fecha debe ser válida y no futura." });
  }
  if (!motivo) return res.status(400).json({ ok: false, error: "Escribe el motivo de la captura o corrección." });

  const campos = ["entrada", "comida_entrada", "comida_salida", "salida"];
  const tiempos = {};
  for (const campo of campos) {
    const original = String(req.body?.[campo] || "").trim();
    const normalizada = original ? normalizarHoraManual(original) : "";
    if (original && !normalizada) return res.status(400).json({ ok: false, error: `La hora de ${campo.replaceAll("_", " ")} no es válida.` });
    tiempos[campo] = normalizada || "";
  }
  if (!tiempos.entrada) return res.status(400).json({ ok: false, error: "La hora de entrada es obligatoria para guardar el registro." });
  if (tiempos.salida && aMinutos(tiempos.salida) < aMinutos(tiempos.entrada)) return res.status(400).json({ ok: false, error: "La salida no puede ser anterior a la entrada." });
  if (tiempos.comida_salida && !tiempos.comida_entrada) return res.status(400).json({ ok: false, error: "Captura primero la salida a comida." });
  if (tiempos.comida_entrada && aMinutos(tiempos.comida_entrada) < aMinutos(tiempos.entrada)) return res.status(400).json({ ok: false, error: "La salida a comida no puede ser anterior a la entrada." });
  if (tiempos.comida_salida && aMinutos(tiempos.comida_salida) < aMinutos(tiempos.comida_entrada)) return res.status(400).json({ ok: false, error: "El regreso de comida debe ser posterior a la salida a comida." });
  if (tiempos.salida && tiempos.comida_entrada && aMinutos(tiempos.salida) < aMinutos(tiempos.comida_entrada)) return res.status(400).json({ ok: false, error: "La salida debe ser posterior a la salida a comida." });
  if (tiempos.salida && tiempos.comida_salida && aMinutos(tiempos.salida) < aMinutos(tiempos.comida_salida)) return res.status(400).json({ ok: false, error: "La salida debe ser posterior al regreso de comida." });

  const fechaObj = fechaDesdeEntrada(fecha);
  const fechaTexto = fechaTextoVisual(fechaObj);
  const existente = findOne("asistencia", a => a.empresa_id === empresaId && a.correo === correo && a.fecha === fecha);
  let calculo = null;
  if (tiempos.salida) calculo = registrarSalida(usuario.area, tiempos.entrada, tiempos.salida, fechaTexto,
    obtenerMinutosComidaUsuario(usuario), correo, empresaId);
  const diaDomingo = fechaObj.getDay() === 0;
  const camposRegistro = {
    empresa_id: empresaId, correo, nombre: usuario.nombre, area: usuario.area || "Sin asignar",
    fecha, fecha_texto: fechaTexto, dia_semana: fechaObj.getDay(), ...tiempos,
    horas_brutas: calculo?.brutas || (diaDomingo ? "0:00" : ""),
    comida: calculo?.comida || (diaDomingo ? "0:00" : ""),
    horas_trabajadas: calculo?.trabajadas || (diaDomingo ? "0:00" : ""),
    horas_extra: calculo?.extra || (diaDomingo ? "0:00" : ""),
    estado: calculo?.estado || (diaDomingo ? "Incidencia: domingo" : "Entrada registrada — salida pendiente"),
    origen: "Manual admin",
    observaciones: `${existente?.observaciones ? `${existente.observaciones}\n` : ""}Captura/corrección manual por ${req.usuario.correo}. Motivo: ${motivo}${calculo?.nota ? `\n${calculo.nota}` : ""}`
  };
  if (existente) update("asistencia", existente.id, camposRegistro);
  else insert("asistencia", camposRegistro);
  if (tiempos.salida) eliminarPendienteWeb(empresaId, correo, fecha, existente?.id);
  else actualizarPendienteWeb(empresaId, correo, usuario.nombre, usuario.area || "Sin asignar", fecha, fechaTexto, tiempos.entrada, existente?.id);
  actualizarPeriodoLaboral(empresaId, correo, fecha);
  res.json({ ok: true, accion: existente ? "actualizado" : "creado", mensaje: existente ? "La asistencia se corrigió." : "La asistencia se registró." });
});

app.get("/api/admin/periodo-laboral", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const filas = find("periodo_laboral", p => p.empresa_id === empresaId);
  // enriquecer con nombre
  const resultado = filas.map(p => {
    const u = findOne("usuarios", us => us.empresa_id === empresaId && us.correo === p.correo);
    return { ...p, nombre: u ? u.nombre : p.correo };
  });
  res.json({ ok: true, periodos: resultado });
});

app.post("/api/admin/periodo-laboral", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { correo, fecha_ini, fecha_fin, dias_laborales } = req.body || {};
  const correoNormalizado = normalizarCorreo(correo);
  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === correoNormalizado);
  if (!usuarioSeUnioConCodigo(usuario) || !fecha_ini || dias_laborales === undefined) {
    return res.status(400).json({ ok: false, error: "Selecciona a una persona que ya se haya unido con el código del equipo y completa los datos." });
  }
  const inicio = claveFecha(String(fecha_ini));
  const fin = fecha_fin ? claveFecha(String(fecha_fin)) : null;
  const diasLaborales = Number(dias_laborales);
  if (!inicio || !Number.isInteger(diasLaborales) || diasLaborales < 0 || (fin && fin < inicio)) {
    return res.status(400).json({ ok: false, error: "Revisa las fechas y los días laborales del periodo." });
  }

  const existente = findOne("periodo_laboral", p => p.empresa_id === empresaId && p.correo === correoNormalizado);
  if (existente) {
    update("periodo_laboral", existente.id, { fecha_ini: inicio, fecha_fin: fin, dias_laborales: diasLaborales });
  } else {
    insert("periodo_laboral", {
      empresa_id: empresaId, correo: correoNormalizado, fecha_ini: inicio, fecha_fin: fin,
      dias_laborales: diasLaborales, dias_trabajados: 0,
      created_at: new Date().toISOString()
    });
  }
  actualizarPeriodoLaboral(empresaId, correoNormalizado, inicio);
  res.json({ ok: true });
});

app.get("/api/admin/dias-trabajados", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { correo, fecha_ini, fecha_fin } = req.query;
  if (!correo || !fecha_ini || !fecha_fin) {
    return res.status(400).json({ ok: false, error: "Se requieren correo, fecha_ini y fecha_fin" });
  }

  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === normalizarCorreo(correo));
  if (!usuarioSeUnioConCodigo(usuario)) return res.status(404).json({ ok: false, error: "No se encontró una persona que se haya unido con el código del equipo." });
  const filas = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === usuario.correo && a.fecha >= fecha_ini && a.fecha <= fecha_fin
      && a.entrada && a.entrada !== "" && a.salida && a.salida !== "" && a.dia_semana !== 0
  );

  const dias = [...new Set(filas.map(f => f.fecha))].sort();
  res.json({ ok: true, dias_trabajados: dias.length, dias });
});

function calcularAniversario(fechaIngreso, fechaReferencia) {
  const inicio = fechaDesdeEntrada(fechaIngreso);
  const referencia = normalizarFechaSoloDia(fechaReferencia);
  if (!inicio || !referencia || isNaN(inicio.getTime()) || isNaN(referencia.getTime())) return null;
  const aniversarioEn = anio => new Date(anio, inicio.getMonth(), Math.min(inicio.getDate(), new Date(anio, inicio.getMonth() + 1, 0).getDate()), 12);
  let anios = referencia.getFullYear() - inicio.getFullYear();
  if (aniversarioEn(referencia.getFullYear()) > referencia) anios--;
  const siguiente = aniversarioEn(inicio.getFullYear() + Math.max(0, anios) + 1);
  const diasRestantes = Math.max(0, Math.round((siguiente - referencia) / 86400000));
  return { anios: Math.max(0, anios), siguiente: claveFecha(siguiente), diasRestantes };
}

app.get("/api/admin/dias-trabajados-resumen", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const hoy = claveFecha(new Date());
  const usuarios = find("usuarios", u => u.empresa_id === empresaId && u.activo === 1 && usuarioSeUnioConCodigo(u))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));
  const resumen = usuarios.map(u => {
    const fechaIngreso = fechaISOValida(u.fecha_ingreso) ? u.fecha_ingreso : null;
    const aniversario = fechaIngreso ? calcularAniversario(fechaIngreso, hoy) : null;
    const periodo = findOne("periodo_laboral", p => p.empresa_id === empresaId && p.correo === u.correo);
    return {
      correo: u.correo, nombre: u.nombre, area: u.area || "Sin asignar",
      fecha_ingreso: fechaIngreso,
      dias_trabajados: fechaIngreso ? contarDiasSinDomingos(fechaIngreso, hoy) : null,
      fecha_ingreso_pendiente: !fechaIngreso,
      antiguedad_anios: aniversario?.anios || 0,
      proximo_aniversario: aniversario?.siguiente || "",
      dias_para_aniversario: aniversario?.diasRestantes ?? null,
      periodo_laboral: periodo ? { fecha_ini: periodo.fecha_ini, fecha_fin: periodo.fecha_fin, dias_laborales: periodo.dias_laborales } : null
    };
  });
  res.json({ ok: true, hoy, trabajadores: resumen });
});

app.post("/api/admin/vacaciones", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { correo, fecha_ini, fecha_fin, dias_solicitados, motivo } = req.body || {};
  const correoNormalizado = normalizarCorreo(correo);
  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === correoNormalizado);
  const dias = Number(dias_solicitados);
  const inicio = fecha_ini ? claveFecha(String(fecha_ini)) : "";
  const fin = fecha_fin ? claveFecha(String(fecha_fin)) : "";
  if (!usuarioSeUnioConCodigo(usuario) || usuario.rol === "admin" || !inicio || !fin || fin < inicio || !Number.isInteger(dias) || dias <= 0) {
    return res.status(400).json({ ok: false, error: "Revisa el trabajador, las fechas y los días solicitados." });
  }
  insert("vacaciones_solicitud", {
    empresa_id: empresaId, correo: correoNormalizado, fecha_ini: inicio, fecha_fin: fin,
    dias_solicitados: dias, motivo: String(motivo || "").trim(),
    estado: "pendiente", creado_en: new Date().toISOString()
  });
  res.json({ ok: true });
});

app.post("/api/admin/vacaciones/:id/estado", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const solicitud = findOne("vacaciones_solicitud", v => v.id === Number(req.params.id) && v.empresa_id === empresaId);
  if (!solicitud) return res.status(404).json({ ok: false, error: "Solicitud de vacaciones no encontrada." });
  const estado = String(req.body?.estado || "").toLowerCase();
  if (!["pendiente", "aprobada", "rechazada"].includes(estado)) {
    return res.status(400).json({ ok: false, error: "Estado de solicitud no válido." });
  }
  update("vacaciones_solicitud", solicitud.id, { estado, revisado_en: new Date().toISOString() });
  res.json({ ok: true });
});

app.get("/api/admin/vacaciones", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const filas = find("vacaciones_solicitud", v => v.empresa_id === empresaId)
    .sort((a, b) => new Date(b.creado_en) - new Date(a.creado_en));
  const resultado = filas.map(v => {
    const u = findOne("usuarios", us => us.empresa_id === empresaId && us.correo === v.correo);
    return { ...v, nombre: u ? u.nombre : v.correo };
  });
  res.json({ ok: true, solicitudes: resultado });
});

// =============================================
// HELPERS PENDIENTES
// =============================================

function actualizarPendienteWeb(empresaId, correo, nombre, area, fecha, fechaTexto, hora, asistenciaId) {
  const existente = findOne(
    "pendientes",
    p => p.empresa_id === empresaId && p.correo === correo && p.fecha === fecha
  );

  if (existente) {
    update("pendientes", existente.id, {
      hora_entrada: hora, marca_texto: fechaTexto + " " + hora,
      titulo_dia: "ASISTENCIA — " + fechaTexto, nombre, area
    });
  } else {
    insert("pendientes", {
      empresa_id: empresaId, correo, nombre, area,
      titulo_dia: "ASISTENCIA — " + fechaTexto, fecha,
      hora_entrada: hora, marca_texto: fechaTexto + " " + hora
    });
  }
}

function eliminarPendienteWeb(empresaId, correo, fecha, asistenciaId) {
  const p = findOne("pendientes", x => x.empresa_id === empresaId && x.correo === correo && x.fecha === fecha);
  if (p) remove("pendientes", p.id);
}

function actualizarPeriodoLaboral(empresaId, correo, fecha) {
  const existente = findOne("periodo_laboral", p => p.empresa_id === empresaId && p.correo === correo);
  if (!existente) return;
  const desde = existente.fecha_ini || "0000-00-00";
  const hasta = existente.fecha_fin || claveFecha(new Date());
  const dias = new Set(find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === correo && a.fecha >= desde && a.fecha <= hasta
      && a.entrada && a.entrada !== "" && a.salida && a.salida !== "" && a.dia_semana !== 0
  ).map(a => a.fecha));
  update("periodo_laboral", existente.id, { dias_trabajados: dias.size });
}

/*********************************************************************************
 * EXPORTACIÓN A EXCEL (.xlsx) — Admin / RH
 *********************************************************************************/

function formatearFechaExcel(fechaTexto) {
  if (!fechaTexto) return "";
  const f = normalizarFechaSoloDia(fechaTexto);
  if (!f) return fechaTexto;
  const dia = String(f.getDate()).padStart(2, "0");
  const mes = String(f.getMonth() + 1).padStart(2, "0");
  const anio = f.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

async function exportarXlsx(res, nombreArchivo, filas, columnas) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Reporte");

  sheet.columns = columnas.map(c => ({ header: c.header, key: c.key, width: c.width || 20 }));
  sheet.views = [{ state: "frozen", ySplit: 1, showGridLines: false }];
  sheet.properties.defaultRowHeight = 21;
  const header = sheet.getRow(1);
  header.height = 28;
  header.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF990000" } };
  header.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  header.eachCell(cell => {
    cell.border = {
      top: { style: "thin", color: { argb: "FF7A0000" } },
      bottom: { style: "thin", color: { argb: "FF7A0000" } },
      left: { style: "thin", color: { argb: "FFB84141" } },
      right: { style: "thin", color: { argb: "FFB84141" } }
    };
  });

  filas.forEach(f => {
    const fila = {};
    columnas.forEach(c => {
      fila[c.key] = c.render ? c.render(f) : (f[c.key] !== undefined ? f[c.key] : "");
    });
    const row = sheet.addRow(fila);
    row.height = 22;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font = { name: "Arial", size: 10, color: { argb: "FF0B1C30" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "thin", color: { argb: "FFDCE4F5" } } };
      if (row.number % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F8FF" } };
    });
    const estado = normalizar(String(f.estado || ""));
    if (estado.includes("falta") || estado.includes("rechaz") || estado.includes("incidencia")) {
      row.eachCell({ includeEmpty: true }, cell => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4E4" } }; });
    } else if (estado.includes("pendiente") || estado.includes("sin registro")) {
      row.eachCell({ includeEmpty: true }, cell => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4D6" } }; });
    } else if (estado.includes("extra")) {
      row.eachCell({ includeEmpty: true }, cell => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF0C2" } }; });
    }
  });

  if (columnas.length && filas.length) {
    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: filas.length + 1, column: columnas.length } };
  }
  sheet.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${nombreArchivo}"`);
  await workbook.xlsx.write(res);
  res.end();
}

async function exportarReporteSemanalExcel(res, nombreArchivo, empresa, desde, hasta, filas, totales) {
  const workbook = new ExcelJS.Workbook();
  const hoja = workbook.addWorksheet("Reporte semanal");
  const dias = Array.from({ length: 6 }, (_, i) => {
    const fecha = fechaDesdeEntrada(desde);
    fecha.setDate(fecha.getDate() + i);
    return { fecha: claveFecha(fecha), titulo: `${NOMBRES_DIA[fecha.getDay()].toUpperCase()}\n${formatearFechaExcel(fecha)}` };
  });
  const anchos = [28, 32, 16, 21, 21, 21, 21, 21, 21, 17, 15, 13, 11];
  hoja.columns = anchos.map(width => ({ width }));
  hoja.properties.defaultRowHeight = 24;
  hoja.views = [{ state: "frozen", ySplit: 4, showGridLines: false }];

  const titulo = hoja.addRow(["REPORTE SEMANAL DE ASISTENCIA"]);
  hoja.mergeCells(1, 1, 1, 13);
  titulo.height = 32;
  titulo.getCell(1).font = { name: "Arial", size: 17, bold: true, color: { argb: "FFFFFFFF" } };
  titulo.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF990000" } };
  titulo.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

  const subtitulo = hoja.addRow([`SEMANA DEL ${formatearFechaExcel(desde)} AL ${formatearFechaExcel(hasta)}  ·  EQUIPO: ${empresa.nombre || ""}`]);
  hoja.mergeCells(2, 1, 2, 13);
  subtitulo.height = 24;
  subtitulo.getCell(1).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF0B1C30" } };
  subtitulo.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE9FF" } };
  subtitulo.getCell(1).alignment = { vertical: "middle", indent: 1 };

  const indicadores = hoja.addRow([]);
  [[1, 3, `PERSONAL ACTIVO\n${filas.length}`], [4, 6, `HORAS TRABAJADAS\n${totales.trabajadas}`],
    [7, 9, `HORAS EXTRA\n${totales.extra}`], [10, 13, `ASISTENCIAS  ·  FALTAS\n${totales.asistencias}  ·  ${totales.faltas}`]]
    .forEach(([inicio, fin, texto]) => {
      hoja.mergeCells(indicadores.number, inicio, indicadores.number, fin);
      const cell = indicadores.getCell(inicio);
      cell.value = texto;
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF0B1C30" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF4FF" } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = { top: { style: "thin", color: { argb: "FFDCE4F5" } }, bottom: { style: "thin", color: { argb: "FFDCE4F5" } }, left: { style: "thin", color: { argb: "FFDCE4F5" } }, right: { style: "thin", color: { argb: "FFDCE4F5" } } };
    });
  indicadores.height = 38;

  const encabezados = ["Trabajador", "Correo", "Área", ...dias.map(d => d.titulo), "Horas trabajadas", "Horas extra", "Asistencias", "Faltas"];
  const encabezado = hoja.addRow(encabezados);
  encabezado.height = 38;
  encabezado.eachCell(cell => {
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = { top: { style: "thin", color: { argb: "FF7A0000" } }, bottom: { style: "thin", color: { argb: "FF7A0000" } }, left: { style: "thin", color: { argb: "FFB84141" } }, right: { style: "thin", color: { argb: "FFB84141" } } };
  });

  const bordes = { bottom: { style: "thin", color: { argb: "FFDCE4F5" } }, left: { style: "thin", color: { argb: "FFDCE4F5" } }, right: { style: "thin", color: { argb: "FFDCE4F5" } } };
  filas.forEach((persona, indice) => {
    const detalles = new Map((persona.detalle || []).map(d => [d.fecha, d]));
    const valoresDias = dias.map(dia => {
      const d = detalles.get(dia.fecha);
      if (!d) return { texto: "—", color: "FFF5F8FF" };
      const estado = normalizar(d.estado || "");
      if (estado.includes("antes del ingreso")) return { texto: "ANTES DEL INGRESO", color: "FFEFF4FF" };
      if (!d.entrada) {
        const falta = estado.includes("falta");
        return { texto: falta ? "FALTA\nSin asistencia" : "PENDIENTE\nSin registro", color: falta ? "FFF4CCCC" : "FFFCE4D6" };
      }
      const salida = d.salida || "Salida pendiente";
      const extra = aMinutos(d.extra) > 0 ? `  ·  +${d.extra} extra` : "";
      const color = estado.includes("incidencia") || estado.includes("rechaz") ? "FFF4CCCC"
        : aMinutos(d.extra) > 0 || estado.includes("extra") ? "FFFFE6B8" : "FFD9EAD3";
      return { texto: `${d.entrada} → ${salida}\n${d.horas || "0:00"} h${extra}`, color };
    });
    const row = hoja.addRow([
      persona.nombre, persona.correo, persona.area,
      ...valoresDias.map(d => d.texto), persona.trabajadas, persona.extra, persona.asistencias, persona.faltas
    ]);
    row.height = 42;
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font = { name: "Arial", size: 10, color: { argb: "FF0B1C30" } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = bordes;
      if (indice % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFF" } };
    });
    row.getCell(1).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    row.getCell(2).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    valoresDias.forEach((d, i) => {
      const cell = row.getCell(i + 4);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: d.color } };
    });
    if (aMinutos(persona.extra) > 0) row.getCell(11).font = { name: "Arial", size: 10, bold: true, color: { argb: "FFB45309" } };
    if (Number(persona.faltas) > 0) row.getCell(13).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF991B1B" } };
  });

  const filaTotal = hoja.addRow(["TOTAL GENERAL", "", "", "", "", "", "", "", "", totales.trabajadas, totales.extra, totales.asistencias, totales.faltas]);
  hoja.mergeCells(filaTotal.number, 1, filaTotal.number, 3);
  filaTotal.height = 30;
  filaTotal.eachCell({ includeEmpty: true }, cell => {
    cell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1C30" } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = bordes;
  });

  if (filas.length) hoja.autoFilter = { from: { row: 4, column: 1 }, to: { row: filas.length + 4, column: encabezados.length } };
  hoja.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${nombreArchivo}"`);
  await workbook.xlsx.write(res);
  res.end();
}

/* Exportar usuarios */
app.get("/api/admin/exportar/usuarios", requireAdmin, async (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const filas = find("usuarios", u => u.empresa_id === empresaId).sort((a, b) => a.nombre.localeCompare(b.nombre));

  const columnas = [
    { header: "Nombre", key: "nombre", width: 30 },
    { header: "Correo", key: "correo", width: 35 },
    { header: "Área", key: "area", width: 15 },
    { header: "Comida (min)", key: "minutos_comida", width: 15 },
    { header: "Activo", key: "activo", width: 10 },
    { header: "Fecha ingreso", key: "fecha_ingreso", width: 15, render: f => formatearFechaExcel(f.fecha_ingreso) },
    { header: "Fecha alta", key: "fecha_alta", width: 20 }
  ];

  await exportarXlsx(res, "usuarios.xlsx", filas, columnas);
});

/* Exportar asistencia semanal de un usuario */
app.get("/api/admin/exportar/asistencia", requireAdmin, async (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { correo, fecha } = req.query;
  if (!correo || !fecha) {
    return res.status(400).json({ ok: false, error: "Se requieren correo y fecha" });
  }
  const correoNormalizado = normalizarCorreo(correo);
  const usuario = findOne("usuarios", u => u.empresa_id === empresaId && u.correo === correoNormalizado);
  if (!usuarioSeUnioConCodigo(usuario)) return res.status(404).json({ ok: false, error: "Selecciona a una persona que ya se haya unido con el código del equipo." });

  const baseFecha = fechaDesdeEntrada(fecha);
  const lunes = lunesDeLaSemana(baseFecha);
  const sabado = new Date(lunes.getTime()); sabado.setDate(sabado.getDate() + 5);
  const desde = claveFecha(lunes);
  const hasta = claveFecha(sabado);

  const filas = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.correo === correoNormalizado && a.fecha >= desde && a.fecha <= hasta
  ).sort((a, b) => a.fecha.localeCompare(b.fecha));

  const columnas = [
    { header: "Fecha", key: "fecha", width: 15, render: f => formatearFechaExcel(f.fecha) },
    { header: "Día", key: "nombre_dia", width: 12, render: f => NOMBRES_DIA[f.dia_semana] || "" },
    { header: "Entrada", key: "entrada", width: 12 },
    { header: "Salida", key: "salida", width: 12 },
    { header: "Horas brutas", key: "horas_brutas", width: 15 },
    { header: "Comida", key: "comida", width: 12 },
    { header: "Horas trabajadas", key: "horas_trabajadas", width: 18 },
    { header: "Horas extra", key: "horas_extra", width: 15 },
    { header: "Estado", key: "estado", width: 25 },
    { header: "Observaciones", key: "observaciones", width: 40 }
  ];

  await exportarXlsx(res, `asistencia_${correo}_${desde}_${hasta}.xlsx`, filas, columnas);
});

/* Exportar reporte semanal general */
app.get("/api/admin/exportar/reporte", requireAdmin, async (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { fecha } = req.query;
  const baseFecha = fecha ? fechaDesdeEntrada(fecha) : new Date();
  const lunes = lunesDeLaSemana(baseFecha);
  const sabado = new Date(lunes.getTime()); sabado.setDate(sabado.getDate() + 5);
  const desde = claveFecha(lunes);
  const hasta = claveFecha(sabado);

  const usuarios = find("usuarios", u => u.empresa_id === empresaId && u.activo === 1).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const filasAsistencia = find(
    "asistencia",
    a => a.empresa_id === empresaId && a.fecha >= desde && a.fecha <= hasta
  );
  const mapa = new Map();
  filasAsistencia.forEach(f => {
    if (!mapa.has(f.correo)) mapa.set(f.correo, new Map());
    mapa.get(f.correo).set(f.fecha, f);
  });
  const primeraAsistenciaPorCorreo = new Map();
  find("asistencia", f => f.empresa_id === empresaId && f.entrada && f.fecha <= hasta).forEach(f => {
    const actual = primeraAsistenciaPorCorreo.get(f.correo);
    if (!actual || f.fecha < actual) primeraAsistenciaPorCorreo.set(f.correo, f.fecha);
  });

  const filas = usuarios.map(u => {
    let trabajadas = 0, extra = 0, asistencias = 0, faltas = 0;
    const diasUsuario = mapa.get(u.correo) || new Map();
    const fechaAlta = fechaISOValida(u.fecha_ingreso) ? u.fecha_ingreso : (primeraAsistenciaPorCorreo.get(u.correo) || desde);
    const detalle = [];
    for (let i = 0; i < 6; i++) {
      const dia = new Date(lunes.getTime());
      dia.setDate(dia.getDate() + i);
      const claveDia = claveFecha(dia);
      if (claveDia < fechaAlta) {
        detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: "Antes del ingreso" });
        continue;
      }
      const f = diasUsuario.get(claveDia);
      if (f?.estado === "Falta") {
        faltas++;
        detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: "Falta", entrada: "", salida: "", horas: "0:00", extra: "0:00" });
        continue;
      }
      if (!f?.entrada) {
        const falta = usuarioDebeHaberTerminado(u, dia, empresaId);
        if (falta) faltas++;
        detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: falta ? "Falta" : "Pendiente", entrada: "", salida: "", horas: "0:00", extra: "0:00" });
        continue;
      }
      asistencias++;
      trabajadas += aMinutos(f.horas_trabajadas) || 0;
      const minutosExtraDia = aMinutos(f.horas_extra) || 0;
      if (dia.getDay() !== 6) extra += minutosExtraDia;
      detalle.push({ fecha: claveDia, dia: NOMBRES_DIA[dia.getDay()], estado: f.estado || "Asistencia",
        entrada: f.entrada || "", salida: f.salida || "", horas: f.horas_trabajadas || "0:00", extra: aTextoHoras(minutosExtraDia) });
    }
    const lunesAnterior = new Date(lunes.getTime());
    lunesAnterior.setDate(lunesAnterior.getDate() - 7);
    const sabadoAnterior = new Date(lunesAnterior.getTime());
    sabadoAnterior.setDate(sabadoAnterior.getDate() + 5);
    const previa = findOne("asistencia", f => f.empresa_id === empresaId && f.correo === u.correo
      && f.fecha === claveFecha(sabadoAnterior) && f.dia_semana === 6);
    if (previa) extra += aMinutos(previa.horas_extra) || 0;
    return {
      nombre: u.nombre,
      correo: u.correo,
      area: u.area,
      trabajadas: aTextoHoras(trabajadas),
      extra: aTextoHoras(extra),
      asistencias,
      faltas,
      detalle
    };
  });
  const totales = filas.reduce((t, fila) => {
    t.trabajadas += aMinutos(fila.trabajadas) || 0;
    t.extra += aMinutos(fila.extra) || 0;
    t.asistencias += fila.asistencias;
    t.faltas += fila.faltas;
    return t;
  }, { trabajadas: 0, extra: 0, asistencias: 0, faltas: 0 });
  const empresa = findOne("empresa", e => e.id === empresaId) || {};
  await exportarReporteSemanalExcel(res, `reporte_semanal_${desde}_${hasta}.xlsx`, empresa, desde, hasta, filas, {
    trabajadas: aTextoHoras(totales.trabajadas), extra: aTextoHoras(totales.extra),
    asistencias: totales.asistencias, faltas: totales.faltas
  });
});

function agregarHojaRespaldo(workbook, nombre, columnas, filas) {
  let nombreHoja = nombre.slice(0, 31);
  let sufijo = 2;
  while (workbook.getWorksheet(nombreHoja)) {
    const marca = ` ${sufijo++}`;
    nombreHoja = `${nombre.slice(0, 31 - marca.length)}${marca}`;
  }
  const hoja = workbook.addWorksheet(nombreHoja);
  hoja.columns = columnas.map(c => ({ header: c.header, key: c.key, width: c.width || 18 }));
  hoja.views = [{ state: "frozen", ySplit: 1, showGridLines: false }];
  hoja.properties.defaultRowHeight = 21;
  const encabezado = hoja.getRow(1);
  encabezado.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
  encabezado.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF990000" } };
  encabezado.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  encabezado.height = 28;
  encabezado.eachCell(cell => {
    cell.border = { top: { style: "thin", color: { argb: "FF7A0000" } }, bottom: { style: "thin", color: { argb: "FF7A0000" } }, left: { style: "thin", color: { argb: "FFB84141" } }, right: { style: "thin", color: { argb: "FFB84141" } } };
  });
  filas.forEach((fila, indice) => {
    const row = hoja.addRow(fila);
    row.height = 21;
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font = { name: "Arial", size: 10, color: { argb: "FF0B1C30" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "thin", color: { argb: "FFDCE4F5" } } };
      if (indice % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F8FF" } };
    });
    const estado = normalizar(String(fila.estado || ""));
    const color = estado.includes("falta") || estado.includes("rechaz") || estado.includes("incidencia") ? "FFFCE4E4"
      : estado.includes("pendiente") || estado.includes("sin registro") ? "FFFCE4D6"
      : estado.includes("extra") ? "FFFFF0C2" : null;
    if (color) row.eachCell({ includeEmpty: true }, cell => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } }; });
  });
  if (columnas.length) {
    hoja.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, filas.length + 1), column: columnas.length } };
  }
  return nombreHoja;
}

function agregarHojaHistorialAsistencia(workbook, registros) {
  const base = "Sistema - Asistencias";
  let nombre = base;
  let sufijo = 2;
  while (workbook.getWorksheet(nombre)) nombre = `${base} ${sufijo++}`;

  const hoja = workbook.addWorksheet(nombre);
  const columnas = [
    { header: "Fecha", width: 13 }, { header: "Trabajador", width: 28 }, { header: "Correo", width: 31 },
    { header: "Área", width: 16 }, { header: "Entrada", width: 11 }, { header: "Salida a comida", width: 15 },
    { header: "Regreso de comida", width: 17 }, { header: "Salida", width: 11 }, { header: "Horas brutas", width: 13 },
    { header: "Comida", width: 11 }, { header: "Horas trabajadas", width: 15 }, { header: "Horas extra", width: 12 },
    { header: "Estado", width: 29 }, { header: "Observaciones", width: 42 }, { header: "GPS", width: 24 },
    { header: "Origen", width: 12 }, { header: "ID", width: 9 }
  ];
  hoja.columns = columnas.map(c => ({ width: c.width }));
  hoja.properties.defaultRowHeight = 21;
  hoja.views = [{ state: "normal", showGridLines: false }];

  const porFecha = new Map();
  registros.slice().sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "") || (a.area || "").localeCompare(b.area || "") || (a.nombre || "").localeCompare(b.nombre || ""))
    .forEach(registro => {
      const fecha = registro.fecha || "Sin fecha";
      if (!porFecha.has(fecha)) porFecha.set(fecha, new Map());
      const porArea = porFecha.get(fecha);
      const area = registro.area || "Sin asignar";
      if (!porArea.has(area)) porArea.set(area, []);
      porArea.get(area).push(registro);
    });

  const bordes = {
    top: { style: "thin", color: { argb: "FFD5D5D5" } }, bottom: { style: "thin", color: { argb: "FFD5D5D5" } },
    left: { style: "thin", color: { argb: "FFD5D5D5" } }, right: { style: "thin", color: { argb: "FFD5D5D5" } }
  };
  let primerDia = true;
  for (const [fecha, areas] of porFecha) {
    if (!primerDia) {
      const espacio = hoja.addRow([]);
      espacio.height = 8;
    }
    primerDia = false;

    const fechaTexto = (areas.values().next().value || [])[0]?.fecha_texto || (fecha === "Sin fecha" ? "Sin fecha" : formatearFechaExcel(fecha));
    const bandaFecha = hoja.addRow([`ASISTENCIA — ${fechaTexto}`]);
    hoja.mergeCells(bandaFecha.number, 1, bandaFecha.number, columnas.length);
    bandaFecha.height = 27;
    bandaFecha.getCell(1).font = { name: "Arial", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
    bandaFecha.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF990000" } };
    bandaFecha.getCell(1).alignment = { horizontal: "center", vertical: "middle" };

    const encabezado = hoja.addRow(columnas.map(c => c.header));
    encabezado.height = 31;
    encabezado.eachCell(cell => {
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = bordes;
    });

    const areasOrdenadas = [...areas.entries()].sort((a, b) => {
      const orden = { "oficina": 0, "tecnico": 1, "sin asignar": 2 };
      return (orden[normalizar(a[0])] ?? 9) - (orden[normalizar(b[0])] ?? 9) || a[0].localeCompare(b[0]);
    });
    for (const [area, filasArea] of areasOrdenadas) {
      const primerRegistro = filasArea[0];
      const empresaId = primerRegistro?.empresa_id;
      const areaBase = findOne("areas", a => a.empresa_id === empresaId && normalizar(a.nombre) === normalizar(area));
      const horario = primerRegistro?.fecha
        ? obtenerHorarioParaFecha(area, fechaDesdeEntrada(primerRegistro.fecha), "", empresaId)
        : areaBase ? { inicio: areaBase.entrada, fin: areaBase.salida } : HORARIOS["Sin asignar"];
      const bandaArea = hoja.addRow([`${area.toLocaleUpperCase("es-MX")}  ·  Horario ${horario.inicio} a ${horario.fin}`]);
      hoja.mergeCells(bandaArea.number, 1, bandaArea.number, columnas.length);
      bandaArea.height = 22;
      bandaArea.getCell(1).font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      bandaArea.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD96868" } };
      bandaArea.getCell(1).alignment = { vertical: "middle" };

      filasArea.forEach(registro => {
        const fila = hoja.addRow([
          registro.fecha_texto || fechaTexto,
          registro.nombre || "",
          registro.correo || "",
          registro.area || "",
          registro.entrada || "",
          registro.comida_entrada || "",
          registro.comida_salida || "",
          registro.salida || "",
          registro.horas_brutas || "",
          registro.comida || "",
          registro.horas_trabajadas || "",
          registro.horas_extra || "",
          registro.estado || "",
          registro.observaciones || "",
          registro.latitud != null && registro.longitud != null ? `${registro.latitud}, ${registro.longitud}` : "",
          registro.origen || "",
          registro.id || ""
        ]);
        fila.height = registro.observaciones ? 30 : 23;
        const estado = normalizar(String(registro.estado || ""));
        const extra = aMinutos(registro.horas_extra) || 0;
        const color = estado.includes("falta") || estado.includes("rechaz") || estado.includes("incidencia") || estado.includes("sin registro") ? "FFF4CCCC"
          : estado.includes("pendiente") || (!registro.salida && registro.entrada) ? "FFF4CCCC"
          : estado.includes("extra") || extra > 0 ? "FFFFE6B8" : "FFD9EAD3";
        fila.eachCell({ includeEmpty: true }, cell => {
          cell.font = { name: "Arial", size: 10, color: { argb: "FF202020" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
          cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
          cell.border = bordes;
        });
        fila.getCell(2).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        fila.getCell(3).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        fila.getCell(13).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF202020" } };
        fila.getCell(14).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      });
    }
  }

  hoja.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return nombre;
}

app.get("/api/admin/exportar/completo", requireAdmin, async (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const empresa = findOne("empresa", e => e.id === empresaId);
  if (!empresa) return res.status(404).json({ ok: false, error: "Empresa no encontrada." });

  const workbook = new ExcelJS.Workbook();
  let archivoOriginalIncluido = false;
  let hojasOriginales = 0;
  if (empresa.archivo_horas_original) {
    const rutaOriginal = path.resolve(__dirname, empresa.archivo_horas_original);
    const carpetaSegura = path.resolve(DIRECTORIO_DATOS) + path.sep;
    if (!rutaOriginal.startsWith(carpetaSegura) || !fs.existsSync(rutaOriginal)) {
      return res.status(503).json({ ok: false, error: "No está disponible la copia del archivo de horas original; no se generó un respaldo incompleto." });
    }
    try {
      await workbook.xlsx.readFile(rutaOriginal);
      archivoOriginalIncluido = true;
      hojasOriginales = workbook.worksheets.length;
    } catch (error) {
      console.error("No se pudo abrir el archivo de horas original:", error.message);
      return res.status(500).json({ ok: false, error: "No se pudo leer el archivo original de horas." });
    }
  }

  const hojasSistema = [];
  const usuarios = find("usuarios", u => u.empresa_id === empresaId);
  const correosPorId = new Map(usuarios.map(u => [u.correo, u.nombre]));

  hojasSistema.push(agregarHojaHistorialAsistencia(workbook, find("asistencia", a => a.empresa_id === empresaId)));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Usuarios", [
    { header: "ID", key: "id", width: 9 }, { header: "Correo", key: "correo", width: 34 },
    { header: "Nombre", key: "nombre", width: 30 }, { header: "Área", key: "area", width: 18 },
    { header: "Perfil", key: "rol", width: 16 }, { header: "Activo", key: "activo", width: 10 },
    { header: "Minutos de comida", key: "minutos_comida", width: 18 }, { header: "Fecha de ingreso", key: "fecha_ingreso", width: 16 },
    { header: "Fecha de alta", key: "fecha_alta", width: 24 }
  ], usuarios.map(u => ({ id: u.id, correo: u.correo, nombre: u.nombre, area: u.area, rol: u.rol,
    activo: Boolean(u.activo), minutos_comida: u.minutos_comida, fecha_ingreso: u.fecha_ingreso, fecha_alta: u.fecha_alta }))));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Areas", [
    { header: "ID", key: "id", width: 9 }, { header: "Área", key: "nombre", width: 28 },
    { header: "Entrada", key: "entrada", width: 14 }, { header: "Salida", key: "salida", width: 14 },
    { header: "Activa", key: "activo", width: 12 }
  ], find("areas", a => a.empresa_id === empresaId).map(a => ({ ...a, activo: Boolean(a.activo) }))));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Pendientes", [
    { header: "ID", key: "id", width: 9 }, { header: "Correo", key: "correo", width: 34 },
    { header: "Trabajador", key: "nombre", width: 30 }, { header: "Área", key: "area", width: 18 },
    { header: "Fecha", key: "fecha", width: 13 }, { header: "Hora de entrada", key: "hora_entrada", width: 16 },
    { header: "Título", key: "titulo_dia", width: 28 }, { header: "Marca original", key: "marca_texto", width: 24 }
  ], find("pendientes", p => p.empresa_id === empresaId)));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Horarios", [
    { header: "ID", key: "id", width: 9 }, { header: "Correo", key: "correo", width: 34 },
    { header: "Nombre de referencia", key: "nombre_referencia", width: 28 }, { header: "Día", key: "dia", width: 16 },
    { header: "Entrada", key: "entrada", width: 12 }, { header: "Salida", key: "salida", width: 12 },
    { header: "Vigente desde", key: "vigente_desde", width: 15 }, { header: "Vigente hasta", key: "vigente_hasta", width: 15 },
    { header: "Activo", key: "activo", width: 10 }
  ], find("horarios_especiales", h => h.empresa_id === empresaId)));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Periodos", [
    { header: "ID", key: "id", width: 9 }, { header: "Correo", key: "correo", width: 34 },
    { header: "Fecha inicial", key: "fecha_ini", width: 15 }, { header: "Fecha final", key: "fecha_fin", width: 15 },
    { header: "Días laborales", key: "dias_laborales", width: 16 }, { header: "Días trabajados", key: "dias_trabajados", width: 17 },
    { header: "Creado", key: "created_at", width: 24 }
  ], find("periodo_laboral", p => p.empresa_id === empresaId)));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Vacaciones", [
    { header: "ID", key: "id", width: 9 }, { header: "Correo", key: "correo", width: 34 },
    { header: "Trabajador", key: "nombre", width: 30 }, { header: "Fecha inicial", key: "fecha_ini", width: 15 },
    { header: "Fecha final", key: "fecha_fin", width: 15 }, { header: "Días solicitados", key: "dias_solicitados", width: 17 },
    { header: "Estado", key: "estado", width: 15 }, { header: "Motivo", key: "motivo", width: 40 },
    { header: "Creado", key: "creado_en", width: 24 }, { header: "Revisado", key: "revisado_en", width: 24 }
  ], find("vacaciones_solicitud", v => v.empresa_id === empresaId)
    .map(v => ({ ...v, nombre: correosPorId.get(v.correo) || v.correo }))));

  hojasSistema.push(agregarHojaRespaldo(workbook, "Sistema - Oficina", [
    { header: "Empresa", key: "nombre", width: 30 }, { header: "Dirección", key: "direccion", width: 40 },
    { header: "Latitud", key: "lat", width: 14 }, { header: "Longitud", key: "lng", width: 14 },
    { header: "Radio en metros", key: "radio_metros", width: 18 }, { header: "Ubicación habilitada", key: "habilitar_ubicacion", width: 20 }
  ], [{ nombre: empresa.nombre, direccion: empresa.direccion, lat: empresa.lat, lng: empresa.lng,
    radio_metros: empresa.radio_metros, habilitar_ubicacion: Boolean(empresa.habilitar_ubicacion) }]));

  agregarHojaRespaldo(workbook, "Sistema - Exportacion", [
    { header: "Dato", key: "dato", width: 30 }, { header: "Valor", key: "valor", width: 60 }
  ], [
    { dato: "Equipo", valor: empresa.nombre },
    { dato: "Exportado", valor: new Intl.DateTimeFormat("es-MX", { timeZone: obtenerZonaHoraria(), dateStyle: "medium", timeStyle: "short" }).format(new Date()) },
    { dato: "Archivo original incluido", valor: archivoOriginalIncluido ? "Sí" : "No" },
    { dato: "Hojas del archivo original", valor: archivoOriginalIncluido ? hojasOriginales : 0 },
    { dato: "Nota", valor: "Las hojas originales se conservan como histórico. Las hojas Sistema muestran los datos actuales al momento de exportar." },
    { dato: "Hojas actuales del sistema", valor: hojasSistema.join(", ") },
    { dato: "Sesiones y contraseñas", valor: "No se incluyen por seguridad." }
  ]);

  const fechaArchivo = claveFecha(new Date());
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="respaldo_horas_completo_${fechaArchivo}.xlsx"`);
  try {
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("No se pudo generar el respaldo completo:", error.message);
    if (res.headersSent) return res.destroy(error);
    res.status(500).json({ ok: false, error: "No se pudo generar el respaldo completo." });
  }
});

/*********************************************************************************
 * NOTIFICACIONES — Aniversario de trabajo (antigüedad)
 *********************************************************************************/

app.get("/api/admin/notificaciones", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const usuarios = find("usuarios", u => u.empresa_id === empresaId && u.activo === 1);
  const hoy = normalizarFechaSoloDia(new Date());
  const hitos = [1, 5, 10, 15, 20, 25, 30];

  const notificaciones = [];
  const fechasPendientes = usuarios.filter(u => !fechaISOValida(u.fecha_ingreso))
    .map(u => ({ nombre: u.nombre, correo: u.correo }));
  const fechasPorConfirmar = usuarios.filter(u => fechaISOValida(u.fecha_ingreso)
      && String(u.fecha_alta || "").slice(0, 10) === u.fecha_ingreso)
    .map(u => ({ nombre: u.nombre, correo: u.correo, fecha_ingreso: u.fecha_ingreso }));

  usuarios.forEach(u => {
    if (!u.fecha_ingreso) return;
    const fi = normalizarFechaSoloDia(u.fecha_ingreso);
    if (isNaN(fi.getTime())) return;

    const aniversarioEn = anio => new Date(anio, fi.getMonth(), Math.min(fi.getDate(), new Date(anio, fi.getMonth() + 1, 0).getDate()), 12);
    const aniversarioActual = aniversarioEn(hoy.getFullYear());
    const aniosEnAniversarioActual = hoy.getFullYear() - fi.getFullYear();
    const esAniversarioHoy = claveFecha(aniversarioActual) === claveFecha(hoy) && aniosEnAniversarioActual > 0;
    if (esAniversarioHoy && hitos.includes(aniosEnAniversarioActual)) {
      notificaciones.push({
        tipo: "aniversario_hito",
        correo: u.correo,
        nombre: u.nombre,
        area: u.area,
        fecha_ingreso: u.fecha_ingreso,
        años: aniosEnAniversarioActual,
        dias: 0,
        mensaje: `${u.nombre} cumple ${aniosEnAniversarioActual} año(s) de antigüedad en la empresa.`
      });
    }

    const proximoAnio = Math.max(1, aniosEnAniversarioActual + (aniversarioActual > hoy ? 0 : 1));
    const proximoAniversario = aniversarioEn(fi.getFullYear() + proximoAnio);
    const diffDias = Math.max(0, Math.round((proximoAniversario - hoy) / 86400000));
    if (hitos.includes(proximoAnio) && diffDias > 0 && diffDias <= 30) {
      notificaciones.push({
        tipo: "aniversario_proximo", correo: u.correo, nombre: u.nombre, area: u.area,
        fecha_ingreso: u.fecha_ingreso, años: proximoAnio, dias: diffDias,
        mensaje: `${u.nombre} cumple ${proximoAnio} año(s) de antigüedad en ${diffDias} día(s).`
      });
    }
  });

  notificaciones.sort((a, b) => (a.dias || 0) - (b.dias || 0));
  res.json({ ok: true, notificaciones, fechasPendientes, fechasPorConfirmar });
});

/*********************************************************************************
 * IMPORTACIÓN DESDE CSV (hoja Asistencia exportada)
 * Formato esperado:
 *   fila 2: ASISTENCIA — DD/MM/YYYY
 *   fila 3: Trabajador,Área,Entrada,Salida,Horas trabajadas,Horas extra,Estado
 *   fila 4: OFICINA  ·  Horario 09:00 a 19:00   (grupo / área)
 *   filas siguientes: Nombre,Área,Entrada,Salida,...
 *********************************************************************************/

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

function parsearFechaCsvDDMMYYYY(texto) {
  const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(texto));
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0, 0);
  if (d.getFullYear() === Number(m[3]) && d.getMonth() === Number(m[2]) - 1 && d.getDate() === Number(m[1])) return d;
  return null;
}

function normalizarNombreCsv(texto) {
  return normalizar(texto || "")
    .replace(/\s*estad[ií]as\s*/gi, "")
    .trim();
}

function fechaIngresoDesdeExcel(valor) {
  const celda = valorCeldaExcel(valor);
  if (celda instanceof Date && !isNaN(celda.getTime())) {
    return String(celda.getUTCFullYear()).padStart(4, "0") + "-" + String(celda.getUTCMonth() + 1).padStart(2, "0") + "-" + String(celda.getUTCDate()).padStart(2, "0");
  }
  if (typeof celda === "number" && Number.isFinite(celda) && celda >= 1) {
    const fecha = new Date(Date.UTC(1899, 11, 30) + Math.floor(celda) * 86400000);
    return String(fecha.getUTCFullYear()).padStart(4, "0") + "-" + String(fecha.getUTCMonth() + 1).padStart(2, "0") + "-" + String(fecha.getUTCDate()).padStart(2, "0");
  }
  const texto = String(celda || "").trim();
  if (fechaISOValida(texto)) return texto;
  const visible = parsearFechaCsvDDMMYYYY(texto);
  return visible ? String(visible.getFullYear()).padStart(4, "0") + "-" + String(visible.getMonth() + 1).padStart(2, "0") + "-" + String(visible.getDate()).padStart(2, "0") : null;
}

function enriquecerUsuarioDesdeExcel(usuario, datosImportados = {}) {
  if (!usuario) return usuario;
  const cambios = {};
  const fechaIngreso = fechaIngresoDesdeExcel(datosImportados.fecha_ingreso);
  if (fechaIngreso && (!fechaISOValida(usuario.fecha_ingreso) || usuario._importado)) cambios.fecha_ingreso = fechaIngreso;
  if ((!usuario.area || usuario.area === "Sin asignar") && datosImportados.area) cambios.area = datosImportados.area;
  if (usuario._importado && Number.isFinite(datosImportados.minutos_comida)) cambios.minutos_comida = datosImportados.minutos_comida;
  if (Object.keys(cambios).length) update("usuarios", usuario.id, cambios);
  return usuario;
}

function buscarOCrearUsuarioPorNombre(empresaId, nombreRaw, datosImportados = {}) {
  const nombreLimpio = String(nombreRaw || "").trim();
  if (!nombreLimpio) return { usuario: null, creado: false };

  const clave = normalizarNombreCsv(nombreLimpio);
  const correoImportado = String(datosImportados.correo || "").trim().toLowerCase();

  if (correoImportado) {
    const usuarioPorCorreo = findOne("usuarios", us =>
      us.empresa_id === empresaId && String(us.correo || "").toLowerCase() === correoImportado
    );
    if (usuarioPorCorreo) return { usuario: enriquecerUsuarioDesdeExcel(usuarioPorCorreo, datosImportados), creado: false };
  }

  // 1) Buscar coincidencia exacta normalizada
  let u = findOne("usuarios", us =>
    us.empresa_id === empresaId &&
    normalizar(us.nombre) === clave
  );

  if (u) return { usuario: enriquecerUsuarioDesdeExcel(u, datosImportados), creado: false };

  // 2) Buscar que el nombre del CSV esté contenido en el nombre del usuario o viceversa
  const todos = find("usuarios", us => us.empresa_id === empresaId && us.activo === 1);
  for (const us of todos) {
    const un = normalizar(us.nombre);
    if (un.includes(clave) || clave.includes(un)) return { usuario: enriquecerUsuarioDesdeExcel(us, datosImportados), creado: false };
  }

  // 3) Crear usuario provisional con correo generado
  const correoGen = correoImportado || (clave.replace(/\s+/g, ".") + "@import.csv.local");
  const yaExiste = findOne("usuarios", us => us.empresa_id === empresaId && us.correo === correoGen);
  if (yaExiste) return { usuario: yaExiste, creado: false };

  const nuevo = insert("usuarios", {
    empresa_id: empresaId,
    correo: correoGen,
    nombre: nombreLimpio,
    area: datosImportados.area || "Sin asignar",
    activo: datosImportados.activo === false ? 0 : 1,
    minutos_comida: datosImportados.minutos_comida ?? MINUTOS_COMIDA_POR_DEFECTO,
    fecha_ingreso: fechaIngresoDesdeExcel(datosImportados.fecha_ingreso),
    fecha_alta: new Date().toISOString(),
    _importado: true
  });
  return { usuario: nuevo, creado: true };
}

app.post("/api/admin/importar/asistencia-csv", requireAdmin, upload.single("archivo"), async (req, res) => {
  const empresaId = req.usuario.empresa_id;

  if (!req.file) {
    return res.status(400).json({ ok: false, error: "No se subió ningún archivo." });
  }

  const texto = req.file.buffer.toString("utf8");
  const lineas = texto.split(/\r?\n/).map(l => l.replace(/\r$/, ""));

  let fechaActual = null;
  let areaActual = null;
  let importados = 0;
  let usuariosCreados = 0;
  const errores = [];

  for (let i = 0; i < lineas.length; i++) {
    const linea = lineas[i];
    if (linea.trim() === "") continue;

    // Detectar título de día: ASISTENCIA — DD/MM/YYYY
    const mDia = /^ASISTENCIA\s*—\s*(\d{1,2}\/\d{1,2}\/\d{4})/i.exec(linea);
    if (mDia) {
      fechaActual = parsearFechaCsvDDMMYYYY(mDia[1]);
      areaActual = null;
      continue;
    }

    // Detectar grupo/área: OFICINA · Horario ...  o  TÉCNICOS · Horario ...
    const mGrupo = /^(OFICINA|TÉCNICOS|Tecnicos|SIN ÁREA|Sin área)\s*·/i.exec(linea);
    if (mGrupo) {
      const g = normalizar(mGrupo[1]);
      if (g.includes("oficina")) areaActual = "Oficina";
      else if (g.includes("tecnico")) areaActual = "Técnico";
      else areaActual = "Sin asignar";
      continue;
    }

    // Detectar fila de encabezado
    if (/^Trabajador/i.test(linea)) continue;

    // Separar columnas por coma (CSV simple)
    const cols = linea.split(",").map(c => c.trim());

    const nombreRaw = cols[0];
    if (!nombreRaw || nombreRaw === "") continue;

    // Saltar si parece encabezado repetido o grupo
    if (normalizar(nombreRaw) === "trabajador") continue;

    const areaCsv = cols[1] || areaActual || "Sin asignar";
    const entradaRaw = cols[2] || "";
    const salidaRaw = cols[3] || "";
    const horasTrabajadasRaw = cols[4] || "";
    const horasExtraRaw = cols[5] || "";
    const estadoRaw = cols[6] || "";

    if (!fechaActual) {
      errores.push(`Fila ${i + 1}: ${nombreRaw} — no se detectó fecha de día`);
      continue;
    }

    const fechaTexto = claveFecha(fechaActual);
    const fechaTextoMostrar = fechaTextoVisual(fechaActual);

    const resultadoUsuario = buscarOCrearUsuarioPorNombre(empresaId, nombreRaw);
    if (!resultadoUsuario.usuario) {
      errores.push(`Fila ${i + 1}: ${nombreRaw} — no se pudo resolver usuario`);
      continue;
    }
    const usuario = resultadoUsuario.usuario;
    if (resultadoUsuario.creado) usuariosCreados++;

    const entrada = normalizarHoraManual(entradaRaw) || "";
    const salida = normalizarHoraManual(salidaRaw) || "";

    // Evitar duplicados exactos
    const yaExiste = findOne("asistencia", a =>
      a.empresa_id === empresaId &&
      a.correo === usuario.correo &&
      a.fecha === fechaTexto &&
      a.entrada === entrada
    );
    if (yaExiste) continue;

    const esDomingo = esDomingoPorFecha(fechaTexto);
    const esSabado = esSabadoPorFecha(fechaTexto);
    const diaSemana = fechaActual.getDay();

    let horasBrutas = "", comida = "", horasTrabajadas = "", horasExtra = "", estado = estadoRaw || "Completado";

    if (entrada && salida && !esDomingo) {
      const minutosComida = obtenerMinutosComidaUsuario(usuario);
      const cal = calcularHoras(entrada, salida, usuario.area, fechaActual, minutosComida, usuario.correo, empresaId);
      horasBrutas = cal.brutas;
      comida = cal.comida;
      horasTrabajadas = cal.trabajadas;
      horasExtra = cal.extra;
      estado = cal.estado;
    } else if (esDomingo) {
      horasBrutas = "0:00"; comida = "0:00"; horasTrabajadas = "0:00"; horasExtra = "0:00";
      estado = "Incidencia: domingo";
    } else if (entrada && !salida) {
      horasBrutas = ""; comida = ""; horasTrabajadas = ""; horasExtra = "";
      estado = "Entrada registrada — salida pendiente";
    }

    insert("asistencia", {
      empresa_id: empresaId,
      correo: usuario.correo,
      nombre: usuario.nombre,
      area: usuario.area || areaCsv,
      fecha: fechaTexto,
      fecha_texto: fechaTextoMostrar,
      dia_semana: diaSemana,
      entrada,
      salida,
      horas_brutas: horasBrutas,
      comida,
      horas_trabajadas: horasTrabajadas,
      horas_extra: horasExtra,
      estado,
      observaciones: "Importado desde CSV hoja Asistencia.",
      origen: "CSV"
    });

    // Si quedó sin salida, agregar pendiente
    if (entrada && !salida && !esDomingo) {
      const yaPendiente = findOne("pendientes", p =>
        p.empresa_id === empresaId && p.correo === usuario.correo && p.fecha === fechaTexto
      );
      if (!yaPendiente) {
        insert("pendientes", {
          empresa_id: empresaId,
          correo: usuario.correo,
          nombre: usuario.nombre,
          area: usuario.area,
          titulo_dia: "ASISTENCIA — " + fechaTexto,
          fecha: fechaTexto,
          hora_entrada: entrada,
          marca_texto: fechaTextoMostrar + " " + entrada
        });
      }
    }

    importados++;
  }

  find("periodo_laboral", p => p.empresa_id === empresaId)
    .forEach(p => actualizarPeriodoLaboral(empresaId, p.correo));

  res.json({
    ok: true,
    importados,
    usuariosCreados,
    errores: errores.slice(0, 50),
    totalErrores: errores.length
  });
});

function valorCeldaExcel(valor) {
  if (valor && typeof valor === "object" && !(valor instanceof Date)) {
    if (Object.prototype.hasOwnProperty.call(valor, "result")) return valor.result;
    if (Array.isArray(valor.richText)) return valor.richText.map(parte => parte.text || "").join("");
    if (Object.prototype.hasOwnProperty.call(valor, "text")) return valor.text;
  }
  return valor;
}

function horaDesdeCeldaExcel(valor) {
  const v = valorCeldaExcel(valor);
  if (v instanceof Date && !isNaN(v.getTime())) {
    // ExcelJS representa las horas como fechas de 1899 en UTC.
    return `${String(v.getUTCHours()).padStart(2, "0")}:${String(v.getUTCMinutes()).padStart(2, "0")}`;
  }
  return normalizarHoraManual(v) || "";
}

function duracionDesdeCeldaExcel(valor) {
  const v = valorCeldaExcel(valor);
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date && !isNaN(v.getTime())) {
    return `${String(v.getUTCHours()).padStart(2, "0")}:${String(v.getUTCMinutes()).padStart(2, "0")}`;
  }
  if (typeof v === "number" && v >= 0 && v < 1) {
    const totalMinutos = Math.round(v * 24 * 60);
    return `${String(Math.floor(totalMinutos / 60)).padStart(2, "0")}:${String(totalMinutos % 60).padStart(2, "0")}`;
  }
  return normalizarHoraManual(v) || String(v).trim();
}

function minutosComidaDesdeExcel(valor) {
  const v = valorCeldaExcel(valor);
  if (v instanceof Date && !isNaN(v.getTime())) return v.getUTCHours() * 60 + v.getUTCMinutes();
  if (typeof v === "number" && v >= 0 && v < 1) return Math.round(v * 24 * 60);
  if (v !== null && v !== undefined && String(v).trim()) return normalizarMinutosComida(v);
  return MINUTOS_COMIDA_POR_DEFECTO;
}

function crearMapaUsuariosDesdeExcel(workbook) {
  const mapa = new Map();
  const hoja = workbook.getWorksheet("Usuarios");
  if (!hoja) return mapa;

  let encabezado = null;
  hoja.eachRow({ includeEmpty: false }, (fila, numero) => {
    if (encabezado) return;
    const nombres = fila.values.slice(1).map(v => normalizar(String(valorCeldaExcel(v) || "")));
    if (nombres.includes("correo") && nombres.includes("nombre")) {
      encabezado = { numero, columnas: nombres };
    }
  });
  if (!encabezado) return mapa;

  const indice = nombre => encabezado.columnas.indexOf(nombre) + 1;
  const colCorreo = indice("correo"), colNombre = indice("nombre");
  const colArea = indice("area"), colActivo = indice("activo");
  const colComida = indice("tiempo de comida");
  const colFechaIngreso = encabezado.columnas.findIndex(nombre =>
    nombre === "fecha de ingreso" || nombre.startsWith("fecha de ingreso ")
    || nombre === "fecha ingreso" || nombre.startsWith("fecha ingreso ")
    || nombre === "ingreso" || nombre === "fecha de alta laboral"
  ) + 1;
  hoja.eachRow({ includeEmpty: false }, (fila, numero) => {
    if (numero <= encabezado.numero) return;
    const nombre = String(valorCeldaExcel(fila.getCell(colNombre).value) || "").trim();
    const correo = String(valorCeldaExcel(fila.getCell(colCorreo).value) || "").trim().toLowerCase();
    if (!nombre || !correo.includes("@")) return;
    const activoTexto = normalizar(String(valorCeldaExcel(fila.getCell(colActivo).value) || "si"));
    mapa.set(normalizarNombreCsv(nombre), {
      correo,
      area: String(valorCeldaExcel(fila.getCell(colArea).value) || "").trim() || "Sin asignar",
      activo: !["no", "inactivo", "false", "0"].includes(activoTexto),
      minutos_comida: minutosComidaDesdeExcel(fila.getCell(colComida).value),
      fecha_ingreso: colFechaIngreso ? fechaIngresoDesdeExcel(fila.getCell(colFechaIngreso).value) : null
    });
  });
  return mapa;
}

function extraerAsistenciasDesdeExcel(hoja) {
  const filas = [];
  let fechaActual = null;
  let areaActual = null;

  hoja.eachRow({ includeEmpty: false }, (fila, numero) => {
    const valores = fila.values.slice(1).map(valorCeldaExcel);
    const primeraCelda = String(valores[0] ?? "").trim();
    if (!primeraCelda) return;

    const tituloDia = /^ASISTENCIA\s*[—–-]\s*(\d{1,2}\/\d{1,2}\/\d{4})/i.exec(primeraCelda);
    if (tituloDia) {
      fechaActual = parsearFechaCsvDDMMYYYY(tituloDia[1]);
      areaActual = null;
      return;
    }

    if (/^Trabajador$/i.test(primeraCelda)) return;

    const grupo = /^(OFICINA|TÉCNICOS|TECNICOS|SIN ÁREA|SIN AREA)\s*·/i.exec(primeraCelda);
    if (grupo) {
      const nombreGrupo = normalizar(grupo[1]);
      areaActual = nombreGrupo.includes("oficina") ? "Oficina"
        : nombreGrupo.includes("tecnico") ? "Técnico" : "Sin asignar";
      return;
    }

    const nombre = primeraCelda;
    if (normalizar(nombre) === "trabajador") return;
    filas.push({
      fila: numero,
      fecha: fechaActual,
      nombre,
      area: String(valores[1] ?? "").trim() || areaActual || "Sin asignar",
      entrada: valores[2],
      salida: valores[3],
      horasTrabajadas: valores[4],
      horasExtra: valores[5],
      estado: String(valores[6] ?? "").trim()
    });
  });
  return filas;
}

function guardarCopiaExcelPorEmpresa(empresaId, buffer) {
  const carpeta = path.join(DIRECTORIO_DATOS, "horas-importadas");
  fs.mkdirSync(carpeta, { recursive: true });
  const nombre = `equipo-${empresaId}-${Date.now()}-${crypto.randomBytes(5).toString("hex")}.xlsx`;
  const destino = path.join(carpeta, nombre);
  const temporal = `${destino}.tmp`;
  try {
    fs.writeFileSync(temporal, buffer, { flag: "wx" });
    fs.renameSync(temporal, destino);
  } catch (error) {
    try { if (fs.existsSync(temporal)) fs.unlinkSync(temporal); } catch (_) {}
    throw error;
  }
  return path.relative(__dirname, destino).split(path.sep).join("/");
}

function importarAsistenciasExcel(empresaId, filas, usuariosExcel) {
  let importados = 0;
  let usuariosCreados = 0;
  let duplicados = 0;
  const errores = [];

  for (const fila of filas) {
    if (!fila.fecha) {
      errores.push(`Fila ${fila.fila}: ${fila.nombre} — no se detectó una fecha de día.`);
      continue;
    }

    const fechaTexto = claveFecha(fila.fecha);
    const fechaTextoMostrar = fechaTextoVisual(fila.fecha);
    const usuarioExcel = usuariosExcel.get(normalizarNombreCsv(fila.nombre)) || {
      area: fila.area
    };
    const resultadoUsuario = buscarOCrearUsuarioPorNombre(empresaId, fila.nombre, usuarioExcel);
    if (!resultadoUsuario.usuario) {
      errores.push(`Fila ${fila.fila}: ${fila.nombre} — no se pudo resolver el usuario.`);
      continue;
    }
    const usuario = resultadoUsuario.usuario;
    if (resultadoUsuario.creado) usuariosCreados++;

    const entrada = horaDesdeCeldaExcel(fila.entrada);
    const salida = horaDesdeCeldaExcel(fila.salida);
    const yaExiste = findOne("asistencia", a =>
      a.empresa_id === empresaId && a.correo === usuario.correo &&
      a.fecha === fechaTexto && a.entrada === entrada
    );
    if (yaExiste) {
      duplicados++;
      continue;
    }

    const esDomingo = esDomingoPorFecha(fechaTexto);
    const fecha = fechaDesdeEntrada(fechaTexto);
    let calculo = null;
    if (entrada && salida && !esDomingo) {
      calculo = calcularHoras(entrada, salida, usuario.area || fila.area, fecha,
        obtenerMinutosComidaUsuario(usuario), usuario.correo, empresaId);
    }

    let horasBrutas = calculo?.brutas || "";
    let comida = calculo?.comida || "";
    let horasTrabajadas = duracionDesdeCeldaExcel(fila.horasTrabajadas) || calculo?.trabajadas || "";
    let horasExtra = duracionDesdeCeldaExcel(fila.horasExtra) || calculo?.extra || "";
    let estado = fila.estado || calculo?.estado || "Completado";

    if (esDomingo && !entrada && !salida) {
      horasBrutas = comida = horasTrabajadas = horasExtra = "0:00";
    }
    if (entrada && !salida && !esDomingo && !fila.estado) {
      estado = "Entrada registrada — salida pendiente";
    }

    insert("asistencia", {
      empresa_id: empresaId,
      correo: usuario.correo,
      nombre: usuario.nombre,
      area: usuario.area || fila.area,
      fecha: fechaTexto,
      fecha_texto: fechaTextoMostrar,
      dia_semana: fecha.getDay(),
      entrada,
      salida,
      horas_brutas: horasBrutas,
      comida,
      horas_trabajadas: horasTrabajadas,
      horas_extra: horasExtra,
      estado,
      observaciones: "Importado desde Excel, hoja Asistencia.",
      origen: "Excel"
    });

    if (entrada && !salida && !esDomingo) {
      const yaPendiente = findOne("pendientes", p =>
        p.empresa_id === empresaId && p.correo === usuario.correo && p.fecha === fechaTexto
      );
      if (!yaPendiente) {
        insert("pendientes", {
          empresa_id: empresaId,
          correo: usuario.correo,
          nombre: usuario.nombre,
          area: usuario.area,
          titulo_dia: "ASISTENCIA — " + fechaTexto,
          fecha: fechaTexto,
          hora_entrada: entrada,
          marca_texto: fechaTextoMostrar + " " + entrada
        });
      }
    }
    importados++;
  }

  find("periodo_laboral", p => p.empresa_id === empresaId)
    .forEach(p => actualizarPeriodoLaboral(empresaId, p.correo));

  return { importados, usuariosCreados, duplicados, errores: errores.slice(0, 50), totalErrores: errores.length };
}

app.post("/api/admin/importar/asistencia-excel", requireAdmin, upload.single("archivo"), async (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: "Selecciona un archivo Excel .xlsx." });
  if (path.extname(req.file.originalname).toLowerCase() !== ".xlsx") {
    return res.status(400).json({ ok: false, error: "Este importador acepta libros Excel .xlsx." });
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(req.file.buffer);
  } catch (error) {
    return res.status(400).json({ ok: false, error: "No se pudo abrir el archivo .xlsx. Verifica que sea un libro de Excel válido." });
  }

  const hoja = workbook.getWorksheet("Asistencia");
  if (!hoja) {
    return res.status(400).json({ ok: false, error: "No encontré la pestaña «Asistencia» en el libro." });
  }

  const filas = extraerAsistenciasDesdeExcel(hoja);
  if (!filas.length) {
    return res.status(400).json({ ok: false, error: "La pestaña «Asistencia» no contiene filas para importar." });
  }

  const empresaId = req.usuario.empresa_id;
  let archivoGuardado;
  try {
    archivoGuardado = guardarCopiaExcelPorEmpresa(empresaId, req.file.buffer);
    const resultado = importarAsistenciasExcel(empresaId, filas, crearMapaUsuariosDesdeExcel(workbook));
    const empresa = findOne("empresa", e => e.id === empresaId);
    if (!empresa) throw new Error("No se encontró el equipo para asociar el libro importado.");
    empresa.archivo_horas_original = archivoGuardado;
    guardarDB();
    return res.json({ ok: true, archivo: req.file.originalname, libroCompletoGuardado: true, ...resultado });
  } catch (error) {
    console.error("No se pudo importar o conservar el libro Excel:", error.message);
    if (archivoGuardado) {
      try { fs.unlinkSync(path.resolve(__dirname, archivoGuardado)); } catch (_) {}
    }
    return res.status(500).json({ ok: false, error: "No se pudo completar la importación y guardar el libro original." });
  }
});

/*********************************************************************************
 * CONFIGURAR UBICACIÓN DE LA EMPRESA (opcional)
 *********************************************************************************/

// Configuración de ubicación de oficina (para el frontend - cualquier usuario autenticado)
app.get("/api/config/oficina", requireAuth, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const oficina = obtenerOficinaConfig(empresaId);
  res.json({ ok: true, oficina });
});

// Configuración de ubicación de oficina (admin)
app.get("/api/admin/config/oficina", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const oficina = obtenerOficinaConfig(empresaId);
  res.json({ ok: true, oficina });
});

app.post("/api/admin/config/oficina", requireAdmin, (req, res) => {
  const empresaId = req.usuario.empresa_id;
  const { lat, lng, radio_metros, direccion, habilitar_ubicacion } = req.body || {};

  const empresa = findOne("empresa", e => e.id === empresaId);
  if (!empresa) return res.status(404).json({ ok: false, error: "Empresa no encontrada" });

  const habilitada = habilitar_ubicacion === true || habilitar_ubicacion === "true";
  const latNumero = lat === "" || lat === null || lat === undefined ? null : Number(lat);
  const lngNumero = lng === "" || lng === null || lng === undefined ? null : Number(lng);
  const radioNumero = radio_metros === "" || radio_metros === null || radio_metros === undefined ? 100 : Number(radio_metros);
  if (habilitada && (!Number.isFinite(latNumero) || latNumero < -90 || latNumero > 90 || !Number.isFinite(lngNumero) || lngNumero < -180 || lngNumero > 180)) {
    return res.status(400).json({ ok: false, error: "Selecciona un punto válido en el mapa antes de habilitar la ubicación." });
  }
  if (habilitada && (!Number.isFinite(radioNumero) || radioNumero < 10 || radioNumero > 50000)) {
    return res.status(400).json({ ok: false, error: "El rango debe estar entre 10 y 50,000 metros." });
  }

  const cambios = {};
  if (lat !== undefined) cambios.lat = lat === "" || lat === null ? null : latNumero;
  if (lng !== undefined) cambios.lng = lng === "" || lng === null ? null : lngNumero;
  if (radio_metros !== undefined) cambios.radio_metros = radioNumero;
  if (direccion !== undefined) cambios.direccion = direccion === "" || direccion === null ? null : String(direccion).trim();
  if (habilitar_ubicacion !== undefined) cambios.habilitar_ubicacion = habilitada;

  update("empresa", empresaId, cambios);
  res.json({ ok: true, oficina: obtenerOficinaConfig(empresaId) });
});

// =============================================
// SUPERADMIN Y CONTROL DE SUSCRIPCIONES
// =============================================
const COOKIE_SUPERADMIN = "jjm_superadmin_session";
const ARCHIVO_SUPERADMIN_SETUP = path.join(DIRECTORIO_DATOS, "superadmin-bootstrap-token.txt");
const intentosSuperadmin = new Map();
let tokenSetupSuperadmin = process.env.SUPERADMIN_SETUP_TOKEN || "";
let tokenSetupGenerado = false;

function prepararTokenSetupSuperadmin() {
  if (db.superadmins.some(superadmin => superadmin.activo === 1)) {
    tokenSetupSuperadmin = "";
    return;
  }
  if (tokenSetupSuperadmin) return;
  if (USAR_SUPABASE) {
    throw new Error("Configura SUPERADMIN_SETUP_TOKEN como secreto antes del primer arranque en Supabase.");
  }
  try {
    fs.mkdirSync(DIRECTORIO_DATOS, { recursive: true });
    if (fs.existsSync(ARCHIVO_SUPERADMIN_SETUP)) {
      tokenSetupSuperadmin = fs.readFileSync(ARCHIVO_SUPERADMIN_SETUP, "utf8").trim();
      tokenSetupGenerado = Boolean(tokenSetupSuperadmin);
    }
    if (!tokenSetupSuperadmin) {
      tokenSetupSuperadmin = crypto.randomBytes(32).toString("base64url");
      fs.writeFileSync(ARCHIVO_SUPERADMIN_SETUP, tokenSetupSuperadmin, { encoding: "utf8", mode: 0o600 });
      tokenSetupGenerado = true;
    }
  } catch (error) {
    console.error("No se pudo preparar el acceso inicial de superadmin:", error.message);
  }
}

function obtenerSuperadminSesion(req) {
  try {
    const token = req.cookies[COOKIE_SUPERADMIN];
    if (!token) return null;
    const sesion = findOne("superadmin_sesiones", s => s.token_hash === hashToken(token));
    if (!sesion || new Date(sesion.expira_en).getTime() <= Date.now()) return null;
    const superadmin = findOne("superadmins", s => s.id === sesion.superadmin_id && s.activo === 1);
    return superadmin ? { id: superadmin.id, correo: superadmin.correo, nombre: superadmin.nombre } : null;
  } catch (error) { return null; }
}

function requireSuperadmin(req, res, next) {
  const superadmin = obtenerSuperadminSesion(req);
  if (!superadmin) return res.status(401).json({ ok: false, error: "Inicia sesión como superadmin." });
  req.superadmin = superadmin;
  next();
}

function crearSesionSuperadmin(res, superadmin) {
  const token = crypto.randomBytes(32).toString("base64url");
  insert("superadmin_sesiones", {
    token_hash: hashToken(token), superadmin_id: superadmin.id,
    expira_en: new Date(Date.now() + DIAS_SESION_MS).toISOString()
  });
  res.cookie(COOKIE_SUPERADMIN, token, {
    httpOnly: true, secure: process.env.NODE_ENV === "production",
    sameSite: cookieSameSite, path: "/", maxAge: DIAS_SESION_MS
  });
}

function registrarAuditoriaSuperadmin(superadmin, accion, empresaId, detalle = {}) {
  insert("auditoria_superadmin", {
    superadmin_id: superadmin.id, superadmin_correo: superadmin.correo,
    empresa_id: empresaId || null, accion, detalle,
    creado_en: new Date().toISOString()
  });
}

function datosEmpresaSuperadmin(empresa) {
  const administrador = findOne("usuarios", u => u.empresa_id === empresa.id && u.rol === "admin" && u.activo === 1);
  const miembros = contarCuposOcupados(empresa.id);
  const pagos = find("pagos_suscripcion", p => p.empresa_id === empresa.id)
    .sort((a, b) => String(b.registrado_en).localeCompare(String(a.registrado_en)));
  return {
    id: empresa.id, nombre: empresa.nombre, creado_en: empresa.creado_en,
    administrador: administrador ? { nombre: administrador.nombre, correo: administrador.correo } : null,
    estado: estadoSuscripcion(empresa), solicitud_estado: empresa.solicitud_codigo_estado || "sin_solicitud",
    solicitud_en: empresa.solicitud_codigo_en || null, cupos_solicitados: Number(empresa.cupos_solicitados || 0),
    cupos_limite: Number(empresa.cupos_limite || 0), cupos_ocupados: miembros,
    codigo_equipo: suscripcionActiva(empresa) ? empresa.codigo : null,
    monto_pagado_mxn: Number(empresa.monto_pagado_mxn || 0),
    monto_ciclo_mxn: empresa.monto_ciclo_mxn === undefined ? null : Number(empresa.monto_ciclo_mxn),
    periodicidad: empresa.periodicidad || "mensual",
    fecha_inicio_suscripcion: empresa.fecha_inicio_suscripcion || null,
    fecha_fin_suscripcion: empresa.fecha_fin_suscripcion || null,
    archivada_en: empresa.archivada_en || null,
    ultimo_pago: pagos[0] || null, pagos
  };
}

app.get("/api/superadmin/setup/status", (req, res) => {
  res.json({ ok: true, configurado: db.superadmins.some(s => s.activo === 1), setup_disponible: !db.superadmins.some(s => s.activo === 1) && Boolean(tokenSetupSuperadmin) });
});

app.post("/api/superadmin/setup", async (req, res) => {
  if (db.superadmins.some(s => s.activo === 1) || !tokenSetupSuperadmin) {
    return res.status(403).json({ ok: false, error: "La configuración inicial ya se utilizó o no está disponible." });
  }
  const { token, nombre, correo, password } = req.body || {};
  const recibido = String(token || "");
  const bufferRecibido = Buffer.from(recibido);
  const bufferEsperado = Buffer.from(tokenSetupSuperadmin);
  if (!recibido || bufferRecibido.length !== bufferEsperado.length
      || !crypto.timingSafeEqual(bufferRecibido, bufferEsperado)) {
    return res.status(401).json({ ok: false, error: "La clave inicial no coincide con la que muestra la consola del servidor." });
  }
  if (!String(nombre || "").trim() || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizarCorreo(correo))) {
    return res.status(400).json({ ok: false, error: "Escribe tu nombre y un correo válido." });
  }
  if (typeof password !== "string" || password.length < 12) {
    return res.status(400).json({ ok: false, error: "Crea una contraseña de al menos 12 caracteres." });
  }
  const passwordHash = await guardarPassword(password);
  if (db.superadmins.some(s => s.activo === 1)) return res.status(409).json({ ok: false, error: "La configuración inicial ya se utilizó." });
  const superadmin = insert("superadmins", {
    nombre: String(nombre).trim(), correo: normalizarCorreo(correo),
    password_hash: passwordHash, activo: 1,
    creado_en: new Date().toISOString()
  });
  tokenSetupSuperadmin = "";
  if (fs.existsSync(ARCHIVO_SUPERADMIN_SETUP)) fs.rmSync(ARCHIVO_SUPERADMIN_SETUP, { force: true });
  crearSesionSuperadmin(res, superadmin);
  registrarAuditoriaSuperadmin({ id: superadmin.id, correo: superadmin.correo }, "superadmin_creado", null);
  res.json({ ok: true, superadmin: { nombre: superadmin.nombre, correo: superadmin.correo } });
});

app.post("/api/superadmin/login", async (req, res) => {
  const ip = req.ip || req.socket?.remoteAddress || "unknown";
  const intento = intentosSuperadmin.get(ip);
  if (intento && intento.hasta > Date.now() && intento.cantidad >= 8) {
    return res.status(429).json({ ok: false, error: "Demasiados intentos. Espera 15 minutos e inténtalo de nuevo." });
  }
  const correo = normalizarCorreo(req.body?.correo);
  const superadmin = findOne("superadmins", s => s.correo === correo && s.activo === 1);
  if (!superadmin || !await verificarPassword(String(req.body?.password || ""), superadmin.password_hash)) {
    const actual = intento && intento.hasta > Date.now() ? intento : { cantidad: 0, hasta: Date.now() + 15 * 60 * 1000 };
    actual.cantidad += 1;
    intentosSuperadmin.set(ip, actual);
    return res.status(401).json({ ok: false, error: "Correo o contraseña incorrectos." });
  }
  intentosSuperadmin.delete(ip);
  crearSesionSuperadmin(res, superadmin);
  res.json({ ok: true, superadmin: { nombre: superadmin.nombre, correo: superadmin.correo } });
});

app.get("/api/superadmin/sesion", requireSuperadmin, (req, res) => res.json({ ok: true, superadmin: req.superadmin }));

app.post("/api/superadmin/logout", (req, res) => {
  const token = req.cookies[COOKIE_SUPERADMIN];
  if (token) {
    const sesion = findOne("superadmin_sesiones", s => s.token_hash === hashToken(token));
    if (sesion) remove("superadmin_sesiones", sesion.id);
  }
  res.clearCookie(COOKIE_SUPERADMIN, { path: "/", sameSite: cookieSameSite, secure: process.env.NODE_ENV === "production" });
  res.json({ ok: true });
});

app.get("/api/superadmin/equipos", requireSuperadmin, (req, res) => {
  const equipos = db.empresa.filter(e => !e.archivada_en).map(datosEmpresaSuperadmin)
    .sort((a, b) => String(b.creado_en || "").localeCompare(String(a.creado_en || "")));
  const resumen = {
    pendientes: equipos.filter(e => ["pendiente", "declinada"].includes(e.estado) || String(e.solicitud_estado || "").startsWith("solicitada")).length,
    activos: equipos.filter(e => e.estado === "activa").length,
    por_vencer: equipos.filter(e => e.estado === "activa" && e.fecha_fin_suscripcion && (new Date(e.fecha_fin_suscripcion).getTime() - Date.now()) <= 7 * 86400000 && new Date(e.fecha_fin_suscripcion).getTime() >= Date.now()).length,
    equipos, auditoria: db.auditoria_superadmin.slice(-12).reverse()
  };
  res.json({ ok: true, ...resumen });
});

app.post("/api/superadmin/equipos/:id/activar", requireSuperadmin, async (req, res) => {
  const empresa = findOne("empresa", e => e.id === Number(req.params.id) && !e.archivada_en);
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  const cupos = Number(req.body?.cupos_limite);
  const montoPagado = Number(req.body?.monto_pagado_mxn || 0);
  const montoCiclo = Number(req.body?.monto_ciclo_mxn);
  const periodicidad = ["mensual", "anual"].includes(req.body?.periodicidad) ? req.body.periodicidad : "mensual";
  const inicio = String(req.body?.fecha_inicio_suscripcion || claveFecha(new Date()));
  const fin = String(req.body?.fecha_fin_suscripcion || "");
  if (!Number.isInteger(cupos) || cupos < 1 || cupos > 10000) return res.status(400).json({ ok: false, error: "Los cupos deben ser un número entre 1 y 10,000." });
  if (cupos < contarCuposOcupados(empresa.id)) return res.status(400).json({ ok: false, error: `No puedes reducir a ${cupos}: el equipo ya tiene ${contarCuposOcupados(empresa.id)} trabajadores activos.` });
  if (!Number.isFinite(montoPagado) || montoPagado < 0 || !Number.isFinite(montoCiclo) || montoCiclo < 0) return res.status(400).json({ ok: false, error: "Revisa los montos capturados." });
  if (!fechaISOValida(inicio) || !fechaISOValida(fin) || fin <= inicio) return res.status(400).json({ ok: false, error: "Selecciona un periodo de suscripción válido." });
  const codigo = empresa.codigo || generarCodigoUnico();
  Object.assign(empresa, {
    codigo, estado_suscripcion: "activa", solicitud_codigo_estado: "aprobada",
    cupos_limite: cupos, monto_ciclo_mxn: montoCiclo, periodicidad,
    fecha_inicio_suscripcion: inicio, fecha_fin_suscripcion: fin,
    archivada_en: null
  });
  if (montoPagado > 0) {
    empresa.monto_pagado_mxn = Number(empresa.monto_pagado_mxn || 0) + montoPagado;
    insert("pagos_suscripcion", {
      empresa_id: empresa.id, monto_mxn: montoPagado, periodicidad,
      periodo_desde: inicio, periodo_hasta: fin,
      nota: String(req.body?.nota || "Pago verificado y activación aprobada.").trim(),
      registrado_por: req.superadmin.correo, registrado_en: new Date().toISOString()
    });
  }
  guardarDB();
  registrarAuditoriaSuperadmin(req.superadmin, "equipo_activado", empresa.id, { cupos, montoPagado, montoCiclo, periodicidad, codigo });
  notificarAdministradores(empresa.id, {
    clave: `suscripcion-activada-${empresa.id}-${fin}`,
    titulo: "Tu equipo ya está activo",
    cuerpo: `Tu código para invitar al equipo está listo. Tienes ${cupos} cupos hasta el ${fin}.`,
    url: "/admin.html"
  });
  // Espera a que Supabase persista ANTES de responder: así cualquier otro
  // dispositivo o instancia que consulte la API ve el código activo de inmediato.
  await esperarPersistenciaDB();
  res.json({ ok: true, equipo: datosEmpresaSuperadmin(empresa) });
});

app.post("/api/superadmin/equipos/:id/codigo", requireSuperadmin, async (req, res) => {
  const empresa = findOne("empresa", e => e.id === Number(req.params.id) && !e.archivada_en);
  if (!empresa || !suscripcionActiva(empresa)) return res.status(409).json({ ok: false, error: "Activa la membresía antes de generar un código." });
  empresa.codigo = generarCodigoUnico();
  guardarDB();
  registrarAuditoriaSuperadmin(req.superadmin, "codigo_regenerado", empresa.id);
  notificarAdministradores(empresa.id, {
    clave: `codigo-renovado-${empresa.id}-${Date.now()}`, titulo: "Código de invitación actualizado",
    cuerpo: "Tu administrador puede consultar el nuevo código desde Código y notificaciones.", url: "/admin.html"
  });
  await esperarPersistenciaDB();
  res.json({ ok: true, codigo_equipo: empresa.codigo });
});

app.post("/api/superadmin/equipos/:id/pagos", requireSuperadmin, async (req, res) => {
  const empresa = findOne("empresa", e => e.id === Number(req.params.id) && !e.archivada_en);
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  const monto = Number(req.body?.monto_mxn);
  const fin = String(req.body?.fecha_fin_suscripcion || "");
  if (!Number.isFinite(monto) || monto <= 0) return res.status(400).json({ ok: false, error: "El monto pagado debe ser mayor que cero." });
  if (!fechaISOValida(fin) || fin <= claveFecha(new Date())) return res.status(400).json({ ok: false, error: "La nueva fecha de vencimiento debe ser posterior a hoy." });
  const desde = empresa.fecha_fin_suscripcion && empresa.fecha_fin_suscripcion > claveFecha(new Date()) ? empresa.fecha_fin_suscripcion : claveFecha(new Date());
  insert("pagos_suscripcion", {
    empresa_id: empresa.id, monto_mxn: monto, periodicidad: empresa.periodicidad || "mensual",
    periodo_desde: desde, periodo_hasta: fin,
    nota: String(req.body?.nota || "Pago de renovación verificado.").trim(),
    registrado_por: req.superadmin.correo, registrado_en: new Date().toISOString()
  });
  empresa.monto_pagado_mxn = Number(empresa.monto_pagado_mxn || 0) + monto;
  empresa.fecha_inicio_suscripcion = empresa.fecha_inicio_suscripcion || desde;
  empresa.fecha_fin_suscripcion = fin;
  empresa.estado_suscripcion = "activa";
  if (!empresa.codigo) empresa.codigo = generarCodigoUnico();
  empresa.solicitud_codigo_estado = "aprobada";
  guardarDB();
  registrarAuditoriaSuperadmin(req.superadmin, "pago_registrado", empresa.id, { monto, fecha_fin_suscripcion: fin });
  notificarAdministradores(empresa.id, {
    clave: `renovacion-${empresa.id}-${fin}`, titulo: "Membresía renovada",
    cuerpo: `El pago quedó registrado. Tu servicio está activo hasta el ${fin}.`, url: "/admin.html"
  });
  await esperarPersistenciaDB();
  res.json({ ok: true, equipo: datosEmpresaSuperadmin(empresa) });
});

app.post("/api/superadmin/equipos/:id/declinar", requireSuperadmin, async (req, res) => {
  const empresa = findOne("empresa", e => e.id === Number(req.params.id) && !e.archivada_en);
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  empresa.estado_suscripcion = "declinada";
  empresa.solicitud_codigo_estado = "declinada";
  empresa.codigo = null;
  empresa.motivo_decline = String(req.body?.motivo || "Solicitud declinada.").trim();
  guardarDB();
  registrarAuditoriaSuperadmin(req.superadmin, "equipo_declinado", empresa.id, { motivo: empresa.motivo_decline });
  await esperarPersistenciaDB();
  res.json({ ok: true });
});

app.delete("/api/superadmin/equipos/:id", requireSuperadmin, async (req, res) => {
  const empresa = findOne("empresa", e => e.id === Number(req.params.id) && !e.archivada_en);
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  empresa.archivada_en = new Date().toISOString();
  empresa.estado_suscripcion = "archivada";
  empresa.codigo = null;
  guardarDB();
  registrarAuditoriaSuperadmin(req.superadmin, "equipo_archivado", empresa.id, { nombre: empresa.nombre });
  await esperarPersistenciaDB();
  res.json({ ok: true });
});

app.get("/api/superadmin/equipos/:id/auditoria", requireSuperadmin, (req, res) => {
  const empresa = findOne("empresa", e => e.id === Number(req.params.id));
  if (!empresa) return res.status(404).json({ ok: false, error: "Equipo no encontrado." });
  const pagos = find("pagos_suscripcion", p => p.empresa_id === empresa.id).sort((a, b) => String(b.registrado_en).localeCompare(String(a.registrado_en)));
  const auditoria = find("auditoria_superadmin", a => a.empresa_id === empresa.id).sort((a, b) => String(b.creado_en).localeCompare(String(a.creado_en)));
  res.json({ ok: true, pagos, auditoria });
});

// =============================================
// NOTIFICACIONES PUSH PARA WEB Y MÓVIL
// =============================================
let WEB_PUSH_PUBLIC_KEY = "";
function configurarWebPush() {
  try {
    fs.mkdirSync(DIRECTORIO_DATOS, { recursive: true });
    const archivo = path.join(DIRECTORIO_DATOS, "web-push-vapid.json");
    let claves;
    if (process.env.VAPID_PUBLIC_KEY && process.env.VAPID_PRIVATE_KEY) {
      claves = { publicKey: process.env.VAPID_PUBLIC_KEY, privateKey: process.env.VAPID_PRIVATE_KEY };
    } else if (fs.existsSync(archivo)) {
      claves = JSON.parse(fs.readFileSync(archivo, "utf8"));
    } else {
      claves = webpush.generateVAPIDKeys();
      fs.writeFileSync(archivo, JSON.stringify(claves, null, 2), { encoding: "utf8", mode: 0o600 });
    }
    webpush.setVapidDetails(process.env.VAPID_SUBJECT || "mailto:admin@example.com", claves.publicKey, claves.privateKey);
    WEB_PUSH_PUBLIC_KEY = claves.publicKey;
  } catch (error) {
    console.error("No se pudo configurar Web Push:", error.message);
  }
}
configurarWebPush();

function subscriptionValida(subscription) {
  try {
    const url = new URL(subscription?.endpoint || "");
    const host = url.hostname.toLowerCase();
    const hostPermitido = host === "fcm.googleapis.com" || host.endsWith(".fcm.googleapis.com")
      || host === "push.services.mozilla.com" || host.endsWith(".push.services.mozilla.com")
      || host === "updates.push.services.mozilla.com" || host === "web.push.apple.com" || host.endsWith(".push.apple.com")
      || host === "notify.windows.com" || host.endsWith(".notify.windows.com");
    return url.protocol === "https:" && hostPermitido
      && typeof subscription?.keys?.p256dh === "string" && subscription.keys.p256dh.length > 20
      && typeof subscription?.keys?.auth === "string" && subscription.keys.auth.length > 8;
  } catch (error) { return false; }
}

async function enviarPush(scope, ownerId, empresaId, payload) {
  if (!WEB_PUSH_PUBLIC_KEY) return;
  const suscripciones = find("push_suscripciones", s => s.scope === scope && s.owner_id === ownerId
    && (empresaId === null || s.empresa_id === empresaId));
  if (!suscripciones.length) return;
  const yaEnviada = findOne("notificaciones_enviadas", n => n.clave === payload.clave && n.scope === scope && n.owner_id === ownerId);
  if (yaEnviada) return;
  insert("notificaciones_enviadas", {
    clave: payload.clave, scope, owner_id: ownerId, empresa_id: empresaId,
    titulo: payload.titulo, creado_en: new Date().toISOString()
  });
  await Promise.allSettled(suscripciones.map(async registro => {
    try {
      await webpush.sendNotification({ endpoint: registro.endpoint, keys: registro.keys }, JSON.stringify({
        title: payload.titulo, body: payload.cuerpo, url: payload.url || "/", tag: payload.clave
      }), { TTL: 60 * 60 });
      update("push_suscripciones", registro.id, { ultimo_envio: new Date().toISOString() });
    } catch (error) {
      if (error.statusCode === 404 || error.statusCode === 410) remove("push_suscripciones", registro.id);
      else console.warn("Web Push no entregado:", error.statusCode || error.message);
    }
  }));
}

function notificarAdministradores(empresaId, payload) {
  const administradores = find("usuarios", u => u.empresa_id === empresaId && u.rol === "admin" && u.activo === 1);
  return Promise.allSettled(administradores.map(u => enviarPush("usuario", u.id, empresaId, payload)));
}

function notificarSuperadmins(payload) {
  const superadmins = find("superadmins", s => s.activo === 1);
  return Promise.allSettled(superadmins.map(s => enviarPush("superadmin", s.id, null, payload)));
}

function notificarAdministradoresCupo(empresaId) {
  const empresa = findOne("empresa", e => e.id === empresaId);
  if (!suscripcionActiva(empresa)) return Promise.resolve();
  const ocupados = contarCuposOcupados(empresaId);
  const limite = Number(empresa.cupos_limite || 0);
  if (!limite) return Promise.resolve();
  const umbral = Math.max(1, Math.ceil(limite * 0.8));
  if (ocupados < umbral) return Promise.resolve();
  const completo = ocupados >= limite;
  return notificarAdministradores(empresaId, {
    clave: `cupos-${empresaId}-${limite}-${completo ? "completo" : "cerca"}`,
    titulo: completo ? "Llegaste al límite de cupos" : "Tu equipo está cerca del límite",
    cuerpo: completo
      ? `Ya ocupas ${ocupados} de ${limite} cupos. Solicita una ampliación para agregar más personas.`
      : `Ya ocupas ${ocupados} de ${limite} cupos. Puedes solicitar una ampliación desde Código y notificaciones.`,
    url: "/admin.html"
  });
}

function configurarRutasPush(scope, requireMiddleware, prefijo) {
  app.get(`${prefijo}/config`, requireMiddleware, (req, res) => {
    res.json({ ok: true, disponible: Boolean(WEB_PUSH_PUBLIC_KEY), clave_publica: WEB_PUSH_PUBLIC_KEY });
  });
  app.post(`${prefijo}/suscripcion`, requireMiddleware, (req, res) => {
    const subscription = req.body?.subscription;
    if (!subscriptionValida(subscription)) return res.status(400).json({ ok: false, error: "La suscripción del navegador no es válida." });
    const ownerId = scope === "usuario" ? req.usuario.id : req.superadmin.id;
    const empresaId = scope === "usuario" ? req.usuario.empresa_id : null;
    const existente = findOne("push_suscripciones", s => s.endpoint === subscription.endpoint && s.scope === scope && s.owner_id === ownerId);
    if (existente) {
      update("push_suscripciones", existente.id, { scope, owner_id: ownerId, empresa_id: empresaId, keys: subscription.keys, actualizado_en: new Date().toISOString() });
    } else {
      insert("push_suscripciones", { scope, owner_id: ownerId, empresa_id: empresaId, endpoint: subscription.endpoint, keys: subscription.keys, creado_en: new Date().toISOString() });
    }
    res.json({ ok: true });
  });
  app.delete(`${prefijo}/suscripcion`, requireMiddleware, (req, res) => {
    const ownerId = scope === "usuario" ? req.usuario.id : req.superadmin.id;
    find("push_suscripciones", s => s.scope === scope && s.owner_id === ownerId).forEach(s => remove("push_suscripciones", s.id));
    res.json({ ok: true });
  });
}

configurarRutasPush("usuario", requireAuth, "/api/notificaciones");
configurarRutasPush("superadmin", requireSuperadmin, "/api/superadmin/notificaciones");

function horaLocalActual() {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: obtenerZonaHoraria(), hour: "2-digit", minute: "2-digit", hourCycle: "h23"
  }).format(new Date());
}

async function revisarRecordatoriosProgramados() {
  const hoy = claveFecha(new Date());
  const ahora = new Date();
  const minutosAhora = aMinutos(horaLocalActual());
  const fecha = normalizarFechaSoloDia(ahora);
  if (!fecha || minutosAhora === null) return;

  for (const usuario of db.usuarios.filter(u => u.activo === 1)) {
    const empresa = findOne("empresa", e => e.id === usuario.empresa_id);
    if (!suscripcionActiva(empresa)) continue;
    const registro = findOne("asistencia", a => a.empresa_id === usuario.empresa_id && a.correo === usuario.correo && a.fecha === hoy);
    if (!registro?.entrada || registro.salida) continue;
    if (registro.comida_entrada && !registro.comida_salida) {
      const salidaComida = aMinutos(registro.comida_entrada);
      const duracionComida = normalizarMinutosComida(usuario.minutos_comida || MINUTOS_COMIDA_POR_DEFECTO);
      if (salidaComida !== null && minutosAhora >= salidaComida + duracionComida) {
        await enviarPush("usuario", usuario.id, usuario.empresa_id, {
          clave: `regreso-comida-${usuario.empresa_id}-${usuario.id}-${hoy}`,
          titulo: "Registra tu regreso de comida",
          cuerpo: "Ya pasó tu tiempo de comida. No olvides marcar tu regreso.",
          url: "/dashboard.html"
        });
      }
    }
    const horario = obtenerHorarioParaFecha(usuario.area, fecha, usuario.correo, usuario.empresa_id);
    const finTurno = aMinutos(horario?.fin);
    if (finTurno !== null && minutosAhora >= finTurno - 15 && minutosAhora < finTurno) {
      await enviarPush("usuario", usuario.id, usuario.empresa_id, {
        clave: `salida-jornada-${usuario.empresa_id}-${usuario.id}-${hoy}`,
        titulo: "Ya casi termina tu turno",
        cuerpo: "No olvides marcar tu salida al terminar tu jornada.",
        url: "/dashboard.html"
      });
    }
  }

  const minutoDia = minutosAhora;
  if (minutoDia >= 9 * 60 && minutoDia < 9 * 60 + 2) {
    for (const empresa of db.empresa.filter(e => estadoSuscripcion(e) === "activa" && e.fecha_fin_suscripcion)) {
      const fin = fechaDesdeEntrada(empresa.fecha_fin_suscripcion);
      const diasRestantes = Math.ceil((new Date(fin.getFullYear(), fin.getMonth(), fin.getDate()).getTime() - new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate()).getTime()) / 86400000);
      if (![7, 3, 1, 0].includes(diasRestantes)) continue;
      await notificarAdministradores(empresa.id, {
        clave: `membresia-${empresa.id}-${hoy}-${diasRestantes}`,
        titulo: diasRestantes === 0 ? "Tu membresía vence hoy" : `Tu membresía vence en ${diasRestantes} días`,
        cuerpo: `Tu equipo tiene ${empresa.cupos_limite || 0} cupos. Contacta a JJM Clock para renovar.`,
        url: "/admin.html"
      });
    }
  }
}

let recordatoriosEnCurso = false;
function iniciarRecordatoriosProgramados() {
  setInterval(async () => {
    if (recordatoriosEnCurso) return;
    recordatoriosEnCurso = true;
    try { await revisarRecordatoriosProgramados(); }
    catch (error) { console.warn("No se pudieron revisar los recordatorios:", error.message); }
    finally { recordatoriosEnCurso = false; }
  }, 60 * 1000);
}

// =============================================
// ARRANQUE
// =============================================

const PORT = process.env.PORT || 3000;
async function iniciarServidor() {
  if (USAR_SUPABASE) await cargarDBSupabase();
  else cargarDB();

  inicializarDatos();
  inicializarAreasExistentes();
  inicializarSuscripcionesEquipos();
  asociarArchivoHorasOriginal();
  prepararCodigosConfiguracionAdmin();
  prepararTokenSetupSuperadmin();
  await esperarPersistenciaDB();

  iniciarRecordatoriosProgramados();
  const server = app.listen(PORT, () => {
    console.log(`🚀 ${APP_NAME} API escuchando en el puerto ${PORT}`);
    if (tokenSetupGenerado && tokenSetupSuperadmin) {
      console.log("🔐 Se generó una clave de configuración inicial para superadmin; almacénala fuera de los logs antes de configurarla en producción.");
    }
  });
  const cerrar = () => {
    esperarPersistenciaDB()
      .catch(error => console.error("No se guardaron los últimos cambios al apagar:", error.message))
      .finally(() => server.close());
  };
  process.once("SIGTERM", cerrar);
  process.once("SIGINT", cerrar);
}

iniciarServidor().catch(error => {
  console.error("No se pudo iniciar JJM Clock:", error.message);
  process.exitCode = 1;
});
