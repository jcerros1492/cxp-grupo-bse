// Sincronización con Google Sheets y funcionamiento sin internet.
//
// - Lo último recibido de la hoja se guarda en el navegador (IndexedDB).
// - Cada cambio se aplica de inmediato en pantalla y queda en una cola de pendientes.
// - Cuando hay internet, la cola se envía al servicio de Google, que la valida y
//   devuelve lo que cambiaron los demás. Si no hay internet, se reintenta luego.

import { API_URL, SEGUNDOS_SINCRONIZAR } from './config.js';

const L = window.Logica;

// ---------------------------------------------------------------- IndexedDB mínimo

function abrirDB() {
  return new Promise((ok, mal) => {
    const req = indexedDB.open('cuentas-por-pagar', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('kv');
    req.onsuccess = () => ok(req.result);
    req.onerror = () => mal(req.error);
  });
}
async function idb(modo, fn) {
  const db = await abrirDB();
  return new Promise((ok, mal) => {
    const tx = db.transaction('kv', modo);
    const r = fn(tx.objectStore('kv'));
    tx.oncomplete = () => ok(r?.result);
    tx.onerror = () => mal(tx.error);
  });
}
const idbGet = (k) => idb('readonly', (s) => s.get(k));
const idbSet = (k, v) => idb('readwrite', (s) => s.put(v, k));
const idbDel = (k) => idb('readwrite', (s) => s.delete(k));

// ---------------------------------------------------------------- Estado

export const estado = {
  sesion: null,        // { token, usuario, nombre, rol, url, debeCambiar }
  servidor: L.vacio(), // copia de lo que hay en la hoja
  pendientes: [],      // cambios hechos aquí que aún no llegan a la hoja
  desde: 0,
  epoca: '',
  hoja: '',
  ultimoSync: null,
  red: navigator.onLine ? 'ok' : 'offline', // ok | sync | offline | error
  error: '',
  cargado: false,
};

let vista = null;
const oyentes = new Set();
const avisadores = new Set();
export const suscribir = (fn) => oyentes.add(fn);
export const alAvisar = (fn) => avisadores.add(fn);
const notificar = () => oyentes.forEach((fn) => fn());
const avisar = (msg, tipo) => avisadores.forEach((fn) => fn(msg, tipo));

/** Datos que se muestran: lo de la hoja más los cambios pendientes de enviar. */
export function datos() {
  if (!vista) recomputar();
  return vista;
}

function recomputar() {
  const d = JSON.parse(JSON.stringify(estado.servidor));
  L.prepararIndices(d);
  const ctx = { usuario: estado.sesion, ahora: Date.now() };
  for (const op of estado.pendientes) L.aplicarOp(d, op, ctx);
  vista = d;
}

let tGuardar;
function guardarLocal() {
  clearTimeout(tGuardar);
  tGuardar = setTimeout(() => {
    const { sesion, servidor, pendientes, desde, epoca, hoja, ultimoSync } = estado;
    idbSet('estado', { sesion, servidor, pendientes, desde, epoca, hoja, ultimoSync }).catch((e) => console.error(e));
  }, 200);
}

export async function iniciar() {
  try {
    const g = await idbGet('estado');
    if (g) Object.assign(estado, g);
  } catch (e) {
    console.error('No se pudo leer la copia local', e);
  }
  estado.cargado = true;
  recomputar();
  notificar();
  window.addEventListener('online', () => { estado.red = 'ok'; notificar(); sincronizar(); });
  window.addEventListener('offline', () => { estado.red = 'offline'; notificar(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') sincronizar(); });
  setInterval(() => { if (document.visibilityState === 'visible') sincronizar(); }, SEGUNDOS_SINCRONIZAR * 1000);
  if (estado.sesion) sincronizar(); else refrescarInicio();
}

// ---------------------------------------------------------------- Cambios

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
export const nuevoId = uid;

/**
 * Aplica un cambio aquí mismo y lo deja en cola para enviarlo a la hoja.
 * Devuelve { error } si no es válido, o { avisos } si se aplicó.
 */
export function ejecutar(op) {
  op.opId = uid();
  const r = L.aplicarOp(datos(), op, { usuario: estado.sesion, ahora: Date.now() });
  if (!r.ok) return { error: r.error };
  estado.pendientes.push(op);
  guardarLocal();
  notificar();
  programarSync(400);
  return { avisos: r.avisos || [] };
}

let tSync;
function programarSync(ms) {
  clearTimeout(tSync);
  tSync = setTimeout(sincronizar, ms);
}

// ---------------------------------------------------------------- Servicio de Google

// Conexión de esta computadora: dirección del servicio y su contraseña.
// Se guarda solo en este navegador; no está en el código publicado en GitHub.
const CLAVE_CONEXION = 'cxp-conexion';

export function conexion() {
  try {
    const c = JSON.parse(localStorage.getItem(CLAVE_CONEXION) || 'null');
    if (c?.url && c?.clave) return c;
  } catch { /* sin almacenamiento */ }
  return null;
}

/** Prueba la dirección y la contraseña del servicio; si funcionan, las guarda en esta computadora. */
export async function configurarConexion(url, clave) {
  const c = { url: String(url || API_URL).trim(), clave: String(clave || '').trim() };
  if (!/^(https:\/\/script\.google(usercontent)?\.com\/|http:\/\/localhost[:/])/.test(c.url)) throw new Error('La dirección debe ser la del servicio de Google (https://script.google.com/macros/s/…/exec)');
  if (!c.clave) throw new Error('Falta la contraseña del servicio');
  const r = await api({ accion: 'ping' }, c);
  if (!r.ok) throw new Error(r.error || 'No se pudo conectar');
  const anterior = conexion();
  c.modoPruebas = !!r.modoPruebas;
  // Los scripts más nuevos responden con estos datos; el script 1.4 solo con { ok }.
  c.pinNativo = 'modoPruebas' in r;
  c.gestionUsuarios = Array.isArray(r.usuarios);
  localStorage.setItem(CLAVE_CONEXION, JSON.stringify(c));
  guardarUsuarios(r.usuarios);
  if (anterior && anterior.url !== c.url) await borrarTodo();
  notificar();
}

// ---------------------------------------------------------------- Usuarios conocidos

const CLAVE_USUARIOS = 'cxp-usuarios';

function guardarUsuarios(lista) {
  if (!Array.isArray(lista)) return;
  try { localStorage.setItem(CLAVE_USUARIOS, JSON.stringify(lista)); } catch { /* sin almacenamiento */ }
}

/** Personas registradas en la hoja (la última lista recibida en este dispositivo). */
export function usuariosConocidos() {
  try { return JSON.parse(localStorage.getItem(CLAVE_USUARIOS) || 'null'); } catch { return null; }
}

export const modoPruebasServidor = () => !!conexion()?.modoPruebas;
export const gestionUsuarios = () => !!conexion()?.gestionUsuarios;

/**
 * El script 1.4 pide contraseñas de al menos 8 caracteres. Para usar un PIN de 4 números
 * sin cambiar el script, el PIN se envía con un formato fijo más largo.
 */
function claveDesdePin(pin) {
  const x = String(pin || '');
  if (conexion()?.pinNativo || !/^\d{4}$/.test(x)) return x;
  return 'pin:' + x + ':cxp-bse';
}

// ---------------------------------------------------------------- Entrada directa (sin PIN)
//
// Cada persona entra tocando su nombre. Con el script 1.4, que exige contraseña, la
// aplicación usa una clave interna por usuario: la primera vez entra con la clave
// temporal «acceso» (escrita en la hoja) y la reemplaza sola por la interna.

export const CLAVE_INICIAL = 'acceso';
const claveInterna = (usuario) => 'cxp:' + usuario + ':acceso-directo';

/** Entra como `usuario`: usa la sesión guardada en este dispositivo o inicia una nueva. */
export async function entrarDirecto(usuario) {
  if (entrarGuardado(usuario)) return;
  if (!navigator.onLine) throw new Error('La primera vez que una persona entra en este dispositivo se necesita internet');
  const intentos = [claveInterna(usuario), CLAVE_INICIAL];
  let r;
  for (const c of intentos) {
    r = await api({ accion: 'login', usuario, contrasena: c, equipo: nombreEquipo() });
    if (r.ok || /desactivado/i.test(r.error || '')) break;
  }
  if (!r.ok) {
    throw new Error(/desactivado/i.test(r.error || '') ? r.error
      : `No se pudo entrar como ${usuario}. En la hoja de Google, pestaña Usuarios, escriba «${CLAVE_INICIAL}» en la columna contrasenaTemporal de esa persona.`);
  }
  if (estado.pendientes.length && estado.sesion?.usuario && estado.sesion.usuario !== r.usuario.usuario) {
    throw new Error('Hay cambios de otra persona que aún no llegan a Google. Espere a que se sincronicen.');
  }
  estado.sesion = { token: r.token, ...r.usuario };
  estado.hoja = r.hoja || estado.hoja;
  if (r.usuario.debeCambiar) {
    const c = await api({ accion: 'cambiarContrasena', token: r.token, actual: CLAVE_INICIAL, nueva: claveInterna(usuario) });
    if (!c.ok) throw new Error(c.error || 'No se pudo preparar el acceso');
    estado.sesion.debeCambiar = false;
  }
  guardarSesionUsuario(estado.sesion);
  guardarLocal();
  notificar();
  await sincronizar();
}

// ---------------------------------------------------------------- Sesiones guardadas en este dispositivo

const CLAVE_SESIONES = 'cxp-sesiones';
function sesionesGuardadas() {
  try { return JSON.parse(localStorage.getItem(CLAVE_SESIONES) || '{}'); } catch { return {}; }
}
function guardarSesionUsuario(s) {
  if (!s?.usuario) return;
  const g = sesionesGuardadas();
  g[s.usuario] = s;
  try { localStorage.setItem(CLAVE_SESIONES, JSON.stringify(g)); } catch { /* sin almacenamiento */ }
}
function olvidarSesionUsuario(usuario) {
  const g = sesionesGuardadas();
  delete g[usuario];
  try { localStorage.setItem(CLAVE_SESIONES, JSON.stringify(g)); } catch { /* sin almacenamiento */ }
}

/** Entrar tocando el nombre si esa persona ya entró antes en este dispositivo. */
export function entrarGuardado(usuario) {
  const s = sesionesGuardadas()[usuario];
  if (!s || s.debeCambiar) return false;
  if (estado.pendientes.length && estado.sesion?.usuario !== usuario) {
    throw new Error('Hay cambios de ' + (estado.sesion?.nombre || 'otra persona') + ' que aún no llegan a Google. Espere a que se sincronicen.');
  }
  estado.sesion = { ...s };
  guardarLocal();
  notificar();
  sincronizar();
  return true;
}

/** Vuelve a la pantalla de usuarios sin borrar los datos de este dispositivo. */
export function cambiarUsuario() {
  if (estado.pendientes.length) throw new Error('Hay cambios que aún no llegan a Google. Espere a que se sincronicen antes de cambiar de usuario.');
  estado.sesion = null;
  guardarLocal();
  notificar();
}

function actualizarModo(valor) {
  const c = conexion();
  if (!c || valor === undefined || !!valor === !!c.modoPruebas) return;
  c.modoPruebas = !!valor;
  localStorage.setItem(CLAVE_CONEXION, JSON.stringify(c));
}

/** Refresca la lista de personas y el modo de pruebas (para la pantalla de inicio). */
export async function refrescarInicio() {
  if (!conexion() || !navigator.onLine) return;
  try {
    const r = await api({ accion: 'ping' });
    if (r.ok) {
      const c = conexion();
      if (c) {
        Object.assign(c, { modoPruebas: !!r.modoPruebas, pinNativo: 'modoPruebas' in r, gestionUsuarios: Array.isArray(r.usuarios) });
        localStorage.setItem(CLAVE_CONEXION, JSON.stringify(c));
      }
      guardarUsuarios(r.usuarios);
      notificar();
    }
  } catch { /* sin conexión: se usa lo guardado */ }
}

/** Sube un comprobante (PDF o imagen) a la carpeta de Drive del servicio. Necesita internet. */
export async function subirArchivo({ nombre, tipo, base64 }) {
  if (!navigator.onLine) throw new Error('Para adjuntar comprobantes se necesita internet');
  const r = await api({ accion: 'subirArchivo', token: estado.sesion?.token, nombre, tipo, base64 });
  if (!r.ok) throw new Error(/desconocida/i.test(r.error || '') ? 'Falta agregar el bloque de adjuntos en Apps Script (versión 1.18)' : (r.error || 'No se pudo subir el archivo'));
  return r;
}

/** Trae un comprobante para verlo (no se guarda en el dispositivo). */
export async function leerArchivo(id) {
  if (!navigator.onLine) throw new Error('Para ver comprobantes se necesita internet');
  const r = await api({ accion: 'leerArchivo', token: estado.sesion?.token, id });
  if (!r.ok) throw new Error(r.error || 'No se pudo abrir el comprobante');
  return r;
}

/** Administración de usuarios (solo administradores). */
export async function adminUsuarios(datos) {
  const r = await api({ accion: 'adminUsuarios', token: estado.sesion.token, ...datos });
  if (!r.ok) throw new Error(r.error || 'No se pudo completar');
  guardarUsuarios(r.usuarios);
  notificar();
  return r;
}

/** Confirma que quien quiere cambiar la conexión conoce la contraseña del servicio. */
export const verificarClaveServicio = (clave) => conexion()?.clave === String(clave || '').trim();

async function api(payload, c = conexion()) {
  if (!c) throw new Error('Esta computadora no tiene configurada la conexión con Google Sheets');
  payload = { ...payload, servicio: c.clave };
  const url = c.url;
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 120000);
  try {
    // Sin encabezados personalizados: así el navegador no hace una consulta previa (CORS).
    const resp = await fetch(url, { method: 'POST', body: JSON.stringify(payload), signal: ctrl.signal, redirect: 'follow' });
    if (!resp.ok) throw new Error('El servicio respondió ' + resp.status);
    const r = await resp.json();
    // Si cambiaron la contraseña del servicio, esta computadora debe configurarse de nuevo.
    const guardada = conexion();
    if (r.servicio === false && guardada && guardada.clave === c.clave) {
      localStorage.removeItem(CLAVE_CONEXION);
      notificar();
    }
    return r;
  } finally {
    clearTimeout(t);
  }
}

let sincronizando = false;

export async function sincronizar() {
  if (sincronizando || !estado.sesion || estado.sesion.debeCambiar) return;
  if (!navigator.onLine) { estado.red = 'offline'; notificar(); return; }
  sincronizando = true;
  estado.red = 'sync';
  notificar();
  const enviadas = estado.pendientes.slice(0, 100);
  try {
    const r = await api({ accion: 'sync', token: estado.sesion.token, desde: estado.desde, epoca: estado.epoca, ops: enviadas });
    if (!r.ok) {
      if (r.auth === false) { olvidarSesionUsuario(estado.sesion?.usuario); await borrarTodo(); avisar(r.error, 'error'); return; }
      if (r.debeCambiar) { estado.sesion.debeCambiar = true; guardarLocal(); return; }
      throw new Error(r.error || 'Error del servicio');
    }
    recibir(r);
    estado.red = 'ok';
    estado.error = '';
    if (estado.pendientes.length) programarSync(100);
  } catch (e) {
    estado.red = navigator.onLine ? 'error' : 'offline';
    estado.error = e.name === 'AbortError' ? 'El servicio tardó demasiado en responder' : e.message;
  } finally {
    sincronizando = false;
    notificar();
  }
}

function recibir(r) {
  const procesadas = new Set((r.resultados || []).map((x) => x.opId));
  const mensajes = [];
  for (const x of r.resultados || []) {
    if (!x.ok) {
      const pista = /tipo de movimiento inv/i.test(x.error || '') ? ' · Para guardar notas de crédito falta agregar una línea en Apps Script (versión 1.14).' : '';
      mensajes.push(['No se guardó un cambio: ' + x.error + pista, 'error']);
    }
    else if (x.avisos?.length) mensajes.push([x.avisos.join(' · '), 'aviso']);
  }
  estado.pendientes = estado.pendientes.filter((op) => !procesadas.has(op.opId));
  if (r.completo) estado.servidor = L.vacio();
  for (const t of L.TABLAS) {
    const clave = L.CLAVES[t];
    for (const e of r.delta?.[t] || []) estado.servidor[t][e[clave]] = e;
  }
  estado.desde = r.ahora;
  estado.epoca = r.epoca || '';
  estado.ultimoSync = Date.now();
  if (r.usuario) Object.assign(estado.sesion, r.usuario);
  if (r.usuarios) guardarUsuarios(r.usuarios);
  actualizarModo(r.modoPruebas);
  recomputar();
  guardarLocal();
  mensajes.forEach(([m, t]) => avisar(m, t));
}

// ---------------------------------------------------------------- Sesión

function nombreEquipo() {
  const ua = navigator.userAgent;
  const so = /Windows/.test(ua) ? 'Windows' : /Mac/.test(ua) ? 'Mac' : /Android/.test(ua) ? 'Android' : /iPhone|iPad/.test(ua) ? 'iPhone/iPad' : 'Otro';
  const nav = /Edg\//.test(ua) ? 'Edge' : /Chrome\//.test(ua) ? 'Chrome' : /Firefox\//.test(ua) ? 'Firefox' : /Safari\//.test(ua) ? 'Safari' : 'Navegador';
  return nav + ' en ' + so;
}

export async function iniciarSesion(usuario, contrasena) {
  // Se prueba el PIN (ya convertido) y, si no coincide, tal cual: así funciona también el PIN
  // temporal escrito en la hoja, aunque Google lo haya guardado sin el cero inicial.
  const raw = String(contrasena || '');
  const candidatos = [...new Set([claveDesdePin(raw), raw, /^\d+$/.test(raw) ? String(Number(raw)) : raw])];
  let r;
  for (const c of candidatos) {
    r = await api({ accion: 'login', usuario: usuario.trim(), contrasena: c, equipo: nombreEquipo() });
    if (r.ok || /desactivado|intentos/i.test(r.error || '')) break;
  }
  if (!r.ok) throw new Error(/incorrect/i.test(r.error || '') ? 'PIN incorrecto' : (r.error || 'No se pudo iniciar sesión'));
  if (estado.sesion?.usuario && estado.sesion.usuario !== r.usuario.usuario) estado.pendientes = [];
  estado.sesion = { token: r.token, ...r.usuario };
  estado.hoja = r.hoja || '';
  guardarSesionUsuario(estado.sesion);
  guardarLocal();
  notificar();
  await sincronizar();
}

export async function cambiarContrasena(actual, nueva) {
  const r = await api({ accion: 'cambiarContrasena', token: estado.sesion.token, actual: claveDesdePin(actual), nueva: claveDesdePin(nueva) });
  if (!r.ok) throw new Error(/contraseña actual/i.test(r.error || '') ? 'El PIN actual no es correcto' : (r.error || 'No se pudo cambiar el PIN'));
  Object.assign(estado.sesion, r.usuario, { debeCambiar: false });
  guardarSesionUsuario(estado.sesion);
  guardarLocal();
  notificar();
  sincronizar();
}

export async function cerrarSesion() {
  const token = estado.sesion?.token;
  olvidarSesionUsuario(estado.sesion?.usuario);
  if (token && navigator.onLine) api({ accion: 'salir', token }).catch(() => {});
  await borrarTodo();
}

export async function borrarTodo() {
  Object.assign(estado, { sesion: null, servidor: L.vacio(), pendientes: [], desde: 0, epoca: '', hoja: '', ultimoSync: null });
  recomputar();
  await idbDel('estado').catch(() => {});
  notificar();
}

/** Envía el Excel ya convertido a la hoja (reemplaza facturas y movimientos). */
export async function importar(paquete) {
  if (!navigator.onLine) throw new Error('Se necesita internet para cargar el Excel');
  const r = await api({ accion: 'importar', token: estado.sesion.token, desde: 0, epoca: '', ops: [], datos: paquete });
  if (!r.ok) throw new Error(r.error || 'No se pudo cargar');
  estado.pendientes = [];
  recibir(r);
  notificar();
}
