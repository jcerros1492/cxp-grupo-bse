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
  sincronizar();
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

const urlServicio = () => estado.sesion?.url || API_URL;

async function api(payload, url = urlServicio()) {
  if (!url) throw new Error('Falta la dirección del servicio de Google Sheets');
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 120000);
  try {
    // Sin encabezados personalizados: así el navegador no hace una consulta previa (CORS).
    const resp = await fetch(url, { method: 'POST', body: JSON.stringify(payload), signal: ctrl.signal, redirect: 'follow' });
    if (!resp.ok) throw new Error('El servicio respondió ' + resp.status);
    return await resp.json();
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
      if (r.auth === false) { await borrarTodo(); avisar(r.error, 'error'); return; }
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
    if (!x.ok) mensajes.push(['No se guardó un cambio: ' + x.error, 'error']);
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

export async function iniciarSesion(usuario, contrasena, url) {
  url = (url || API_URL || '').trim();
  const r = await api({ accion: 'login', usuario: usuario.trim(), contrasena, equipo: nombreEquipo() }, url);
  if (!r.ok) throw new Error(r.error || 'No se pudo iniciar sesión');
  if (estado.sesion?.url && estado.sesion.url !== url) await borrarTodo();
  if (estado.sesion?.usuario && estado.sesion.usuario !== r.usuario.usuario) estado.pendientes = [];
  estado.sesion = { token: r.token, url, ...r.usuario };
  estado.hoja = r.hoja || '';
  guardarLocal();
  notificar();
  await sincronizar();
}

export async function cambiarContrasena(actual, nueva) {
  const r = await api({ accion: 'cambiarContrasena', token: estado.sesion.token, actual, nueva });
  if (!r.ok) throw new Error(r.error || 'No se pudo cambiar la contraseña');
  Object.assign(estado.sesion, r.usuario, { debeCambiar: false });
  guardarLocal();
  notificar();
  sincronizar();
}

export async function cerrarSesion() {
  const token = estado.sesion?.token;
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
