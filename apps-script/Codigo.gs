/**
 * Cuentas por Pagar — servicio en Google Sheets.
 *
 * Este script vive dentro de la hoja de cálculo (Extensiones → Apps Script) junto con
 * Logica.gs, que contiene las mismas reglas que usa la aplicación web.
 *
 * La aplicación web le envía los cambios hechos (aunque se hayan hecho sin internet);
 * el script los valida con un bloqueo para que dos personas no choquen, los guarda en
 * la hoja y devuelve lo que cambió desde la última sincronización.
 *
 * Primera vez: ejecute la función configurarInicio() desde el editor.
 */

var VERSION_SERVICIO = 1;

// Pestañas de la hoja. Tipos: t = texto, n = número, b = sí/no.
var ESQUEMA = {
  facturas: { hoja: 'Facturas', cols: [
    ['id', 't'], ['sociedad', 't'], ['codigo', 't'], ['proveedor', 't'], ['documento', 't'], ['fechaFactura', 't'],
    ['fechaVence', 't'], ['venceEstimado', 'b'], ['moneda', 't'], ['monto', 'n'], ['pagado', 'n'], ['saldo', 'n'],
    ['estado', 't'], ['revision', 'b'], ['observaciones', 't'], ['creadaPor', 't'], ['v', 'n']] },
  movimientos: { hoja: 'Movimientos', cols: [
    ['id', 't'], ['facturaId', 't'], ['tipo', 't'], ['monto', 'n'], ['fecha', 't'], ['comprobante', 't'], ['por', 't'],
    ['nota', 't'], ['programadoPor', 't'], ['loteId', 't'], ['v', 'n']] },
  programaciones: { hoja: 'Programaciones', cols: [
    ['id', 't'], ['facturaId', 't'], ['monto', 'n'], ['fechaPago', 't'], ['nota', 't'], ['estado', 't'], ['por', 't'],
    ['creado', 't'], ['cerradoPor', 't'], ['cerrado', 't'], ['comprobante', 't'], ['motivo', 't'], ['v', 'n']] },
  proveedores: { hoja: 'Proveedores', cols: [
    ['codigo', 't'], ['nombre', 't'], ['alias', 't'], ['ibanCRC', 't'], ['ibanUSD', 't'], ['plazo', 'n'], ['v', 'n']] },
  sociedades: { hoja: 'Sociedades', cols: [
    ['clave', 't'], ['corto', 't'], ['cuentaCRC', 't'], ['cuentaUSD', 't'], ['descripcionBanco', 't'], ['v', 'n']] },
  lotes: { hoja: 'Lotes', cols: [
    ['id', 't'], ['sociedad', 't'], ['moneda', 't'], ['fecha', 't'], ['comprobante', 't'], ['total', 'n'],
    ['aplicadoPor', 't'], ['aplicado', 't'], ['detalle', 't'], ['v', 'n']] },
  config: { hoja: 'Config', cols: [['clave', 't'], ['valor', 't'], ['v', 'n']] },
};
var HOJA_USUARIOS = 'Usuarios';
// contrasenaTemporal: la escribe el administrador; al entrar con ella se pide crear una nueva.
// hash y sal: la contraseña cifrada (nunca se guarda la contraseña real).
var COLS_USUARIOS = ['usuario', 'nombre', 'rol', 'activo', 'contrasenaTemporal', 'hash', 'sal'];
var HOJA_SESIONES = 'Sesiones';
var COLS_SESIONES = ['token', 'usuario', 'creada', 'ultimoUso', 'equipo'];
var HOJA_BITACORA = 'Bitacora';
var COLS_BITACORA = ['opId', 'fecha', 'usuario', 'tipo', 'resultado', 'detalle'];

// ---------------------------------------------------------------- Entrada web

function doGet() {
  return json_({ ok: true, servicio: 'cuentas-por-pagar', version: VERSION_SERVICIO });
}

function doPost(e) {
  var res;
  try {
    var p = JSON.parse(e.postData.contents);
    res = manejar_(p);
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return json_(res);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function manejar_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var libro = SpreadsheetApp.getActive();
    if (p.accion === 'login') return login_(libro, p);
    var usuario = autenticar_(libro, p.token);
    if (!usuario) return { ok: false, auth: false, error: 'La sesión expiró o el usuario fue desactivado. Inicie sesión de nuevo.' };
    if (p.accion === 'cambiarContrasena') return cambiarContrasena_(libro, usuario, p);
    if (p.accion === 'salir') { cerrarSesion_(libro, p.token); return { ok: true }; }
    if (usuario.debeCambiar) return { ok: false, debeCambiar: true, error: 'Debe crear su contraseña antes de continuar' };
    if (p.accion === 'quienSoy') return { ok: true, usuario: publico_(usuario), hoja: libro.getUrl() };

    var ahora = Date.now();
    var alm = cargarTodo_(libro);
    var d = alm.datos;

    if (p.accion === 'importar') {
      if (usuario.rol !== 'admin') return { ok: false, error: 'Solo el administrador puede cargar el Excel' };
      importar_(libro, alm, p.datos, ahora, usuario);
      alm = cargarTodo_(libro);
      d = alm.datos;
    } else if (p.accion !== 'sync') {
      return { ok: false, error: 'Acción desconocida' };
    }

    // Aplicar los cambios recibidos, en orden, una sola vez cada uno.
    var resultados = [], cambiados = {}, bitacora = [];
    var hechas = opsHechas_(libro);
    (p.ops || []).forEach(function (op) {
      if (!op || !op.opId) return;
      if (hechas[op.opId]) { resultados.push({ opId: op.opId, ok: true, repetida: true }); return; }
      var r = Logica.aplicarOp(d, op, { usuario: usuario, ahora: ahora });
      if (r.ok) r.cambios.forEach(function (c) { (cambiados[c[0]] = cambiados[c[0]] || {})[c[1]] = true; });
      resultados.push({ opId: op.opId, ok: r.ok, error: r.error || '', avisos: r.avisos || [] });
      bitacora.push([op.opId, new Date(ahora), usuario.nombre, op.tipo, r.ok ? 'ok' : 'rechazada',
        r.ok ? (r.avisos || []).join(' · ') : r.error]);
      hechas[op.opId] = true;
    });
    Object.keys(cambiados).forEach(function (tabla) { guardarCambios_(libro, alm, tabla, Object.keys(cambiados[tabla])); });
    if (bitacora.length) agregarFilas_(hoja_(libro, HOJA_BITACORA, COLS_BITACORA), bitacora);

    var epoca = Logica.cfg(d, 'epoca');
    var completo = !p.desde || String(p.epoca || '') !== epoca;
    return {
      ok: true, ahora: ahora, epoca: epoca, completo: completo, usuario: publico_(usuario),
      resultados: resultados, delta: exportar_(d, completo ? -1 : Number(p.desde)),
    };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- Usuarios y sesiones

function leerUsuarios_(libro) {
  var sh = hoja_(libro, HOJA_USUARIOS, COLS_USUARIOS);
  var n = sh.getLastRow() - 1;
  var filas = n > 0 ? sh.getRange(2, 1, n, COLS_USUARIOS.length).getValues() : [];
  return { sh: sh, lista: filas.map(function (f, i) {
    var a = String(f[3]).toUpperCase();
    return {
      fila: i + 2, usuario: String(f[0]).trim().toLowerCase(), nombre: String(f[1]), rol: String(f[2] || 'aplica').toLowerCase(),
      activo: f[3] === true || a === 'TRUE' || a === 'SI' || a === 'SÍ', temporal: String(f[4] || ''), hash: String(f[5] || ''), sal: String(f[6] || ''),
    };
  }) };
}

function publico_(u) { return { usuario: u.usuario, nombre: u.nombre, rol: u.rol, debeCambiar: !!u.debeCambiar }; }

function cifrar_(contrasena, sal) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + contrasena, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) hex += ('0' + (bytes[i] & 0xff).toString(16)).slice(-2);
  // Varias vueltas para que adivinar contraseñas sea lento.
  for (var k = 0; k < 1000; k++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hex + sal, Utilities.Charset.UTF_8);
    hex = '';
    for (var j = 0; j < bytes.length; j++) hex += ('0' + (bytes[j] & 0xff).toString(16)).slice(-2);
  }
  return hex;
}

function login_(libro, p) {
  var nombre = String(p.usuario || '').trim().toLowerCase();
  var clave = String(p.contrasena || '');
  var us = leerUsuarios_(libro).lista;
  var u = us.filter(function (x) { return x.usuario === nombre; })[0];
  var error = { ok: false, error: 'Usuario o contraseña incorrectos' };
  if (!u || !nombre || !clave) return error;
  if (!u.activo) return { ok: false, error: 'Este usuario está desactivado. Hable con el administrador.' };
  var debeCambiar = false;
  if (u.temporal && clave === u.temporal) debeCambiar = true;
  else if (!(u.hash && cifrar_(clave, u.sal) === u.hash)) return error;
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var ahora = new Date();
  agregarFilas_(hoja_(libro, HOJA_SESIONES, COLS_SESIONES), [[token, u.usuario, ahora, ahora, String(p.equipo || '').slice(0, 120)]]);
  u.debeCambiar = debeCambiar;
  agregarFilas_(hoja_(libro, HOJA_BITACORA, COLS_BITACORA), [['login-' + token.slice(0, 12), ahora, u.nombre, 'iniciar sesión', 'ok', String(p.equipo || '').slice(0, 120)]]);
  return { ok: true, token: token, usuario: publico_(u), hoja: libro.getUrl() };
}

/** Busca la sesión por su token y devuelve el usuario (o null). */
function autenticar_(libro, token) {
  if (!token) return null;
  var sh = hoja_(libro, HOJA_SESIONES, COLS_SESIONES);
  var n = sh.getLastRow() - 1;
  if (n < 1) return null;
  var filas = sh.getRange(2, 1, n, 2).getValues();
  var nombre = null, fila = 0;
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][0]) === String(token)) { nombre = String(filas[i][1]); fila = i + 2; break; }
  }
  if (!nombre) return null;
  var u = leerUsuarios_(libro).lista.filter(function (x) { return x.usuario === nombre; })[0];
  if (!u || !u.activo) return null;
  u.debeCambiar = !!u.temporal || !u.hash;
  sh.getRange(fila, 4).setValue(new Date());
  return u;
}

function cambiarContrasena_(libro, u, p) {
  var nueva = String(p.nueva || '');
  if (nueva.length < 8) return { ok: false, error: 'La contraseña debe tener al menos 8 caracteres' };
  if (nueva === u.temporal) return { ok: false, error: 'La nueva contraseña debe ser distinta de la temporal' };
  if (!u.debeCambiar && !(u.hash && cifrar_(String(p.actual || ''), u.sal) === u.hash)) {
    return { ok: false, error: 'La contraseña actual no es correcta' };
  }
  var sal = Utilities.getUuid().replace(/-/g, '');
  var us = leerUsuarios_(libro);
  us.sh.getRange(u.fila, 5, 1, 3).setNumberFormat('@').setValues([['', cifrar_(nueva, sal), sal]]);
  agregarFilas_(hoja_(libro, HOJA_BITACORA, COLS_BITACORA), [['clave-' + Date.now(), new Date(), u.nombre, 'cambiar contraseña', 'ok', '']]);
  u.debeCambiar = false;
  return { ok: true, usuario: publico_(u) };
}

function cerrarSesion_(libro, token) {
  var sh = hoja_(libro, HOJA_SESIONES, COLS_SESIONES);
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var filas = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][0]) === String(token)) { sh.getRange(i + 2, 1, 1, COLS_SESIONES.length).clearContent(); return; }
  }
}

// ---------------------------------------------------------------- Lectura y escritura

function hoja_(libro, nombre, cols) {
  var sh = libro.getSheetByName(nombre);
  if (!sh) {
    sh = libro.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function nombresCols_(tabla) { return ESQUEMA[tabla].cols.map(function (c) { return c[0]; }); }

/** Lee todas las pestañas y devuelve los datos y la fila de cada registro. */
function cargarTodo_(libro) {
  var datos = Logica.vacio(), filas = {};
  Logica.TABLAS.forEach(function (tabla) {
    var esq = ESQUEMA[tabla], clave = Logica.CLAVES[tabla];
    var sh = hoja_(libro, esq.hoja, nombresCols_(tabla));
    filas[tabla] = {};
    var n = sh.getLastRow() - 1;
    if (n < 1) return;
    var valores = sh.getRange(2, 1, n, esq.cols.length).getValues();
    for (var i = 0; i < valores.length; i++) {
      var obj = {};
      esq.cols.forEach(function (c, j) { obj[c[0]] = convertir_(valores[i][j], c[1]); });
      var k = obj[clave];
      if (k === '') continue;
      datos[tabla][k] = obj;
      filas[tabla][k] = i + 2;
    }
  });
  Logica.prepararIndices(datos);
  return { datos: datos, filas: filas };
}

function convertir_(v, tipo) {
  if (tipo === 'n') return v === '' || v === null ? 0 : Number(v);
  if (tipo === 'b') return v === true || String(v).toUpperCase() === 'TRUE';
  // Si Google convirtió un texto en fecha, se vuelve a texto con la zona horaria de la hoja.
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'America/Costa_Rica', 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function fila_(tabla, obj) {
  return ESQUEMA[tabla].cols.map(function (c) {
    var v = obj[c[0]];
    if (c[1] === 'n') return Number(v) || 0;
    if (c[1] === 'b') return !!v;
    return v === null || v === undefined ? '' : String(v);
  });
}

function formatear_(sh, tabla, filas) {
  if (filas < 1) return;
  ESQUEMA[tabla].cols.forEach(function (c, j) {
    if (c[1] === 't') sh.getRange(2, j + 1, filas, 1).setNumberFormat('@');
  });
}

function guardarCambios_(libro, alm, tabla, claves) {
  var sh = hoja_(libro, ESQUEMA[tabla].hoja, nombresCols_(tabla));
  var nuevos = [];
  claves.forEach(function (k) {
    var obj = alm.datos[tabla][k];
    if (!obj) return;
    var valores = fila_(tabla, obj);
    var r = alm.filas[tabla][k];
    if (r) sh.getRange(r, 1, 1, valores.length).setNumberFormats([formatos_(tabla)]).setValues([valores]);
    else nuevos.push(valores);
  });
  if (nuevos.length) {
    var inicio = sh.getLastRow() + 1;
    sh.getRange(inicio, 1, nuevos.length, nuevos[0].length).setNumberFormats(nuevos.map(function () { return formatos_(tabla); })).setValues(nuevos);
    nuevos.forEach(function (v, i) { alm.filas[tabla][v[0]] = inicio + i; });
  }
}

function formatos_(tabla) {
  return ESQUEMA[tabla].cols.map(function (c) { return c[1] === 't' ? '@' : c[1] === 'n' ? (c[0] === 'v' ? '0' : '#,##0.00') : 'General'; });
}

function agregarFilas_(sh, filas) {
  if (!filas.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
}

function opsHechas_(libro) {
  var sh = hoja_(libro, HOJA_BITACORA, COLS_BITACORA);
  var hechas = {};
  var n = sh.getLastRow() - 1;
  if (n < 1) return hechas;
  sh.getRange(2, 1, n, 1).getValues().forEach(function (f) { hechas[String(f[0])] = true; });
  return hechas;
}

/** Registros con versión >= desde (o todos si desde < 0). */
function exportar_(d, desde) {
  var out = {};
  Logica.TABLAS.forEach(function (tabla) {
    out[tabla] = [];
    Object.keys(d[tabla]).forEach(function (k) {
      var e = d[tabla][k];
      if (desde < 0 || Number(e.v) >= desde) out[tabla].push(e);
    });
  });
  return out;
}

// ---------------------------------------------------------------- Carga del Excel

/**
 * Reemplaza facturas, movimientos y programaciones con lo que viene del Excel.
 * Conserva IBAN, alias y cuentas ya registrados en Proveedores y Sociedades.
 */
function importar_(libro, alm, x, ahora, usuario) {
  var d = Logica.vacio();
  (x.facturas || []).forEach(function (f) { f.v = ahora; d.facturas[f.id] = f; });
  (x.movimientos || []).forEach(function (m) { m.v = ahora; d.movimientos[m.id] = m; });
  (x.proveedores || []).forEach(function (p) {
    var ant = alm.datos.proveedores[p.codigo];
    if (ant) { p.alias = ant.alias; p.ibanCRC = ant.ibanCRC; p.ibanUSD = ant.ibanUSD; p.plazo = ant.plazo || p.plazo; }
    p.v = ahora;
    d.proveedores[p.codigo] = p;
  });
  Object.keys(alm.datos.proveedores).forEach(function (k) { if (!d.proveedores[k]) d.proveedores[k] = alm.datos.proveedores[k]; });
  (x.sociedades || []).forEach(function (s) {
    var ant = alm.datos.sociedades[s.clave];
    if (ant) { s.corto = ant.corto || s.corto; s.cuentaCRC = ant.cuentaCRC; s.cuentaUSD = ant.cuentaUSD; s.descripcionBanco = ant.descripcionBanco; }
    s.v = ahora;
    d.sociedades[s.clave] = s;
  });
  d.config = alm.datos.config;
  d.config.epoca = { clave: 'epoca', valor: String(ahora), v: ahora };
  d.config.importado = { clave: 'importado', valor: new Date(ahora).toISOString() + ' · ' + usuario.nombre, v: ahora };
  Object.keys(d.proveedores).forEach(function (k) { d.proveedores[k].v = ahora; });
  Logica.recalcularTodo(d);
  Object.keys(d.facturas).forEach(function (id) {
    var f = d.facturas[id];
    var p = d.proveedores[f.codigo];
    if (p) f.proveedor = p.nombre;
  });

  ['facturas', 'movimientos', 'programaciones', 'proveedores', 'sociedades', 'lotes', 'config'].forEach(function (tabla) {
    var esq = ESQUEMA[tabla];
    var sh = hoja_(libro, esq.hoja, nombresCols_(tabla));
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    var filas = Object.keys(d[tabla]).map(function (k) { return fila_(tabla, d[tabla][k]); });
    if (filas.length) {
      formatear_(sh, tabla, filas.length);
      sh.getRange(2, 1, filas.length, esq.cols.length).setValues(filas);
    }
  });
  agregarFilas_(hoja_(libro, HOJA_BITACORA, COLS_BITACORA), [['importar-' + ahora, new Date(ahora), usuario.nombre, 'importar', 'ok',
    (x.facturas || []).length + ' facturas, ' + (x.movimientos || []).length + ' movimientos']]);
}

// ---------------------------------------------------------------- Instalación

/**
 * Ejecute esta función UNA vez desde el editor de Apps Script.
 * Crea las pestañas y la pestaña Usuarios con un usuario y una contraseña temporal para cada persona.
 *
 * Para agregar a alguien: escriba una fila nueva en Usuarios (usuario, nombre, rol, TRUE y una
 * contraseña temporal). Roles: admin, aplica, aprueba.
 * Si alguien olvida su contraseña: escriba una contraseña temporal nueva en su fila.
 * Para quitarle el acceso: ponga FALSE en la columna activo.
 */
function configurarInicio() {
  var libro = SpreadsheetApp.getActive();
  Logica.TABLAS.forEach(function (tabla) { hoja_(libro, ESQUEMA[tabla].hoja, nombresCols_(tabla)); });
  hoja_(libro, HOJA_BITACORA, COLS_BITACORA);
  hoja_(libro, HOJA_SESIONES, COLS_SESIONES);
  var us = hoja_(libro, HOJA_USUARIOS, COLS_USUARIOS);
  if (us.getLastRow() < 2) {
    var temporal = function () { return 'Cxp-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6); };
    us.getRange(2, 1, 3, COLS_USUARIOS.length).setNumberFormat('@').setValues([
      ['jonathan', 'Jonathan Cerros', 'admin', 'TRUE', temporal(), '', ''],
      ['norman', 'Norman Alpízar', 'aprueba', 'TRUE', temporal(), '', ''],
      ['daniel', 'Daniel Barquero', 'aplica', 'TRUE', temporal(), '', ''],
    ]);
  }
  var hoja1 = libro.getSheetByName('Hoja 1') || libro.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && libro.getSheets().length > 1) libro.deleteSheet(hoja1);
  SpreadsheetApp.flush();
  Logger.log('Listo. En la pestaña Usuarios está el usuario y la contraseña temporal de cada persona.');
}
