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

// ================================================================
//  CONTRASEÑA DEL SERVICIO
//  Escriba aquí, entre las comillas, la contraseña que se pedirá en cada
//  computadora junto con el enlace /exec. Es la misma para todas.
//  Después de cambiarla: guarde (Ctrl + S) e implemente una nueva versión.
// ================================================================
var CLAVE_SERVICIO = 'ESCRIBA-AQUI-SU-CONTRASEÑA';

// ================================================================
//  MODO DE PRUEBAS
//  true  = se entra tocando el nombre, sin PIN (solo para probar).
//  false = cada persona entra con su PIN de 4 dígitos.
//  La contraseña del servicio se sigue pidiendo en ambos casos.
//  Después de cambiarlo: guarde (Ctrl + S) e implemente una nueva versión.
// ================================================================
var MODO_PRUEBAS = true;

// Intentos fallidos de PIN antes de bloquear a la persona por 15 minutos.
var INTENTOS_PIN = 5;


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
    // Sin la contraseña del servicio no se atiende ninguna consulta, aunque se conozca el enlace.
    if (!claveServicioCorrecta_(p.servicio)) return { ok: false, servicio: false, error: 'La contraseña del servicio no es correcta' };
    if (p.accion === 'ping') return { ok: true, modoPruebas: !!MODO_PRUEBAS, usuarios: listaPublica_(libro) };
    if (p.accion === 'login') return login_(libro, p);
    var usuario = autenticar_(libro, p.token);
    if (!usuario) return { ok: false, auth: false, error: 'La sesión expiró o el usuario fue desactivado. Inicie sesión de nuevo.' };
    if (p.accion === 'cambiarContrasena') return cambiarContrasena_(libro, usuario, p);
    if (p.accion === 'salir') { cerrarSesion_(libro, p.token); return { ok: true }; }
    if (usuario.debeCambiar) return { ok: false, debeCambiar: true, error: 'Debe crear su contraseña antes de continuar' };
    if (p.accion === 'quienSoy') return { ok: true, usuario: publico_(usuario), hoja: libro.getUrl() };
    if (p.accion === 'adminUsuarios') return adminUsuarios_(libro, usuario, p);

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
      modoPruebas: !!MODO_PRUEBAS, usuarios: listaPublica_(libro, usuario.rol === 'admin'),
      resultados: resultados, delta: exportar_(d, completo ? -1 : Number(p.desde)),
    };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- Contraseña del servicio

/** La contraseña del servicio es la que está escrita arriba, en CLAVE_SERVICIO. */
function claveServicio_() {
  var c = String(CLAVE_SERVICIO || '').trim();
  if (!c || c === 'ESCRIBA-AQUI-SU-CONTRASEÑA') return '';
  return c;
}

function claveServicioCorrecta_(clave) {
  var real = claveServicio_();
  return !!real && String(clave || '') === real;
}

/** Muestra en el registro de ejecución la contraseña del servicio. */
function mostrarClaveServicio() {
  var c = claveServicio_();
  Logger.log(c ? 'Contraseña del servicio: ' + c : 'Falta escribir la contraseña en CLAVE_SERVICIO, al inicio del código.');
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
      activo: f[3] === true || a === 'TRUE' || a === 'SI' || a === 'SÍ', temporal: normalizarTemporal_(f[4]), hash: String(f[5] || ''), sal: String(f[6] || ''),
    };
  }) };
}

function normalizarTemporal_(v) {
  var t = String(v === null || v === undefined ? '' : v).trim();
  return /^\d{1,4}$/.test(t) ? ('000' + t).slice(-4) : t;
}

function publico_(u) { return { usuario: u.usuario, nombre: u.nombre, rol: u.rol, debeCambiar: !!u.debeCambiar }; }

/** Nombres y roles para las tarjetas de inicio (nunca PIN ni cifrados). */
function listaPublica_(libro, conEstado) {
  return leerUsuarios_(libro).lista.filter(function (u) { return u.usuario && (conEstado || u.activo); }).map(function (u) {
    var x = { usuario: u.usuario, nombre: u.nombre, rol: u.rol, activo: u.activo };
    if (conEstado) x.pendiente = !!u.temporal || !u.hash;
    return x;
  });
}

// Bloqueo temporal por intentos fallidos (guardado en las propiedades del script).
function fallos_(usuario) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('fallos:' + usuario) || '{}'); } catch (e) { return {}; }
}
function anotarFallo_(usuario) {
  var f = fallos_(usuario);
  f.n = (f.n || 0) + 1;
  if (f.n >= INTENTOS_PIN) { f.hasta = Date.now() + 15 * 60 * 1000; f.n = 0; }
  PropertiesService.getScriptProperties().setProperty('fallos:' + usuario, JSON.stringify(f));
}
function limpiarFallos_(usuario) { PropertiesService.getScriptProperties().setProperty('fallos:' + usuario, '{}'); }

function pinTemporal_() {
  var n = parseInt(Utilities.getUuid().replace(/[^0-9]/g, '').slice(0, 6) || '0', 10) % 10000;
  return ('000' + n).slice(-4);
}

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
  var error = { ok: false, error: 'PIN incorrecto' };
  if (!u || !nombre) return { ok: false, error: 'Usuario desconocido' };
  if (!u.activo) return { ok: false, error: 'Este usuario está desactivado. Hable con el administrador.' };
  var debeCambiar = false;
  if (!MODO_PRUEBAS) {
    var f = fallos_(u.usuario);
    if (f.hasta && f.hasta > Date.now()) {
      return { ok: false, error: 'Demasiados intentos. Espere ' + Math.ceil((f.hasta - Date.now()) / 60000) + ' minutos o pida a un administrador un PIN nuevo.' };
    }
    if (!clave) return error;
    if (u.temporal && clave === u.temporal) debeCambiar = true;
    else if (!(u.hash && cifrar_(clave, u.sal) === u.hash)) { anotarFallo_(u.usuario); return error; }
    limpiarFallos_(u.usuario);
  }
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
  u.debeCambiar = !MODO_PRUEBAS && (!!u.temporal || !u.hash);
  sh.getRange(fila, 4).setValue(new Date());
  return u;
}

function cambiarContrasena_(libro, u, p) {
  var nueva = String(p.nueva || '');
  if (!/^\d{4}$/.test(nueva)) return { ok: false, error: 'El PIN debe tener exactamente 4 números' };
  if (nueva === u.temporal) return { ok: false, error: 'El PIN nuevo debe ser distinto del temporal' };
  if (/^(\d)\1{3}$/.test(nueva) || '0123456789'.indexOf(nueva) >= 0 || '9876543210'.indexOf(nueva) >= 0) {
    return { ok: false, error: 'Elija un PIN menos fácil de adivinar (no 1111, 1234, 4321…)' };
  }
  var tieneActual = u.hash && !u.temporal;
  if (!u.debeCambiar && tieneActual && !MODO_PRUEBAS && !(cifrar_(String(p.actual || ''), u.sal) === u.hash)) {
    return { ok: false, error: 'El PIN actual no es correcto' };
  }
  var sal = Utilities.getUuid().replace(/-/g, '');
  var us = leerUsuarios_(libro);
  us.sh.getRange(u.fila, 5, 1, 3).setNumberFormat('@').setValues([['', cifrar_(nueva, sal), sal]]);
  agregarFilas_(hoja_(libro, HOJA_BITACORA, COLS_BITACORA), [['clave-' + Date.now(), new Date(), u.nombre, 'cambiar PIN', 'ok', '']]);
  u.debeCambiar = false;
  return { ok: true, usuario: publico_(u) };
}

/**
 * Administración de usuarios desde la aplicación (solo rol admin).
 * op: crear {usuario, nombre, rol} · pin {usuario} (genera PIN temporal) · activo {usuario, activo}
 */
function adminUsuarios_(libro, admin, p) {
  if (admin.rol !== 'admin') return { ok: false, error: 'Solo un administrador puede administrar usuarios' };
  var us = leerUsuarios_(libro);
  var nombreU = String(p.usuario || '').trim().toLowerCase();
  var u = us.lista.filter(function (x) { return x.usuario === nombreU; })[0];
  var registro = function (detalle) {
    agregarFilas_(hoja_(libro, HOJA_BITACORA, COLS_BITACORA), [['usuarios-' + Date.now(), new Date(), admin.nombre, 'usuarios', 'ok', detalle]]);
  };
  if (p.op === 'crear') {
    if (!/^[a-z][a-z0-9._-]{1,29}$/.test(nombreU)) return { ok: false, error: 'El usuario debe empezar con una letra y usar solo minúsculas, números, punto o guion' };
    if (u) return { ok: false, error: 'Ya existe el usuario ' + nombreU };
    var rol = ['admin', 'aplica', 'aprueba'].indexOf(p.rol) >= 0 ? p.rol : 'aplica';
    var nombre = String(p.nombre || '').trim();
    if (!nombre) return { ok: false, error: 'Falta el nombre' };
    var pin = pinTemporal_();
    agregarFilas_(us.sh, [[nombreU, nombre, rol, 'TRUE', pin, '', '']], true);
    registro('creó a ' + nombreU + ' (' + rol + ')');
    return { ok: true, pin: pin, usuarios: listaPublica_(libro, true) };
  }
  if (!u) return { ok: false, error: 'No existe el usuario ' + nombreU };
  if (p.op === 'pin') {
    var nuevo = pinTemporal_();
    us.sh.getRange(u.fila, 5).setNumberFormat('@').setValue(nuevo);
    limpiarFallos_(u.usuario);
    registro('PIN temporal nuevo para ' + u.usuario);
    return { ok: true, pin: nuevo, usuarios: listaPublica_(libro, true) };
  }
  if (p.op === 'activo') {
    if (u.usuario === admin.usuario && !p.activo) return { ok: false, error: 'No puede desactivarse a sí mismo' };
    us.sh.getRange(u.fila, 4).setNumberFormat('@').setValue(p.activo ? 'TRUE' : 'FALSE');
    registro((p.activo ? 'activó a ' : 'desactivó a ') + u.usuario);
    return { ok: true, usuarios: listaPublica_(libro, true) };
  }
  return { ok: false, error: 'Operación desconocida' };
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

function agregarFilas_(sh, filas, comoTexto) {
  if (!filas.length) return;
  var r = sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length);
  if (comoTexto) r.setNumberFormat('@'); // así un PIN como 0123 no se convierte en 123
  r.setValues(filas);
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
  if (us.getLastRow() >= 2) {
    var existentes = leerUsuarios_(libro).lista.map(function (x) { return x.usuario; });
    if (existentes.indexOf('fernando') < 0) {
      agregarFilas_(us, [['fernando', 'Fernando', 'admin', 'TRUE', pinTemporal_(), '', '']], true);
      Logger.log('Se agregó el usuario fernando (administrador). Su PIN temporal está en la pestaña Usuarios.');
    }
  }
  if (us.getLastRow() < 2) {
    us.getRange(2, 1, 4, COLS_USUARIOS.length).setNumberFormat('@').setValues([
      ['jonathan', 'Jonathan Cerros', 'admin', 'TRUE', pinTemporal_(), '', ''],
      ['fernando', 'Fernando', 'admin', 'TRUE', pinTemporal_(), '', ''],
      ['norman', 'Norman Alpízar', 'aprueba', 'TRUE', pinTemporal_(), '', ''],
      ['daniel', 'Daniel Barquero', 'aplica', 'TRUE', pinTemporal_(), '', ''],
    ]);
  }
  var hoja1 = libro.getSheetByName('Hoja 1') || libro.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && libro.getSheets().length > 1) libro.deleteSheet(hoja1);
  SpreadsheetApp.flush();
  Logger.log('Listo. En la pestaña Usuarios está el usuario y el PIN temporal de cada persona.' + (MODO_PRUEBAS ? ' (MODO_PRUEBAS está activo: se entra sin PIN.)' : ''));
  if (!claveServicio_()) Logger.log('ATENCIÓN: falta escribir la contraseña del servicio en CLAVE_SERVICIO, al inicio del código. Sin ella la aplicación no se puede conectar.');
  else if (claveServicio_().length < 10) Logger.log('Sugerencia: use una contraseña del servicio de al menos 10 caracteres.');
  else Logger.log('Contraseña del servicio configurada.');
}
