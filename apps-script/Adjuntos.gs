// ============================================================================
// Cuentas por Pagar · Grupo BSE · Adjuntos y anotaciones (versión 1.18)
//
// Pegue este bloque completo AL FINAL del script, y agregue en la función
// manejar_ esta línea justo debajo de la que dice  if (p.accion === 'quienSoy') ...
//
//     var extra = adjuntos_(p, usuario); if (extra) return extra;
//
// Después elija la función autorizarDrive y haga clic en Ejecutar (una sola vez,
// para dar permiso de usar Google Drive) e implemente una versión nueva.
// ============================================================================

var CARPETA_COMPROBANTES = 'CxP Comprobantes';
var TAMANO_MAXIMO_ARCHIVO = 8 * 1024 * 1024; // 8 MB

/** Ejecutar una vez desde el editor para dar permiso de usar Google Drive. */
function autorizarDrive() {
  var c = carpetaComprobantes_();
  Logger.log('Carpeta lista: ' + c.getName() + ' (' + c.getUrl() + ')');
}

/** Atiende las acciones de archivos y habilita la operación «anotar». Devuelve null si no le corresponde. */
function adjuntos_(p, usuario) {
  instalarAnotar_();
  if (p.accion === 'subirArchivo') return subirArchivo_(p, usuario);
  if (p.accion === 'leerArchivo') return leerArchivo_(p);
  return null;
}

/**
 * Operación «anotar»: guarda en la pestaña Config un dato compartido entre todos los dispositivos
 * (enviado al banco, N° de documento por proveedor, comprobantes adjuntos). Valor vacío = sin dato.
 */
function instalarAnotar_() {
  if (Logica.__anotar) return;
  var original = Logica.aplicarOp;
  Logica.aplicarOp = function (d, op, ctx) {
    if (!op || op.tipo !== 'anotar') return original(d, op, ctx);
    var clave = String(op.clave || '');
    if (!/^(envio|adj|doc):/.test(clave)) return { ok: false, error: 'Anotación no permitida', cambios: [] };
    d.config[clave] = { clave: clave, valor: String(op.valor == null ? '' : op.valor), v: ctx.ahora };
    return { ok: true, cambios: [['config', clave]] };
  };
  Logica.__anotar = true;
}

function carpetaComprobantes_() {
  var it = DriveApp.getFoldersByName(CARPETA_COMPROBANTES);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_COMPROBANTES);
}

function subirArchivo_(p, usuario) {
  if (!p.base64) return { ok: false, error: 'Falta el archivo' };
  var bytes = Utilities.base64Decode(p.base64);
  if (bytes.length > TAMANO_MAXIMO_ARCHIVO) return { ok: false, error: 'El archivo pesa más de 8 MB' };
  var tipo = /^(application\/pdf|image\/(png|jpeg))$/.test(p.tipo || '') ? p.tipo : 'application/pdf';
  var nombre = String(p.nombre || 'comprobante').replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 120);
  var archivo = carpetaComprobantes_().createFile(Utilities.newBlob(bytes, tipo, nombre));
  archivo.setDescription('Subido por ' + usuario.nombre + ' · ' + new Date().toISOString());
  return { ok: true, id: archivo.getId(), nombre: nombre, tipo: tipo, tamano: bytes.length };
}

function leerArchivo_(p) {
  var archivo;
  try { archivo = DriveApp.getFileById(String(p.id || '')); } catch (e) { return { ok: false, error: 'No se encontró el comprobante' }; }
  // Solo se entregan archivos de la carpeta de comprobantes.
  var padres = archivo.getParents(), dentro = false;
  while (padres.hasNext()) if (padres.next().getName() === CARPETA_COMPROBANTES) dentro = true;
  if (!dentro) return { ok: false, error: 'Ese archivo no es un comprobante' };
  var blob = archivo.getBlob();
  return { ok: true, nombre: archivo.getName(), tipo: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
}
