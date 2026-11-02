// ============================================================================
// Cuentas por Pagar · Grupo BSE · Facturas recibidas por correo (versión 1.21)
//
// Este es un ARCHIVO NUEVO del proyecto de Apps Script (botón «+» junto a «Archivos»
// → «Secuencia de comandos», nombre: Recibidas). Pegue aquí todo este contenido.
//
// Además, en el archivo Código.gs, dentro de la función manejar_, agregue DOS líneas:
//
//  1) Justo debajo de:   if (p.accion === 'ping') return { ok: true };
//         if (p.accion === 'recibirFacturas') return recibirFacturas_(libro, p);
//
//  2) Justo debajo de:   var extra = adjuntos_(p, usuario); if (extra) return extra;
//         var rec = recibidas_(p, usuario); if (rec) return rec;
//
// Guarde e implemente una versión nueva. La hoja «Recibidas» se crea sola.
// ============================================================================

var HOJA_RECIBIDAS = 'Recibidas';
/** Las facturas con fecha anterior a esta se marcan «anterior · revisar» (el archivo de GTI llegaba hasta el 25/10). */
var RECIBIDAS_INICIO = '2026-10-26';
var COLS_RECIBIDAS = ['clave', 'tipo', 'consecutivo', 'fecha', 'cedulaReceptor', 'nombreReceptor', 'cedulaEmisor', 'nombreEmisor',
  'nombreComercial', 'condicion', 'plazo', 'moneda', 'tipoCambio', 'neto', 'impuesto', 'total', 'referencia',
  'estado', 'hacienda', 'haciendaDetalle', 'pdfId', 'pdfNombre', 'recibida', 'origen',
  'condicionManual', 'plazoManual', 'nota', 'facturaId', 'cambiadoPor', 'cambiado'];

// ---------------------------------------------------------------- Lectura del XML de Hacienda

var TIPOS_HACIENDA_ = {
  FacturaElectronica: 'FE', TiqueteElectronico: 'TE', NotaCreditoElectronica: 'NC', NotaDebitoElectronica: 'ND',
  FacturaElectronicaCompra: 'FEC', FacturaElectronicaExportacion: 'FEE', MensajeHacienda: 'MH', MensajeReceptor: 'MR',
};

function xmlDecodificar_(t) {
  return String(t == null ? '' : t)
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, n) { return String.fromCharCode(Number(n)); })
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')
    .trim();
}
/** Contenido de la primera etiqueta con ese nombre (con o sin prefijo de espacio de nombres). */
function xmlEtiqueta_(xml, nombre) {
  var r = new RegExp('<(?:[\\w.-]+:)?' + nombre + '(?:\\s[^>]*)?>([\\s\\S]*?)</(?:[\\w.-]+:)?' + nombre + '\\s*>');
  var m = r.exec(xml || '');
  return m ? m[1] : '';
}
function xmlValor_(xml, nombre) { return xmlDecodificar_(xmlEtiqueta_(xml, nombre)); }

/**
 * Lee un comprobante electrónico de Costa Rica (v4.3 / v4.4).
 * Devuelve null si no es un XML de Hacienda.
 */
function leerComprobante_(xml) {
  xml = String(xml || '').replace(/^﻿/, '');
  var raiz = /<(?:[\w.-]+:)?([A-Za-z]+)[\s>]/.exec(xml.replace(/<\?xml[\s\S]*?\?>/, '').replace(/<!--[\s\S]*?-->/g, ''));
  var tipo = raiz && TIPOS_HACIENDA_[raiz[1]];
  if (!tipo) return null;
  var clave = xmlValor_(xml, 'Clave').replace(/\D/g, '');
  if (tipo === 'MH') {
    var m = xmlValor_(xml, 'Mensaje');
    return { tipo: 'MH', clave: clave, hacienda: m === '1' ? 'aceptada' : m === '2' ? 'parcial' : m === '3' ? 'rechazada' : '',
      haciendaDetalle: xmlValor_(xml, 'DetalleMensaje').slice(0, 500), cedulaEmisor: xmlValor_(xml, 'NumeroCedulaEmisor') };
  }
  if (tipo === 'MR') return { tipo: 'MR', clave: clave };
  var emisor = xmlEtiqueta_(xml, 'Emisor'), receptor = xmlEtiqueta_(xml, 'Receptor'), resumen = xmlEtiqueta_(xml, 'ResumenFactura');
  var monedaXml = xmlEtiqueta_(resumen, 'CodigoTipoMoneda');
  var ref = xmlEtiqueta_(xml, 'InformacionReferencia');
  var fecha = xmlValor_(xml, 'FechaEmision').slice(0, 10);
  return {
    tipo: tipo, clave: clave, consecutivo: xmlValor_(xml, 'NumeroConsecutivo'), fecha: fecha,
    cedulaReceptor: xmlValor_(xmlEtiqueta_(receptor, 'Identificacion'), 'Numero'), nombreReceptor: xmlValor_(receptor, 'Nombre'),
    cedulaEmisor: xmlValor_(xmlEtiqueta_(emisor, 'Identificacion'), 'Numero'), nombreEmisor: xmlValor_(emisor, 'Nombre'),
    nombreComercial: xmlValor_(emisor, 'NombreComercial'),
    condicion: xmlValor_(xml, 'CondicionVenta'), plazo: xmlValor_(xml, 'PlazoCredito').replace(/[^\d]/g, ''),
    moneda: xmlValor_(monedaXml, 'CodigoMoneda') || 'CRC', tipoCambio: xmlValor_(monedaXml, 'TipoCambio') || '1',
    neto: xmlValor_(resumen, 'TotalVentaNeta'), impuesto: xmlValor_(resumen, 'TotalImpuesto') || '0', total: xmlValor_(resumen, 'TotalComprobante'),
    referencia: xmlValor_(ref, 'Numero'),
  };
}

// ---------------------------------------------------------------- Hoja «Recibidas»

function hojaRecibidas_(libro) {
  var sh = libro.getSheetByName(HOJA_RECIBIDAS);
  if (!sh) {
    sh = libro.insertSheet(HOJA_RECIBIDAS);
    sh.getRange(1, 1, 1, COLS_RECIBIDAS.length).setValues([COLS_RECIBIDAS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** { sh, filas: [obj], pos: { clave: númeroDeFila } } */
function leerRecibidas_(libro) {
  var sh = hojaRecibidas_(libro);
  var n = sh.getLastRow() - 1, pos = {}, filas = [];
  var vals = n > 0 ? sh.getRange(2, 1, n, COLS_RECIBIDAS.length).getValues() : [];
  vals.forEach(function (v, i) {
    var o = {};
    COLS_RECIBIDAS.forEach(function (c, j) { o[c] = v[j] instanceof Date ? v[j].toISOString() : String(v[j] == null ? '' : v[j]); });
    if (!o.clave) return;
    filas.push(o); pos[o.clave] = i + 2;
  });
  return { sh: sh, filas: filas, pos: pos };
}

function escribirRecibida_(rec, o) {
  var fila = COLS_RECIBIDAS.map(function (c) { return o[c] == null ? '' : String(o[c]); });
  var formatos = [COLS_RECIBIDAS.map(function () { return '@'; })]; // todo como texto: claves y cédulas no se vuelven números
  var r = rec.pos[o.clave];
  if (!r) { r = rec.sh.getLastRow() + 1; rec.pos[o.clave] = r; rec.filas.push(o); }
  rec.sh.getRange(r, 1, 1, fila.length).setNumberFormats(formatos).setValues([fila]);
}

function buscarRecibida_(rec, clave) {
  for (var i = 0; i < rec.filas.length; i++) if (rec.filas[i].clave === clave) return rec.filas[i];
  return null;
}

/** Guarda el PDF en la carpeta de comprobantes (la del bloque de adjuntos). */
function guardarPdfRecibida_(pdf, o) {
  if (!pdf || !pdf.base64 || typeof carpetaComprobantes_ !== 'function') return null;
  try {
    var bytes = Utilities.base64Decode(pdf.base64);
    if (bytes.length > 8 * 1024 * 1024) return null;
    var nombre = ('Factura ' + (o.consecutivo || o.clave) + ' ' + (o.nombreEmisor || '')).replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 110) + '.pdf';
    var archivo = carpetaComprobantes_().createFile(Utilities.newBlob(bytes, 'application/pdf', nombre));
    archivo.setDescription('Factura recibida por correo · ' + o.clave);
    return { id: archivo.getId(), nombre: nombre };
  } catch (e) { return null; }
}

/**
 * Guarda una lista de documentos [{ xml, pdf: {nombre, base64} }].
 * Devuelve por cada uno { clave, tipo, estado: nueva | repetida | respuesta | ignorado | error, error }.
 */
function guardarDocumentos_(libro, docs, origen, quien) {
  var rec = leerRecibidas_(libro), ahora = new Date().toISOString();
  return (docs || []).map(function (doc) {
    var x = leerComprobante_(doc && doc.xml);
    if (!x) return { estado: 'error', error: 'No es un comprobante electrónico de Hacienda' };
    if (!x.clave || x.clave.length < 40) return { estado: 'error', error: 'El comprobante no trae la clave de Hacienda' };
    if (x.tipo === 'MR' || x.tipo === 'FEE') return { clave: x.clave, tipo: x.tipo, estado: 'ignorado' };
    var ya = buscarRecibida_(rec, x.clave);
    if (x.tipo === 'MH') {
      var o = ya || { clave: x.clave, tipo: 'MH', cedulaEmisor: x.cedulaEmisor, estado: 'respuesta', recibida: ahora, origen: origen };
      o.hacienda = x.hacienda; o.haciendaDetalle = x.haciendaDetalle;
      escribirRecibida_(rec, o);
      return { clave: x.clave, tipo: 'MH', estado: 'respuesta', hacienda: x.hacienda };
    }
    if (ya && ya.tipo !== 'MH') {
      if (!ya.pdfId && doc.pdf) { var p0 = guardarPdfRecibida_(doc.pdf, ya); if (p0) { ya.pdfId = p0.id; ya.pdfNombre = p0.nombre; escribirRecibida_(rec, ya); } }
      return { clave: x.clave, tipo: x.tipo, estado: 'repetida' };
    }
    var nuevo = {};
    for (var k in x) nuevo[k] = x[k];
    nuevo.estado = 'nueva'; nuevo.recibida = ahora; nuevo.origen = origen; nuevo.cambiadoPor = quien || '';
    if (ya) { nuevo.hacienda = ya.hacienda; nuevo.haciendaDetalle = ya.haciendaDetalle; } // la respuesta de Hacienda llegó antes
    var pdf = guardarPdfRecibida_(doc.pdf, nuevo);
    if (pdf) { nuevo.pdfId = pdf.id; nuevo.pdfNombre = pdf.nombre; }
    escribirRecibida_(rec, nuevo);
    return { clave: x.clave, tipo: x.tipo, estado: 'nueva' };
  });
}

// ---------------------------------------------------------------- Entradas

/** El lector de cxp.grupobse (ya validó la contraseña del servicio en manejar_). */
function recibirFacturas_(libro, p) {
  var res = guardarDocumentos_(libro, p.documentos, 'correo', 'Correo');
  // Resumen del día: hora de la última lectura y cuántas facturas nuevas llegaron hoy.
  var props = PropertiesService.getScriptProperties(), ahora = new Date();
  var hoy = Utilities.formatDate(ahora, 'America/Costa_Rica', 'yyyy-MM-dd'), ant = {};
  try { ant = JSON.parse(props.getProperty('RECIBIDAS_ULTIMA') || '{}'); } catch (e) { ant = {}; }
  var mismoDia = ant.dia === hoy;
  props.setProperty('RECIBIDAS_ULTIMA', JSON.stringify({ fecha: ahora.toISOString(), dia: hoy,
    nuevasHoy: (mismoDia ? Number(ant.nuevasHoy) || 0 : 0) + res.filter(function (r) { return r.estado === 'nueva'; }).length,
    correosHoy: (mismoDia ? Number(ant.correosHoy) || 0 : 0) + 1 }));
  return { ok: true, resultados: res };
}

/** Acciones de la aplicación (usuario con sesión). Devuelve null si no le corresponde. */
function recibidas_(p, usuario) {
  if (!/^recibidas/.test(p.accion || '')) return null;
  if (usuario.rol === 'aprueba') return { ok: false, error: 'Solo quienes aplican pagos ven las facturas recibidas' };
  var libro = SpreadsheetApp.getActive();
  if (p.accion === 'recibidasLista') {
    var ultima = PropertiesService.getScriptProperties().getProperty('RECIBIDAS_ULTIMA');
    return { ok: true, filas: leerRecibidas_(libro).filas, inicio: RECIBIDAS_INICIO, ultima: ultima ? JSON.parse(ultima) : null };
  }
  if (p.accion === 'recibidasSubir') {
    if ((p.documentos || []).length > 40) return { ok: false, error: 'Máximo 40 archivos por envío' };
    return { ok: true, resultados: guardarDocumentos_(libro, p.documentos, 'subida', usuario.nombre) };
  }
  if (p.accion === 'recibidasCambiar') {
    var rec = leerRecibidas_(libro), c = p.cambios || {}, ahora = new Date().toISOString(), n = 0;
    var permitidos = ['estado', 'condicionManual', 'plazoManual', 'nota', 'facturaId'];
    if (c.estado && ['nueva', 'descartada', 'cxp'].indexOf(c.estado) < 0) return { ok: false, error: 'Estado no válido' };
    (p.claves || []).forEach(function (clave) {
      var o = buscarRecibida_(rec, String(clave));
      if (!o) return;
      var extra = (p.porClave || {})[clave] || {};
      permitidos.forEach(function (k) { if (k in c) o[k] = c[k]; if (k in extra) o[k] = extra[k]; });
      o.cambiadoPor = usuario.nombre; o.cambiado = ahora;
      escribirRecibida_(rec, o); n++;
    });
    return { ok: true, cambiadas: n };
  }
  return { ok: false, error: 'Acción desconocida: ' + p.accion };
}
