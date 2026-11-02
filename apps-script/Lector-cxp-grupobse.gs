// ============================================================================
// Cuentas por Pagar · Grupo BSE · LECTOR DE FACTURAS (versión 1.21)
//
// Va en la cuenta cxp.grupobse@gmail.com (NO en la hoja de cuentas por pagar):
//   1. Entre a https://script.google.com con cxp.grupobse@gmail.com → «Nuevo proyecto».
//   2. Borre lo que trae y pegue todo este archivo. Nombre del proyecto: Lector de facturas.
//   3. Escriba abajo la dirección del servicio y la contraseña del servicio (las mismas de la aplicación).
//   4. Arriba elija la función «instalar» y haga clic en Ejecutar. Acepte los permisos.
//
// Desde entonces, cada hora: lee los correos reenviados desde facturacion@bse.cr, manda los XML de
// Hacienda (y su PDF) a la aplicación y manda el correo a la papelera. Si un correo no se pudo leer,
// queda con la etiqueta «CxP-revisar» y no se borra.
// ============================================================================

var URL_SERVICIO = 'PEGUE-AQUI-LA-DIRECCION-DEL-SERVICIO';   // https://script.google.com/macros/s/…/exec
var CLAVE_SERVICIO = 'PEGUE-AQUI-LA-CONTRASEÑA-DEL-SERVICIO';
var ETIQUETA_REVISAR = 'CxP-revisar';
var CORREOS_POR_VUELTA = 40;

/** Ejecutar UNA vez: revisa la conexión, programa la lectura cada hora y lee lo que ya hay. */
function instalar() {
  probarConexion();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'procesar') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('procesar').timeBased().everyHours(1).create();
  Logger.log('Listo: se revisará el correo cada hora.');
  procesar();
}

function probarConexion() {
  var r = llamarServicio_({ accion: 'ping' });
  if (!r || !r.ok) throw new Error('No se pudo conectar con la aplicación: ' + ((r && r.error) || 'revise la dirección y la contraseña del servicio'));
  Logger.log('Conexión con la aplicación: correcta.');
}

/** Lee los correos pendientes (lo llama el reloj cada hora; también se puede ejecutar a mano). */
function procesar() {
  var etiqueta = GmailApp.getUserLabelByName(ETIQUETA_REVISAR) || GmailApp.createLabel(ETIQUETA_REVISAR);
  var hilos = GmailApp.search('in:inbox -label:' + ETIQUETA_REVISAR, 0, CORREOS_POR_VUELTA);
  var resumen = { correos: 0, nuevas: 0, repetidas: 0, respuestas: 0, sinFactura: 0, revisar: 0 };
  for (var i = 0; i < hilos.length; i++) {
    var mensajes = hilos[i].getMessages(), revisar = false;
    for (var j = 0; j < mensajes.length; j++) {
      var m = mensajes[j];
      if (m.isInTrash()) continue;
      resumen.correos++;
      var docs;
      try { docs = documentosDelCorreo_(m.getRawContent()); } catch (e) { revisar = true; continue; }
      if (!docs.length) { resumen.sinFactura++; m.moveToTrash(); continue; }
      var r = llamarServicio_({ accion: 'recibirFacturas', documentos: docs });
      if (!r || !r.ok) { Logger.log('La aplicación no respondió; se intentará en la próxima vuelta. ' + ((r && r.error) || '')); return resumen; }
      var conError = false;
      (r.resultados || []).forEach(function (x) {
        if (x.estado === 'nueva') resumen.nuevas++;
        else if (x.estado === 'repetida') resumen.repetidas++;
        else if (x.estado === 'respuesta') resumen.respuestas++;
        else if (x.estado === 'error') conError = true;
      });
      if (conError) revisar = true; else m.moveToTrash();
    }
    if (revisar) { hilos[i].addLabel(etiqueta); hilos[i].markRead(); resumen.revisar++; }
  }
  Logger.log(JSON.stringify(resumen));
  return resumen;
}

function llamarServicio_(datos) {
  datos.servicio = CLAVE_SERVICIO;
  try {
    var resp = UrlFetchApp.fetch(URL_SERVICIO, { method: 'post', contentType: 'text/plain', payload: JSON.stringify(datos), muteHttpExceptions: true, followRedirects: true });
    return JSON.parse(resp.getContentText());
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

// ---------------------------------------------------------------- Lectura del correo (MIME)

/**
 * Del correo completo saca los XML de Hacienda y sus PDF, aunque vengan reenviados como
 * adjunto (.eml) dentro de otro correo. Devuelve [{ xml, pdf: {nombre, base64} | null }].
 */
function documentosDelCorreo_(raw) {
  var partes = partesMime_(String(raw || ''), 0);
  var xmls = [], pdfs = [];
  partes.forEach(function (p) {
    if (p.tipo === 'xml') {
      var texto = textoDeBase64_(p.base64);
      if (/<(?:[\w.-]+:)?(FacturaElectronica|TiqueteElectronico|NotaCreditoElectronica|NotaDebitoElectronica|FacturaElectronicaCompra|MensajeHacienda)[\s>]/.test(texto)) {
        xmls.push({ nombre: p.nombre, xml: texto, mh: /<(?:[\w.-]+:)?MensajeHacienda[\s>]/.test(texto),
          clave: ((/<(?:[\w.-]+:)?Clave>\s*(\d+)/.exec(texto) || [])[1] || ''),
          consecutivo: ((/<(?:[\w.-]+:)?NumeroConsecutivo>\s*(\d+)/.exec(texto) || [])[1] || '') });
      }
    } else if (p.tipo === 'pdf') pdfs.push(p);
  });
  // PDF de cada factura: el que lleva su clave o consecutivo en el nombre; si el correo trae una sola factura, su PDF.
  var usados = {}, comprobantes = xmls.filter(function (x) { return !x.mh; });
  var docs = xmls.map(function (x) {
    if (x.mh) return { xml: x.xml, pdf: null };
    var pdf = null;
    for (var k = 0; k < pdfs.length && !pdf; k++) {
      var n = pdfs[k].nombre || '';
      if (!usados[k] && ((x.clave && n.indexOf(x.clave) >= 0) || (x.consecutivo && n.indexOf(x.consecutivo) >= 0) || (x.nombre && n.replace(/\.pdf$/i, '') === x.nombre.replace(/\.xml$/i, '')))) { pdf = pdfs[k]; usados[k] = true; }
    }
    if (!pdf && comprobantes.length === 1) for (var q = 0; q < pdfs.length && !pdf; q++) if (!usados[q]) { pdf = pdfs[q]; usados[q] = true; }
    return { xml: x.xml, pdf: pdf ? { nombre: pdf.nombre, base64: pdf.base64 } : null };
  });
  return docs;
}

function partesMime_(raw, prof) {
  if (prof > 8) return [];
  var m = /\r?\n\r?\n/.exec(raw);
  var cab = m ? raw.slice(0, m.index) : raw, cuerpo = m ? raw.slice(m.index + m[0].length) : '';
  var h = cabeceras_(cab);
  var ct = h['content-type'] || 'text/plain';
  var tipo = ct.split(';')[0].trim().toLowerCase();
  var cte = (h['content-transfer-encoding'] || '').trim().toLowerCase();
  var nombre = parametro_(h['content-disposition'] || '', 'filename') || parametro_(ct, 'name') || '';
  if (/^multipart\//.test(tipo)) {
    var b = parametro_(ct, 'boundary');
    if (!b) return [];
    var trozos = cuerpo.split('--' + b), out = [];
    for (var i = 1; i < trozos.length; i++) {
      var t = trozos[i];
      if (/^--/.test(t)) break;
      out = out.concat(partesMime_(t.replace(/^[ \t]*\r?\n/, ''), prof + 1));
    }
    return out;
  }
  if (tipo === 'message/rfc822' || /\.eml$/i.test(nombre)) {
    return partesMime_(cte === 'base64' ? textoDeBase64_(cuerpo.replace(/\s+/g, '')) : cte === 'quoted-printable' ? qpTexto_(cuerpo) : cuerpo, prof + 1);
  }
  var esXml = /\.xml$/i.test(nombre) || /^(application|text)\/xml$/.test(tipo);
  var esPdf = /\.pdf$/i.test(nombre) || tipo === 'application/pdf';
  if (!esXml && !esPdf) return [];
  var b64;
  if (cte === 'base64') b64 = cuerpo.replace(/[^A-Za-z0-9+/=]/g, '');
  else if (cte === 'quoted-printable') b64 = Utilities.base64Encode(Utilities.newBlob(qpTexto_(cuerpo)).getBytes());
  else b64 = Utilities.base64Encode(Utilities.newBlob(cuerpo.replace(/\r?\n$/, '')).getBytes());
  return [{ nombre: nombre, tipo: esXml ? 'xml' : 'pdf', base64: b64 }];
}

function cabeceras_(texto) {
  var h = {}, ultimo = null;
  texto.split(/\r?\n/).forEach(function (l) {
    if (/^[ \t]/.test(l) && ultimo) { h[ultimo] += ' ' + l.trim(); return; }
    var i = l.indexOf(':');
    if (i > 0) { ultimo = l.slice(0, i).trim().toLowerCase(); h[ultimo] = l.slice(i + 1).trim(); }
  });
  return h;
}

/** Valor de un parámetro de cabecera: name="x.xml", filename*=UTF-8''x.xml… */
function parametro_(valor, nombre) {
  // Nombre largo partido en pedazos (RFC 2231): filename*0*=utf-8''FE_00100…; filename*1*=…_XML.xml
  var piezas = [], rp = new RegExp(nombre + '\\*(\\d+)(\\*?)=("([^"]*)"|[^;]+)', 'ig'), mp;
  while ((mp = rp.exec(valor))) piezas[Number(mp[1])] = { v: (mp[4] != null ? mp[4] : mp[3]).trim(), cod: mp[2] === '*' };
  if (piezas.length) {
    var junto = piezas.map(function (x, i) {
      if (!x) return '';
      var v = x.cod && i === 0 ? v0_(x.v) : x.v;
      if (!x.cod) return v;
      try { return decodeURIComponent(v); } catch (e) { return v; }
    }).join('');
    return junto;
  }
  var ext = new RegExp(nombre + "\\*=([^']*)'[^']*'([^;]+)", 'i').exec(valor);
  if (ext) { try { return decodeURIComponent(ext[2].trim()); } catch (e) { return ext[2].trim(); } }
  var m = new RegExp('(?:^|;)\\s*' + nombre + '\\s*=\\s*("([^"]*)"|[^;]+)', 'i').exec(valor);
  if (!m) return '';
  var v = (m[2] != null ? m[2] : m[1]).trim();
  var enc = /^=\?([^?]+)\?([bq])\?([^?]*)\?=$/i.exec(v); // =?UTF-8?B?...?=
  if (enc) {
    try { v = enc[2].toLowerCase() === 'b' ? Utilities.newBlob(Utilities.base64Decode(enc[3])).getDataAsString(enc[1]) : qpTexto_(enc[3].replace(/_/g, ' ')); } catch (e) { /* se deja como venía */ }
  }
  return v;
}

/** Quita el «utf-8''» del primer pedazo codificado. */
function v0_(v) { var i = v.indexOf("''"); return i >= 0 ? v.slice(i + 2) : v; }

function textoDeBase64_(b64) {
  return Utilities.newBlob(Utilities.base64Decode(String(b64).replace(/\s+/g, ''))).getDataAsString('UTF-8');
}

function qpTexto_(t) {
  var bytes = [], s = t.replace(/=\r?\n/g, '');
  for (var i = 0; i < s.length; i++) {
    var b = s.charCodeAt(i) & 255;
    if (s[i] === '=' && /^[0-9A-F]{2}$/i.test(s.substr(i + 1, 2))) { b = parseInt(s.substr(i + 1, 2), 16); i += 2; }
    bytes.push(b > 127 ? b - 256 : b); // Apps Script usa bytes con signo
  }
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}
