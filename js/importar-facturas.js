// Importación de facturas nuevas desde un CSV o un Excel.
// Reconoce el reporte de GTI ("DocumentosAceptacion…", con o sin encabezados) y cualquier otro
// archivo por el título de sus columnas; el usuario puede corregirlas.
//
// GTI no trae el código contable del proveedor, sino su cédula jurídica. La aplicación busca el
// proveedor por la cédula (si ya se importó antes) o por el nombre, y el usuario confirma o elige.
// La cédula queda escrita en las observaciones de la factura ("Céd. 3101…"), así la próxima vez
// —en cualquier computadora— el proveedor se reconoce solo.

import { excelSerialToISO, parseFecha, parseNum, round2, sumarDias } from './util.js';

export const CAMPOS = [
  { k: 'sociedad', n: 'Sociedad', fijo: true },
  { k: 'cedulaReceptor', n: 'Cédula de la sociedad', opcional: true },
  { k: 'codigo', n: 'Código de proveedor', opcional: true },
  { k: 'cedulaEmisor', n: 'Cédula del proveedor', opcional: true },
  { k: 'proveedor', n: 'Nombre del proveedor' },
  { k: 'tipoDoc', n: 'Tipo de documento', opcional: true },
  { k: 'documento', n: 'N° de factura' },
  { k: 'fechaFactura', n: 'Fecha de factura' },
  { k: 'fechaVence', n: 'Fecha de vencimiento', opcional: true },
  { k: 'moneda', n: 'Moneda', fijo: true },
  { k: 'monto', n: 'Monto' },
  { k: 'observaciones', n: 'Observaciones', opcional: true },
];

/** Cédulas jurídicas conocidas de las sociedades (por su nombre corto). */
export const CEDULAS_SOCIEDAD = { 3101006847: 'Santa Eduviges', 3101006133: 'La Hilda' };

// Columnas del reporte de GTI. El CSV que se ajusta para el sistema contable trae estas 13, sin encabezado.
export const GTI_CSV = ['TipoDocRecibido', 'FechaEmisiónDocumento', 'ConsecutivoDocumento', 'IdentificaciónReceptor',
  'NombreReceptor', 'IdentificaciónEmisor', 'NombreEmisor', 'TotalNeto', 'TotalImpuesto', 'TotalComprobante',
  'CódigoMoneda', 'TipoCambio', 'Total Colones'];
const GTI_MAPEO = {
  tipoDoc: 'tipodocrecibido', fechaFactura: 'fechaemisiondocumento', documento: 'consecutivodocumento',
  cedulaReceptor: 'identificacionreceptor', sociedad: 'nombrereceptor', cedulaEmisor: 'identificacionemisor',
  proveedor: 'nombreemisor', monto: 'totalcomprobante', moneda: 'codigomoneda',
};

const PISTAS = {
  sociedad: ['sociedad', 'empresa', 'compania', 'nombre receptor', 'receptor'],
  cedulaReceptor: ['identificacion receptor', 'cedula receptor', 'cedula sociedad'],
  codigo: ['codigo proveedor', 'cod proveedor', 'codigo', 'cod', 'id proveedor'],
  cedulaEmisor: ['identificacion emisor', 'cedula emisor', 'cedula proveedor', 'cedula juridica', 'cedula'],
  proveedor: ['nombre del proveedor', 'proveedor', 'nombre emisor', 'emisor', 'razon social', 'nombre'],
  tipoDoc: ['tipo doc', 'tipo de documento', 'tipo documento'],
  documento: ['consecutivo documento', 'numero documento', 'numero de factura', 'factura', 'documento', 'consecutivo', 'numero', 'no factura', 'n factura'],
  fechaFactura: ['fecha emision documento', 'fecha factura', 'fec factura', 'fecha de factura', 'fecha emision', 'fecha documento', 'fecha'],
  fechaVence: ['fecha vence', 'fec vence', 'vencimiento', 'vence', 'fecha de vencimiento'],
  moneda: ['codigo moneda', 'moneda', 'divisa', 'tipo moneda'],
  monto: ['total comprobante', 'monto total', 'total factura', 'monto', 'total', 'importe', 'valor'],
  observaciones: ['observaciones', 'detalle', 'descripcion', 'nota', 'comentario'],
};

const normal = (t) => String(t ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
  .replace(/[^a-z0-9]+/g, ' ').trim();
const pegado = (t) => normal(t).replace(/ /g, '');
const soloDigitos = (v) => String(v ?? '').replace(/\D/g, '');

/** Lee el archivo (CSV con coma o punto y coma, o Excel) y devuelve las filas como listas de celdas. */
export function leerFilas(contenido, XLSX, esTexto) {
  let wb;
  if (esTexto) {
    const texto = contenido.replace(/^﻿/, '');
    const primera = texto.split(/\r?\n/).find((l) => l.trim()) || '';
    const sep = (primera.match(/;/g) || []).length > (primera.match(/,/g) || []).length ? ';' : ',';
    wb = XLSX.read(texto, { type: 'string', raw: true, FS: sep });
  } else {
    wb = XLSX.read(contenido, { type: 'array' });
  }
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' })
    .map((f) => f.map((c) => (typeof c === 'string' ? c.replace(/�/g, 'ó').trim() : c)))
    .filter((f) => f.some((c) => String(c).trim() !== ''));
}

/** Primera fila que parece de encabezados (al menos 3 textos sin números de documento). */
export function filaEncabezado(filas) {
  const i = filas.findIndex((f) => f.filter((c) => typeof c === 'string' && /[a-zA-Z]/.test(c) && !/\d{6,}/.test(c)).length >= 3);
  return i < 0 ? 0 : i;
}

/**
 * Reconoce el reporte de GTI. Devuelve { filas, inicio, mapeo } (con encabezado agregado si el CSV
 * no lo trae) o null si el archivo tiene otro formato.
 */
export function detectarGTI(filas) {
  const conEnc = filas.findIndex((f) => {
    const n = f.map(pegado);
    return n.includes('consecutivodocumento') && n.includes('identificacionemisor');
  });
  if (conEnc >= 0) return { filas, inicio: conEnc, mapeo: mapeoGTI(filas[conEnc]) };
  const pareceGTI = (f) => f.length >= 11 && /factura|nota|tiquete/i.test(String(f[0])) && /^\d{20}$/.test(soloDigitos(f[2])) && /^\d{9,12}$/.test(soloDigitos(f[5]));
  if (filas.length && filas.every(pareceGTI)) {
    return { filas: [GTI_CSV, ...filas], inicio: 0, mapeo: mapeoGTI(GTI_CSV) };
  }
  return null;
}

function mapeoGTI(enc) {
  const n = enc.map(pegado);
  const mapeo = {};
  for (const [k, h] of Object.entries(GTI_MAPEO)) { const i = n.indexOf(h); if (i >= 0) mapeo[k] = i; }
  return mapeo;
}

/** Propone qué columna corresponde a cada dato según el nombre del encabezado. */
export function sugerirMapeo(encabezados) {
  const gti = mapeoGTI(encabezados);
  if ('documento' in gti && 'cedulaEmisor' in gti) return gti;
  const enc = encabezados.map(normal);
  const usado = new Set();
  const mapeo = {};
  for (const c of CAMPOS) {
    for (const pista of PISTAS[c.k]) {
      const i = enc.findIndex((h, j) => !usado.has(j) && (h === pista || h.includes(pista)));
      if (i >= 0) { mapeo[c.k] = i; usado.add(i); break; }
    }
  }
  return mapeo;
}

/**
 * Número de factura como se registra en la aplicación: el consecutivo electrónico de 20 dígitos
 * (sucursal + terminal + tipo + número) queda solo con el número, sin ceros: 00100001010000107091 → 107091.
 */
export function numeroDocumento(v) {
  const t = String(v ?? '').trim();
  if (/^\d{20}$/.test(t)) return String(Number(t.slice(10)));
  return t;
}
const docComparable = (d) => String(d ?? '').trim().replace(/^0+(?=\d)/, '').toUpperCase();

function fecha(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v > 20000 && v < 80000 ? excelSerialToISO(v) : '';
  return parseFecha(String(v).trim().split(/\s+/)[0]) || '';
}

function moneda(v) {
  const t = normal(v);
  if (!t) return '';
  if (/usd|dolar|\$|^us/.test(t) || String(v).includes('$')) return 'USD';
  if (/crc|col|colon/.test(t) || String(v).includes('₡')) return 'CRC';
  return '';
}

/** Nombre para comparar: sin espacios, puntos, comas ni "S.A.", "SRL", "Sociedad Anónima"… */
export function nombreClave(n) {
  let s = pegado(n);
  for (let k = 0; k < 2; k++) s = s.replace(/(sociedadanonima|sociedadderesponsabilidadlimitada|limitada|ltda|srl|sa)$/, '');
  return s;
}

/** Cédula → código de proveedor, aprendido de las facturas importadas antes ("Céd. 3101…"). */
export function cedulasConocidas(datos) {
  const m = {};
  for (const f of Object.values(datos.facturas)) {
    const x = /C[eé]d\.\s*(\d{9,12})/.exec(f.observaciones || '');
    if (x && f.codigo && datos.proveedores[f.codigo]) m[x[1]] = f.codigo;
  }
  return m;
}

/** Busca un proveedor por nombre (exacto sin puntuación, o uno contenido en el otro). */
function buscarPorNombre(nombre, provs) {
  const c = nombreClave(nombre);
  if (c.length < 3) return null;
  const exacto = provs.filter((p) => nombreClave(p.nombre) === c || (p.alias && nombreClave(p.alias) === c));
  if (exacto.length === 1) return exacto[0];
  // Mismas palabras en otro orden: "CESAR ALBERTO CARRANZA GARITA" = "CARRANZA GARITA CESAR ALBERTO"
  const palabras = (t) => normal(t).split(' ').filter((w) => w.length > 1 && !/^(s|a|sa|srl|ltda|de|del|la|y|e)$/.test(w)).sort().join(' ');
  const pc = palabras(nombre);
  if (pc.split(' ').length >= 2) {
    const mismas = provs.filter((p) => palabras(p.nombre) === pc);
    if (mismas.length === 1) return mismas[0];
  }
  if (c.length < 6) return null;
  const parecidos = provs.filter((p) => { const q = nombreClave(p.nombre); return q.length >= 6 && (q.includes(c) || c.includes(q)); });
  return parecidos.length === 1 ? parecidos[0] : null;
}

/**
 * Convierte las filas en facturas listas para agregar.
 * @param filas      todas las filas del archivo
 * @param inicio     índice de la fila de encabezados
 * @param mapeo      { campo: índiceColumna }
 * @param fijos      { sociedad, moneda } cuando no vienen en una columna
 * @param datos      { facturas, proveedores, sociedades } de la aplicación
 * @param opciones   { asignados: { claveEmisor: código | '' (nuevo) }, memoria: { cédula: código } }
 * @returns { filas: [{n, estado, mensaje, factura, proveedor, proveedorNuevo, emisor}], emisores, resumen }
 *          estado: 'nueva' | 'existe' | 'omitida' | 'error'
 */
export function prepararFacturas(filas, inicio, mapeo, fijos, datos, opciones = {}) {
  const asignados = opciones.asignados || {};
  const memoria = { ...cedulasConocidas(datos), ...(opciones.memoria || {}) };
  const tiene = (k) => mapeo[k] != null && mapeo[k] !== '';
  const col = (fila, k) => (tiene(k) ? fila[Number(mapeo[k])] : '');
  const socs = Object.values(datos.sociedades);
  const socPorCorto = (corto) => socs.find((s) => normal(s.corto) === normal(corto))?.clave || null;
  const buscarSociedad = (ced, nombre) => {
    const c = soloDigitos(ced);
    if (c && CEDULAS_SOCIEDAD[c]) { const s = socPorCorto(CEDULAS_SOCIEDAD[c]); if (s) return s; }
    const t = pegado(nombre);
    if (!t) return null;
    return socs.find((s) => pegado(s.clave) === t || pegado(s.corto) === t || t.includes(pegado(s.corto)) || pegado(s.clave).includes(t))?.clave || null;
  };
  const provs = Object.values(datos.proveedores);
  const facturas = Object.values(datos.facturas);
  const existentes = new Set(facturas.map((f) => [f.sociedad, f.codigo, docComparable(f.documento), f.moneda].join('|')));
  const mismoNumeroYMonto = new Set(facturas.map((f) => [f.sociedad, docComparable(f.documento), round2(f.monto)].join('|')));
  const enArchivo = new Set();
  let siguienteCodigo = Math.max(0, ...provs.map((p) => Number(p.codigo)).filter(Number.isFinite)) + 1;
  const emisores = new Map(); // clave → { clave, cedula, nombre, codigo, como, filas }

  /** Decide el proveedor de cada emisor una sola vez. */
  const emisorDe = (cedula, nombre, codigoArchivo) => {
    const clave = cedula || (nombre ? 'n:' + nombreClave(nombre) : 'c:' + codigoArchivo);
    if (emisores.has(clave)) return emisores.get(clave);
    let codigo = '';
    let como = 'nuevo';
    if (clave in asignados) { codigo = asignados[clave]; como = codigo ? 'elegido' : 'nuevo'; }
    else if (codigoArchivo && datos.proveedores[codigoArchivo]) { codigo = codigoArchivo; como = 'codigo'; }
    else if (cedula && memoria[cedula] && datos.proveedores[memoria[cedula]]) { codigo = memoria[cedula]; como = 'cedula'; }
    else { const p = nombre ? buscarPorNombre(nombre, provs) : null; if (p) { codigo = p.codigo; como = 'nombre'; } }
    const e = { clave, cedula, nombre, codigo, como, filas: 0, nuevo: null };
    if (!codigo && nombre) {
      const cod = codigoArchivo && !datos.proveedores[codigoArchivo] ? codigoArchivo : String(siguienteCodigo++);
      e.nuevo = { codigo: cod, nombre: nombreLegible(nombre), alias: '', ibanCRC: '', ibanUSD: '', plazo: 30 };
    }
    emisores.set(clave, e);
    return e;
  };

  const salida = [];
  for (let i = inicio + 1; i < filas.length; i++) {
    const fila = filas[i];
    const r = { n: i + (filas[0] === GTI_CSV ? 0 : 1), estado: 'nueva', mensaje: '' };
    const tipo = String(col(fila, 'tipoDoc') ?? '').trim();
    const sociedad = tiene('sociedad') || tiene('cedulaReceptor')
      ? buscarSociedad(col(fila, 'cedulaReceptor'), col(fila, 'sociedad')) || (tiene('sociedad') ? null : fijos.sociedad)
      : fijos.sociedad;
    const mon = tiene('moneda') ? moneda(col(fila, 'moneda')) : fijos.moneda;
    const documento = numeroDocumento(col(fila, 'documento'));
    const nombreProv = String(col(fila, 'proveedor') ?? '').trim().replace(/\s{2,}/g, ' ');
    const codigoArchivo = String(col(fila, 'codigo') ?? '').trim();
    const cedula = soloDigitos(col(fila, 'cedulaEmisor'));
    const fFactura = fecha(col(fila, 'fechaFactura'));
    const monto = round2(parseNum(col(fila, 'monto')));

    // Filas de totales o vacías
    if (!documento && !nombreProv) continue;
    if (/^total/i.test(documento) || /^total/i.test(nombreProv)) continue;

    const e = (nombreProv || codigoArchivo || cedula) ? emisorDe(cedula, nombreProv, codigoArchivo) : null;
    const prov = e ? (e.codigo ? datos.proveedores[e.codigo] : e.nuevo) : null;

    const errores = [];
    if (!sociedad) errores.push('sociedad no reconocida');
    if (!mon) errores.push('moneda no reconocida');
    if (!documento) errores.push('falta el número de factura');
    if (!prov) errores.push('falta el proveedor');
    if (!fFactura) errores.push('fecha de factura inválida');
    if (!(monto > 0) && !/nota/i.test(tipo)) errores.push('monto inválido');

    const fVence = fecha(col(fila, 'fechaVence')) || (fFactura ? sumarDias(fFactura, Number(prov?.plazo) || 30) : '');
    const obs = [String(col(fila, 'observaciones') ?? '').trim(), cedula ? `Céd. ${cedula}` : ''].filter(Boolean).join(' · ');
    const factura = {
      sociedad, codigo: prov?.codigo || '', documento, moneda: mon, fechaFactura: fFactura, fechaVence: fVence,
      monto, observaciones: obs,
    };
    const clave = [sociedad, factura.codigo, docComparable(documento), mon].join('|');
    if (/nota\s*de\s*cr|nota\s*cr/i.test(tipo) || (monto < 0 && !errores.length)) {
      r.estado = 'omitida'; r.mensaje = 'nota de crédito: aplíquela en la factura que corresponde';
    } else if (/nota/i.test(tipo)) {
      r.estado = 'omitida'; r.mensaje = `${tipo.toLowerCase()}: regístrela a mano`;
    } else if (errores.length) { r.estado = 'error'; r.mensaje = errores.join(', '); }
    else if (existentes.has(clave)) { r.estado = 'existe'; r.mensaje = 'ya está registrada'; }
    else if (mismoNumeroYMonto.has([sociedad, docComparable(documento), monto].join('|'))) { r.estado = 'existe'; r.mensaje = 'ya hay una factura con ese número y monto'; }
    else if (enArchivo.has(clave)) { r.estado = 'existe'; r.mensaje = 'repetida en el archivo'; }
    else enArchivo.add(clave);
    if (e && r.estado === 'nueva') e.filas++;
    Object.assign(r, { factura, proveedor: prov, proveedorNuevo: e && !e.codigo ? e.nuevo : null, nombreProv, emisor: e?.clave, tipo });
    salida.push(r);
  }
  const lista = [...emisores.values()];
  const resumen = {
    nuevas: salida.filter((x) => x.estado === 'nueva').length,
    existentes: salida.filter((x) => x.estado === 'existe').length,
    omitidas: salida.filter((x) => x.estado === 'omitida').length,
    errores: salida.filter((x) => x.estado === 'error').length,
    proveedoresNuevos: lista.filter((e) => !e.codigo && e.nuevo && e.filas > 0).map((e) => e.nuevo),
  };
  return { filas: salida, emisores: lista, resumen };
}

/** "DistribuidoraAgroComercialS.A." → "DISTRIBUIDORA AGRO COMERCIAL S.A." */
function nombreLegible(n) {
  return String(n).replace(/([a-záéíóúñ])([A-ZÁÉÍÓÚÑ])/g, '$1 $2').replace(/([a-zA-Z])(S\.A\.|S\.R\.L\.)$/i, '$1 $2')
    .replace(/\s{2,}/g, ' ').trim().toUpperCase();
}
