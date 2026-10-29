// Importación de facturas nuevas desde un CSV (delimitado por comas) o un Excel.
// La aplicación reconoce las columnas por su encabezado; el usuario puede corregirlas.

import { excelSerialToISO, parseFecha, parseNum, round2, sumarDias } from './util.js';

export const CAMPOS = [
  { k: 'sociedad', n: 'Sociedad', fijo: true },
  { k: 'codigo', n: 'Código de proveedor', opcional: true },
  { k: 'proveedor', n: 'Nombre del proveedor' },
  { k: 'documento', n: 'N° de factura' },
  { k: 'fechaFactura', n: 'Fecha de factura' },
  { k: 'fechaVence', n: 'Fecha de vencimiento', opcional: true },
  { k: 'moneda', n: 'Moneda', fijo: true },
  { k: 'monto', n: 'Monto' },
  { k: 'observaciones', n: 'Observaciones', opcional: true },
];

const PISTAS = {
  sociedad: ['sociedad', 'empresa', 'compania', 'origen', 'receptor'],
  codigo: ['codigo proveedor', 'cod proveedor', 'codigo', 'cod', 'id proveedor'],
  proveedor: ['nombre del proveedor', 'proveedor', 'nombre emisor', 'emisor', 'razon social', 'nombre'],
  documento: ['numero documento', 'numero de factura', 'factura', 'documento', 'consecutivo', 'numero', 'no factura', 'n factura'],
  fechaFactura: ['fecha factura', 'fec factura', 'fecha de factura', 'fecha emision', 'fecha documento', 'fecha'],
  fechaVence: ['fecha vence', 'fec vence', 'vencimiento', 'vence', 'fecha de vencimiento'],
  moneda: ['moneda', 'divisa', 'tipo moneda'],
  monto: ['monto total', 'total comprobante', 'total factura', 'monto', 'total', 'importe', 'valor'],
  observaciones: ['observaciones', 'detalle', 'descripcion', 'nota', 'comentario'],
};

const normal = (t) => String(t ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
  .replace(/[^a-z0-9]+/g, ' ').trim();

/** Lee el archivo (CSV o Excel) y devuelve las filas como listas de celdas. */
export function leerFilas(contenido, XLSX, esTexto) {
  const wb = esTexto
    ? XLSX.read(contenido.replace(/^﻿/, ''), { type: 'string', raw: true })
    : XLSX.read(contenido, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' })
    .filter((f) => f.some((c) => String(c).trim() !== ''));
}

/** Primera fila que parece de encabezados (al menos 3 textos). */
export function filaEncabezado(filas) {
  const i = filas.findIndex((f) => f.filter((c) => typeof c === 'string' && /[a-zA-Z]/.test(c)).length >= 3);
  return i < 0 ? 0 : i;
}

/** Propone qué columna corresponde a cada dato según el nombre del encabezado. */
export function sugerirMapeo(encabezados) {
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

const nombreParaComparar = (n) => normal(n).replace(/\b(s a|sa|s r l|srl|sociedad anonima|limitada|ltda)\b/g, '').replace(/\s+/g, ' ').trim();

/**
 * Convierte las filas en facturas listas para agregar.
 * @param filas      todas las filas del archivo
 * @param inicio     índice de la fila de encabezados
 * @param mapeo      { campo: índiceColumna }
 * @param fijos      { sociedad, moneda } cuando no vienen en una columna
 * @param datos      { facturas, proveedores, sociedades } de la aplicación
 * @returns { filas: [{n, estado, mensaje, factura, proveedor, proveedorNuevo}], resumen }
 */
export function prepararFacturas(filas, inicio, mapeo, fijos, datos) {
  const col = (fila, k) => (mapeo[k] == null || mapeo[k] === '' ? '' : fila[Number(mapeo[k])]);
  const socs = Object.values(datos.sociedades);
  const buscarSociedad = (v) => {
    const t = normal(v);
    if (!t) return null;
    return socs.find((s) => normal(s.clave) === t || normal(s.corto) === t || normal(s.clave).includes(t) || t.includes(normal(s.corto)))?.clave || null;
  };
  const provs = Object.values(datos.proveedores);
  const porNombre = new Map(provs.map((p) => [nombreParaComparar(p.nombre), p]));
  const existentes = new Set(Object.values(datos.facturas).map((f) => [f.sociedad, f.codigo, String(f.documento), f.moneda].join('|')));
  const enArchivo = new Set();
  let siguienteCodigo = Math.max(0, ...provs.map((p) => Number(p.codigo)).filter(Number.isFinite)) + 1;
  const nuevosProv = new Map();

  const salida = [];
  for (let i = inicio + 1; i < filas.length; i++) {
    const fila = filas[i];
    const r = { n: i + 1, estado: 'nueva', mensaje: '' };
    const sociedad = mapeo.sociedad != null && mapeo.sociedad !== '' ? buscarSociedad(col(fila, 'sociedad')) : fijos.sociedad;
    const mon = mapeo.moneda != null && mapeo.moneda !== '' ? moneda(col(fila, 'moneda')) : fijos.moneda;
    const documento = String(col(fila, 'documento') ?? '').trim();
    const nombreProv = String(col(fila, 'proveedor') ?? '').trim().replace(/\s{2,}/g, ' ');
    const codigoArchivo = String(col(fila, 'codigo') ?? '').trim();
    const fFactura = fecha(col(fila, 'fechaFactura'));
    const monto = round2(parseNum(col(fila, 'monto')));

    // Filas de totales o vacías
    if (!documento && !nombreProv) continue;
    if (/^total/i.test(documento) || /^total/i.test(nombreProv)) continue;

    const errores = [];
    if (!sociedad) errores.push('sociedad no reconocida');
    if (!mon) errores.push('moneda no reconocida');
    if (!documento) errores.push('falta el número de factura');
    if (!nombreProv && !codigoArchivo) errores.push('falta el proveedor');
    if (!fFactura) errores.push('fecha de factura inválida');
    if (!(monto > 0)) errores.push('monto inválido');

    // Proveedor: por código, luego por nombre; si no existe, se crea.
    let prov = codigoArchivo ? datos.proveedores[codigoArchivo] : null;
    if (!prov && nombreProv) prov = porNombre.get(nombreParaComparar(nombreProv)) || null;
    let proveedorNuevo = null;
    if (!prov && nombreProv) {
      const clave = nombreParaComparar(nombreProv);
      if (!nuevosProv.has(clave)) {
        const codigo = codigoArchivo && !datos.proveedores[codigoArchivo] ? codigoArchivo : String(siguienteCodigo++);
        nuevosProv.set(clave, { codigo, nombre: nombreProv.toUpperCase(), alias: '', ibanCRC: '', ibanUSD: '', plazo: 30 });
      }
      proveedorNuevo = nuevosProv.get(clave);
      prov = proveedorNuevo;
    }

    const fVence = fecha(col(fila, 'fechaVence')) || (fFactura ? sumarDias(fFactura, Number(prov?.plazo) || 30) : '');
    const factura = {
      sociedad, codigo: prov?.codigo || '', documento, moneda: mon, fechaFactura: fFactura, fechaVence: fVence,
      monto, observaciones: String(col(fila, 'observaciones') ?? '').trim(),
    };
    const clave = [sociedad, factura.codigo, documento, mon].join('|');
    if (errores.length) { r.estado = 'error'; r.mensaje = errores.join(', '); }
    else if (existentes.has(clave)) { r.estado = 'existe'; r.mensaje = 'ya está registrada'; }
    else if (enArchivo.has(clave)) { r.estado = 'existe'; r.mensaje = 'repetida en el archivo'; }
    else enArchivo.add(clave);
    Object.assign(r, { factura, proveedor: prov, proveedorNuevo, nombreProv });
    salida.push(r);
  }
  const resumen = {
    nuevas: salida.filter((x) => x.estado === 'nueva').length,
    existentes: salida.filter((x) => x.estado === 'existe').length,
    errores: salida.filter((x) => x.estado === 'error').length,
    proveedoresNuevos: [...nuevosProv.values()].filter((p) => salida.some((x) => x.estado === 'nueva' && x.proveedorNuevo === p)),
  };
  return { filas: salida, resumen };
}
