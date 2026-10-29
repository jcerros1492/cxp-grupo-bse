// Importa el Excel "CUENTAS POR PAGAR" y lo convierte al modelo de la aplicación.
//
// Reglas acordadas:
//  - Filas con monto negativo son abonos o notas de crédito de la factura con la misma
//    sociedad + código de proveedor + número de documento + moneda.
//  - Si no hay fecha de vencimiento, se estima como fecha de factura + 30 días.
//  - Facturas "Cancelado" sin fecha de pago se dejan así (antes no se registraba).

import { excelSerialToISO, parseFecha, sumarDias, round2, uid } from './util.js';

const COLUMNAS = {
  sociedad: ['origen de la factura'],
  codigo: ['código', 'codigo'],
  proveedor: ['nombre del proveedor'],
  fechaFactura: ['fec/factura', 'fecha factura'],
  fechaVence: ['fec/vence', 'fecha vence'],
  documento: ['número documento', 'numero documento', 'n° documento'],
  colones: ['colones'],
  dolares: ['dólares', 'dolares'],
  estado: ['estado'],
  moneda: ['moneda'],
  fechaPago: ['fecha de pago'],
  transferencia: ['n° transferencia', 'no transferencia', 'nº transferencia'],
  observaciones: ['observaciones'],
};

export const NOMBRES_CORTOS = {
  'BENEFICIADORA SANTA EDUVIGES S,A': 'Santa Eduviges',
  'SOC, AGRICOLA COMERCIAL LA HILDA S,A': 'La Hilda',
  'NOSAVAR S,A,': 'Nosavar',
  'HACIENDA LOMA BONITA S,A,': 'Loma Bonita',
};

function valorFecha(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v > 20000 && v < 80000 ? excelSerialToISO(v) : null;
  if (v instanceof Date) return excelSerialToISO(v.getTime() / 86400000 + 25569);
  return parseFecha(v);
}

function texto(v) {
  if (v == null) return '';
  if (typeof v === 'number') return String(Math.round(v) === v ? v : v);
  return String(v).trim();
}

/** "EL COLONO AGROPECUARIO S,A," -> "EL COLONO AGROPECUARIO S.A." */
function limpiarNombre(n) {
  return n.replace(/\s{2,}/g, ' ').replace(/\bS,\s?R,\s?L,?/g, 'S.R.L.').replace(/\bS,\s?A,?/g, 'S.A.').replace(/,\s*$/, '').trim();
}

function numero(v) {
  if (v == null || v === '') return null;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

/** Recibe un ArrayBuffer del .xlsx y la librería XLSX. Devuelve { facturas, proveedores, sociedades, avisos }. */
export function importarExcel(buffer, XLSX) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  return importarFilas(filas);
}

/** Igual que importarExcel pero a partir de las filas (primera fila = encabezados). */
export function importarFilas(filas) {
  const enc = (filas[0] || []).map((h) => String(h ?? '').trim().toLowerCase());
  const idx = {};
  for (const [campo, nombres] of Object.entries(COLUMNAS)) {
    idx[campo] = enc.findIndex((h) => nombres.includes(h));
  }
  const faltan = ['sociedad', 'codigo', 'proveedor', 'documento', 'estado'].filter((c) => idx[c] < 0);
  if (faltan.length) throw new Error('No se encontraron estas columnas en el Excel: ' + faltan.join(', '));

  const get = (fila, campo) => (idx[campo] >= 0 ? fila[idx[campo]] : null);
  const avisos = [];
  const facturas = [];
  const porClave = new Map();
  const nombresPorCodigo = new Map();
  const sociedades = new Set();
  const negativos = [];

  filas.slice(1).forEach((fila, i) => {
    const nFila = i + 2;
    const sociedad = texto(get(fila, 'sociedad'));
    if (!sociedad) return; // filas vacías o de TOTALES
    const estadoTxt = texto(get(fila, 'estado')).toLowerCase();
    let moneda = texto(get(fila, 'moneda')).toUpperCase() === 'USD' ? 'USD' : 'CRC';
    const col = numero(get(fila, 'colones'));
    const usd = numero(get(fila, 'dolares'));
    let monto = moneda === 'USD' ? usd : col;
    if (monto == null) {
      if (usd != null) { monto = usd; moneda = 'USD'; } else if (col != null) { monto = col; moneda = 'CRC'; }
    }
    if (monto == null) { avisos.push(`Fila ${nFila}: sin monto, no se importó`); return; }

    sociedades.add(sociedad);
    const codigo = texto(get(fila, 'codigo')) || 'SIN-CODIGO';
    const nombre = limpiarNombre(texto(get(fila, 'proveedor')));
    const cuenta = nombresPorCodigo.get(codigo) || new Map();
    cuenta.set(nombre, (cuenta.get(nombre) || 0) + 1);
    nombresPorCodigo.set(codigo, cuenta);

    const documento = texto(get(fila, 'documento'));
    const fechaFactura = valorFecha(get(fila, 'fechaFactura'));
    const fechaPagoRaw = get(fila, 'fechaPago');
    const fechaPago = valorFecha(fechaPagoRaw);
    const transferencia = texto(get(fila, 'transferencia'));
    let observaciones = texto(get(fila, 'observaciones'));
    if (fechaPagoRaw != null && !fechaPago) observaciones = [observaciones, texto(fechaPagoRaw)].filter(Boolean).join(' · ');
    if (transferencia && !/^\d+$/.test(transferencia)) observaciones = [observaciones, transferencia].filter(Boolean).join(' · ');
    const comprobante = /^\d+$/.test(transferencia) ? transferencia : '';
    const clave = [sociedad, codigo, documento, moneda].join('|');

    let fechaVence = valorFecha(get(fila, 'fechaVence'));
    let venceEstimado = false;
    if (!fechaVence && fechaFactura) { fechaVence = sumarDias(fechaFactura, 30); venceEstimado = true; }

    const f = {
      id: uid(), sociedad, codigo, proveedorExcel: nombre, documento, fechaFactura, fechaVence, venceEstimado,
      moneda, monto: round2(monto), revision: estadoTxt.startsWith('revis'), canceladaExcel: estadoTxt.startsWith('cancel'),
      observaciones, movimientos: [], pagoExcel: { fecha: fechaPago, comprobante },
    };

    if (monto < 0) {
      // En el Excel los abonos a veces están antes que la factura: se unen al final.
      negativos.push({ clave, nFila, f, mov: { id: uid(), tipo: 'abono', monto: round2(-monto), fecha: fechaPago, comprobante, por: 'Excel', nota: observaciones } });
      return;
    }
    facturas.push(f);
    if (!porClave.has(clave)) porClave.set(clave, []);
    porClave.get(clave).push(f);
  });

  for (const n of negativos) {
    const destino = porClave.get(n.clave)?.[0];
    if (destino) { destino.movimientos.push(n.mov); continue; }
    avisos.push(`Fila ${n.nFila}: monto negativo sin factura (doc. ${n.f.documento}); se importó como nota de crédito suelta`);
    facturas.push(n.f);
  }

  // Las canceladas en el Excel: el saldo que quede se registra como pago histórico.
  for (const f of facturas) {
    const abonado = f.movimientos.reduce((s, m) => s + m.monto, 0);
    const resto = round2(f.monto - abonado);
    if (f.canceladaExcel && resto > 0) {
      f.movimientos.push({ id: uid(), tipo: 'pago', monto: resto, fecha: f.pagoExcel.fecha, comprobante: f.pagoExcel.comprobante, por: 'Excel', nota: '' });
    }
    f.movimientos.sort((a, b) => String(a.fecha || '').localeCompare(String(b.fecha || '')));
    delete f.pagoExcel;
    delete f.canceladaExcel;
  }

  const proveedores = {};
  for (const [codigo, nombres] of nombresPorCodigo) {
    const masUsado = [...nombres.entries()].sort((a, b) => b[1] - a[1])[0][0];
    proveedores[codigo] = { codigo, nombre: masUsado, alias: '', ibanCRC: '', ibanUSD: '', plazo: 30 };
  }

  const socs = {};
  for (const s of sociedades) {
    socs[s] = { nombre: s, corto: NOMBRES_CORTOS[s] || s.replace(/,/g, '.').replace(/\s+S\.A\.?\.?$/i, ''), cuentaCRC: '', cuentaUSD: '', descripcionBanco: '' };
  }
  return { facturas, proveedores, sociedades: socs, avisos };
}
