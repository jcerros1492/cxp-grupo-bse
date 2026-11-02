// Archivo para el sistema contable: el mismo formato del reporte de GTI que ya se sube hoy
// (13 columnas separadas por «;», sin encabezado, Windows-1252, fin de línea CRLF), ya limpio:
// decimales con punto y sin separador de miles, 0 en vez de «-», nombres sin espacios.

const TIPOS = { FE: 'FacturaElectrónica', TE: 'TiqueteElectrónico', NC: 'NotaCréditoElectrónica', ND: 'NotaDébitoElectrónica', FEC: 'FacturaElectrónicaCompra' };

/** «12294.4000» → «12294.4»; vacío o inválido → «0». */
export function numero(v) {
  const n = Number(String(v ?? '').replace(/,/g, '.'));
  if (!Number.isFinite(n) || n === 0) return '0';
  return String(Math.round(n * 100000) / 100000);
}
const fechaDMY = (iso) => { const [a, m, d] = String(iso || '').slice(0, 10).split('-'); return a ? `${Number(d)}/${Number(m)}/${a}` : ''; };
const sinEspacios = (t) => String(t || '').replace(/\s+/g, '').replace(/;/g, ',');

/** Una línea por comprobante recibido (fila de la hoja «Recibidas»). */
export function lineaContable(r) {
  const tc = Number(r.tipoCambio) || 1;
  const colones = (Math.round(Number(r.total || 0) * tc * 100) / 100).toFixed(2);
  return [TIPOS[r.tipo] || r.tipo, fechaDMY(r.fecha), r.consecutivo, r.cedulaReceptor, sinEspacios(r.nombreReceptor), r.cedulaEmisor,
    sinEspacios(r.nombreEmisor), numero(r.neto), numero(r.impuesto), numero(r.total), r.moneda || 'CRC', numero(r.tipoCambio || 1), colones].join(';');
}

/** Bytes en Windows-1252 (las letras con tilde como las escribe Excel); lo que no existe ahí queda «?». */
export function aWindows1252(texto) {
  const extra = { '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89, 'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e,
    '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f };
  const out = new Uint8Array(texto.length);
  for (let i = 0; i < texto.length; i++) {
    const c = texto.charCodeAt(i);
    out[i] = c < 256 && !(c >= 0x80 && c < 0xa0) ? c : (extra[texto[i]] ?? 63);
  }
  return out;
}

/** Contenido completo del archivo (bytes) para una lista de filas, ordenadas por fecha. */
export function archivoContable(filas) {
  const lineas = [...filas].sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)) || String(a.consecutivo).localeCompare(String(b.consecutivo))).map(lineaContable);
  return aWindows1252(lineas.map((l) => l + '\r\n').join(''));
}
