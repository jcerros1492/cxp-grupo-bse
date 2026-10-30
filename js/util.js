// Utilidades generales: fechas, montos, IBAN y HTML.

export const MONEDAS = { CRC: { simbolo: '₡', nombre: 'Colones' }, USD: { simbolo: '$', nombre: 'Dólares' } };

const nf = new Intl.NumberFormat('es-CR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/** Formatea un monto con su símbolo: ₡1.234,50 / $1.234,50 */
export function money(monto, moneda) {
  const s = MONEDAS[moneda]?.simbolo ?? '';
  const n = Number(monto) || 0;
  return (n < 0 ? '-' : '') + s + fmtNum(Math.abs(n));
}

/** 1234.5 -> "1.234,50" (separadores de Costa Rica) */
export function fmtNum(n) {
  // Intl en algunos navegadores usa espacio como separador de miles para es-CR; se normaliza a punto.
  return nf.format(round2(n)).replace(/\s/g, '.').replace(/ /g, '.');
}

/**
 * Monto escrito en un campo de la aplicación: puntos de miles y coma para céntimos ("1.000.000,50").
 * A diferencia de parseNum, "1.000" es mil (no 1).
 */
export function parseMonto(texto) {
  if (typeof texto === 'number') return texto;
  const s = String(texto ?? '').replace(/[^\d,]/g, '').replace(',', '.').replace(/,/g, '');
  return s && s !== '.' ? Number(s) : NaN;
}

/** Da formato mientras se escribe: "1000000,5" → "1.000.000,5" (máximo 2 céntimos). */
export function formatearMontoEscrito(texto) {
  let s = String(texto ?? '').replace(/[^\d,]/g, '');
  const i = s.indexOf(',');
  let ent = i < 0 ? s : s.slice(0, i);
  const dec = i < 0 ? null : s.slice(i + 1).replace(/,/g, '').slice(0, 2);
  ent = ent.replace(/^0+(?=\d)/, '');
  if (dec !== null && !ent) ent = '0';
  ent = ent.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return dec === null ? ent : ent + ',' + dec;
}

/** Lee un número escrito por una persona: "1.234,56", "1234.56", "1,234.56", "$ 500" */
export function parseNum(texto) {
  if (typeof texto === 'number') return texto;
  let s = String(texto ?? '').replace(/[^\d.,-]/g, '');
  if (!s) return NaN;
  const lastComma = s.lastIndexOf(','), lastDot = s.lastIndexOf('.');
  if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.');
  else if (lastDot > lastComma && lastComma !== -1) s = s.replace(/,/g, '');
  else if (lastComma === -1 && (s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
  return Number(s);
}

export const round2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;

// ---------- Fechas (se guardan como texto ISO "AAAA-MM-DD") ----------

export function hoyISO() {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
const pad2 = (n) => String(n).padStart(2, '0');

/** "2026-10-29" -> "29/10/2026" */
export function fmtFecha(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

/** "29/10/2026" o "2026-10-29" -> ISO, o null si no es válida */
export function parseFecha(texto) {
  const s = String(texto ?? '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return validISO(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
  if (m) return validISO(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  return null;
}
function validISO(y, m, d) {
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  return `${y}-${pad2(m)}-${pad2(d)}`;
}

export function sumarDias(iso, dias) {
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + dias));
  return `${dt.getUTCFullYear()}-${pad2(dt.getUTCMonth() + 1)}-${pad2(dt.getUTCDate())}`;
}

/** Días desde `iso` hasta hoy (positivo = vencida hace N días) */
export function diasDesde(iso, hoy = hoyISO()) {
  const a = Date.UTC(...iso.split('-').map((v, i) => (i === 1 ? v - 1 : +v)));
  const b = Date.UTC(...hoy.split('-').map((v, i) => (i === 1 ? v - 1 : +v)));
  return Math.round((b - a) / 86400000);
}

/** Número de serie de Excel -> ISO */
export function excelSerialToISO(serial) {
  const ms = Math.round((serial - 25569) * 86400000);
  const dt = new Date(ms);
  return `${dt.getUTCFullYear()}-${pad2(dt.getUTCMonth() + 1)}-${pad2(dt.getUTCDate())}`;
}

/** "29102026" para el archivo del banco */
export function fechaDDMMAAAA(iso) {
  const [y, m, d] = iso.split('-');
  return `${d}${m}${y}`;
}

// ---------- IBAN (Costa Rica: CR + 20 dígitos = 22 caracteres) ----------

export function limpiarIBAN(texto) {
  return String(texto ?? '').replace(/\s+/g, '').toUpperCase();
}

/** Valida formato y dígitos de control (mod 97). Devuelve null si es válido o el motivo del error. */
export function errorIBAN(texto) {
  const iban = limpiarIBAN(texto);
  if (!iban) return 'Vacío';
  if (!/^CR\d{20}$/.test(iban)) return 'Debe ser CR seguido de 20 dígitos (22 caracteres)';
  const reord = iban.slice(4) + iban.slice(0, 4);
  const num = reord.replace(/[A-Z]/g, (c) => String(c.charCodeAt(0) - 55));
  let resto = 0;
  for (const ch of num) resto = (resto * 10 + Number(ch)) % 97;
  return resto === 1 ? null : 'Los dígitos de control no coinciden: revise el número';
}

export function fmtIBAN(texto) {
  return limpiarIBAN(texto).replace(/(.{4})/g, '$1 ').trim();
}

// ---------- HTML ----------

export function esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

export const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);

/** Quita tildes y caracteres no ASCII (para el archivo del banco) */
export function ascii(texto) {
  return String(texto ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\x20-\x7E]/g, '');
}

/** Hora corta para "Sincronizado 10:42" (o fecha si no es de hoy) */
export function fechaHora(ms) {
  const d = new Date(ms), hoy = new Date();
  const hora = `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  return d.toDateString() === hoy.toDateString() ? hora : `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)} ${hora}`;
}
