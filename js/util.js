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
