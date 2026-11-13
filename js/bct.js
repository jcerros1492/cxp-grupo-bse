// Archivo de pagos para el banco BCT (Enl@ce BCT · «Plantilla para ingreso de planillas o proveedores»).
// Una línea por proveedor, campos separados por coma, sin encabezado ni totales, Windows-1252, CRLF:
//   tipoId,cédula,nombre,cuentaCliente,monto,moneda,descripción,O,correo,concepto
// tipoId: 0 física nacional · 1 física residente/DIMEX · 2 gobierno · 3 jurídica · 4 institución autónoma.
// Solo se usa en Beneficiadora Santa Eduviges; el archivo del BCR (banco.js) no cambia.
import { aWindows1252 } from './contable.js';

/** Mayúsculas, sin comas (separan campos) y sin espacios de más. */
const campo = (t) => String(t ?? '').replace(/[,\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
/** Nombre como lo escribe la plantilla: sin tildes ni puntos («DUWEST CAFESA, S.A.» → «DUWEST CAFESA S A»). */
export const nombreBCT = (t) => campo(String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[.]+/g, ' '));

/** Cuenta cliente (17 dígitos) a partir del IBAN: CRkk0 + 17 dígitos. */
export function cuentaCliente(iban) {
  const s = String(iban || '').replace(/\s+/g, '').toUpperCase();
  return /^CR\d{20}$/.test(s) ? s.slice(5) : '';
}

/** IBAN a partir de la cuenta cliente de 17 dígitos (dígitos de control ISO 13616). */
export function ibanDeCuenta(cuenta) {
  const c = String(cuenta || '').replace(/\D/g, '');
  if (c.length !== 17) return '';
  const bban = '0' + c;
  let resto = 0;
  for (const ch of bban + '1227' + '00') resto = (resto * 10 + Number(ch)) % 97; // C=12 R=27
  return 'CR' + String(98 - resto).padStart(2, '0') + bban;
}

/** Cédula con guiones: 10 dígitos → 3-101-268981 · 9 dígitos → 1-0483-0067. */
export function cedulaBCT(ced) {
  const d = String(ced || '').replace(/\D/g, '');
  if (d.length === 10) return `${d[0]}-${d.slice(1, 4)}-${d.slice(4)}`;
  if (d.length === 9) return `${d[0]}-${d.slice(1, 5)}-${d.slice(5)}`;
  return d;
}

/** Tipo de identificación que se deduce de la cédula. */
export function tipoDe(ced) {
  const d = String(ced || '').replace(/\D/g, '');
  if (d.length === 9) return '0';
  if (d.length === 10) return ['2', '3', '4'].includes(d[0]) ? d[0] : '3';
  if (d.length === 11 || d.length === 12) return '1';
  return '';
}

/** Monto como en la plantilla: punto decimal, sin ceros de más (6624.29 · 2200 · 181.8). */
export const montoBCT = (n) => String(Math.round(Number(n || 0) * 100) / 100);

/** Problemas de una línea antes de generar el archivo. */
export function faltaBCT(x) {
  const out = [];
  if (!cuentaCliente(x.iban)) out.push('cuenta (IBAN)');
  if (!String(x.cedula || '').replace(/\D/g, '')) out.push('cédula');
  if (!/^\S+@\S+\.\S+$/.test(String(x.correo || '').trim())) out.push('correo');
  return out;
}

export function lineaBCT(x) {
  return [
    x.tipo || tipoDe(x.cedula) || '3', cedulaBCT(x.cedula), nombreBCT(x.nombre), cuentaCliente(x.iban), montoBCT(x.monto),
    x.moneda === 'USD' ? '2' : '1', campo(x.descripcion).toUpperCase(), 'O', campo(x.correo), campo(x.concepto).toUpperCase(),
  ].join(',');
}

/** Bytes del .txt (Windows-1252, CRLF), solo con las líneas de monto mayor a cero. */
export function archivoBCT(lineas) {
  const texto = lineas.filter((x) => Number(x.monto) > 0.004).map(lineaBCT).join('\r\n') + '\r\n';
  return { texto, bytes: aWindows1252(texto) };
}

/** Lee un .txt del BCT (el que genera la plantilla de Excel). */
export function leerArchivoBCT(texto) {
  const filas = [];
  for (const linea of String(texto || '').split(/\r?\n/)) {
    const c = linea.split(',');
    if (c.length < 10 || !/^\d$/.test(c[0].trim())) continue;
    filas.push({
      tipo: c[0].trim(), cedula: c[1].trim(), nombre: c[2].trim(), cuenta: c[3].replace(/\D/g, ''),
      monto: Number(c[4]) || 0, moneda: c[5].trim() === '2' ? 'USD' : 'CRC', descripcion: c[6].trim(),
      correo: c[8].trim(), concepto: c.slice(9).join(' ').trim(),
    });
  }
  return filas;
}
