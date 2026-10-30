// Generador del archivo .txt de pago a proveedores para el banco.
//
// Mismo formato que genera la plantilla de Excel del BCR («BCR Comercial»), confirmado con
// los archivos reales de las cuatro sociedades:
//  - Todas las líneas miden 214 caracteres, rellenadas con espacios, y terminan en CRLF.
//  - Encabezado:
//      1-15  «Cédula jurídica» de la plantilla (ej. 6133, 37251 o 3101032890), con ceros a la izquierda
//      16-18 consecutivo del archivo (el banco rechaza uno repetido)
//      19-26 fecha DDMMAAAA
//      27-38 cédula del empleado que sube el archivo, con ceros a la izquierda
//      39-50 espacios
//      51-56 000000
//  - Detalle (primera línea = débito a la cuenta de la sociedad; luego un crédito por proveedor):
//      1-3   000
//      4-25  IBAN
//      26-28 colones: 014 débito / 012 crédito · dólares: 024 débito / 022 crédito
//      29-32 espacios
//      33-52 monto en centavos (2 decimales, sin punto), 20 dígitos
//      53-84 espacios; en el débito, desde la posición 65 la «cédula jurídica» de la plantilla
//      85-214 descripción

import { ascii, fechaDDMMAAAA, limpiarIBAN, errorIBAN, round2 } from './util.js';

export const LARGO_LINEA = 214;
export const LARGO_DESCRIPCION = LARGO_LINEA - 84;
/** Códigos de movimiento: el segundo dígito es la moneda (1 colones, 2 dólares). */
const CODIGOS = { CRC: { debito: '014', credito: '012' }, USD: { debito: '024', credito: '022' } };

function padDer(texto, n) {
  const s = String(texto ?? '');
  return s.length >= n ? s.slice(0, n) : s + ' '.repeat(n - s.length);
}

function ceros(valor, n, campo) {
  const s = String(valor);
  if (!/^\d*$/.test(s)) throw new Error(`${campo}: solo puede tener dígitos (“${s}”)`);
  if (s.length > n) throw new Error(`${campo}: no cabe en ${n} dígitos (“${s}”)`);
  return s.padStart(n, '0');
}

export const centavos = (monto) => Math.round(round2(monto) * 100);

function lineaDetalle({ iban, codigo, monto, medio, descripcion }) {
  const err = errorIBAN(iban);
  if (err) throw new Error(`IBAN inválido ${iban || '(vacío)'}: ${err}`);
  const linea = '000' + limpiarIBAN(iban) + codigo + '    ' +
    ceros(centavos(monto), 20, 'Monto') + medio + ascii(descripcion).slice(0, LARGO_DESCRIPCION);
  return padDer(linea, LARGO_LINEA);
}

/**
 * Construye el texto del archivo.
 * @param {object} p
 * @param {string} p.cliente      «Cédula jurídica» de la plantilla del BCR (ej. "6133")
 * @param {number} p.consecutivo  consecutivo del archivo (1-999)
 * @param {string} p.fecha        fecha ISO "AAAA-MM-DD"
 * @param {string} p.control      cédula del empleado que sube el archivo (posiciones 27-38)
 * @param {'CRC'|'USD'} p.moneda
 * @param {{iban:string, descripcion:string}} p.debito  cuenta de la sociedad
 * @param {{iban:string, monto:number, descripcion:string}[]} p.creditos  un pago por proveedor
 * @returns {string} contenido con saltos CRLF
 */
export function construirArchivo({ cliente, consecutivo, fecha, control = '', moneda, debito, creditos }) {
  if (!creditos?.length) throw new Error('No hay pagos en el lote');
  for (const c of creditos) {
    if (!(c.monto > 0)) throw new Error(`Monto inválido para ${c.descripcion || c.iban}`);
  }
  const cod = CODIGOS[moneda];
  if (!cod) throw new Error('Moneda inválida');
  const total = creditos.reduce((s, c) => s + centavos(c.monto), 0) / 100;
  const clienteTxt = String(cliente ?? '').trim();
  if (!clienteTxt) throw new Error('Falta la cédula jurídica de la sociedad para el BCR (Configuración)');
  const n = Number(consecutivo);
  if (!(Number.isInteger(n) && n >= 1 && n <= 999)) throw new Error('El consecutivo del archivo debe ser un número de 1 a 999');

  const encabezado = padDer(
    ceros(clienteTxt, 15, 'Cédula jurídica') +
    ceros(n, 3, 'Consecutivo') +
    fechaDDMMAAAA(fecha) +
    ceros(String(control ?? '').trim() || '0', 12, 'Cédula del empleado') +
    ' '.repeat(12) + '000000',
    LARGO_LINEA,
  );

  const lineas = [
    encabezado,
    lineaDetalle({
      iban: debito.iban, codigo: cod.debito, monto: total,
      medio: ' '.repeat(12) + padDer(clienteTxt, 20),
      descripcion: debito.descripcion,
    }),
    ...creditos.map((c) => lineaDetalle({
      iban: c.iban, codigo: cod.credito, monto: c.monto, medio: ' '.repeat(32), descripcion: c.descripcion,
    })),
  ];
  return lineas.join('\r\n') + '\r\n';
}
