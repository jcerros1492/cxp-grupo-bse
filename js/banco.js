// Generador del archivo .txt de pago a proveedores para el banco.
//
// Formato deducido del archivo real que se sube hoy (ver README):
//  - Todas las líneas miden 214 caracteres, rellenadas con espacios, y terminan en CRLF.
//  - Encabezado:
//      1-15  código de cliente, con ceros a la izquierda (ej. 000000000001234)
//      16-18 cantidad de líneas de detalle (débito + créditos)
//      19-26 fecha DDMMAAAA
//      27-38 campo de 12 dígitos aún SIN CONFIRMAR (se toma de Configuración)
//      39-50 espacios
//      51-56 000000
//  - Detalle (primera línea = débito a la cuenta de la sociedad; luego un crédito por proveedor):
//      1-3   000
//      4-25  IBAN
//      26-28 024 = débito / 022 = crédito
//      29-32 espacios
//      33-52 monto en centavos (2 decimales, sin punto), 20 dígitos
//      53-84 espacios; en el débito, posiciones 65-68 = código de cliente
//      85-214 descripción

import { ascii, fechaDDMMAAAA, limpiarIBAN, errorIBAN, round2 } from './util.js';

export const LARGO_LINEA = 214;
export const LARGO_DESCRIPCION = LARGO_LINEA - 84;
const COD_DEBITO = '024';
const COD_CREDITO = '022';

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
 * @param {string} p.cliente   código de cliente en el banco (ej. "1234")
 * @param {string} p.fecha     fecha ISO "AAAA-MM-DD"
 * @param {string} [p.control] campo de 12 dígitos del encabezado (posiciones 27-38)
 * @param {{iban:string, descripcion:string}} p.debito  cuenta de la sociedad
 * @param {{iban:string, monto:number, descripcion:string}[]} p.creditos  un pago por proveedor
 * @returns {string} contenido con saltos CRLF
 */
export function construirArchivo({ cliente, fecha, control = '', debito, creditos }) {
  if (!creditos?.length) throw new Error('No hay pagos en el lote');
  for (const c of creditos) {
    if (!(c.monto > 0)) throw new Error(`Monto inválido para ${c.descripcion || c.iban}`);
  }
  const total = creditos.reduce((s, c) => s + centavos(c.monto), 0) / 100;
  const clienteTxt = String(cliente ?? '').trim();
  if (!clienteTxt) throw new Error('Falta el código de cliente del banco (Configuración)');

  const encabezado = padDer(
    ceros(clienteTxt, 15, 'Código de cliente') +
    ceros(creditos.length + 1, 3, 'Cantidad de líneas') +
    fechaDDMMAAAA(fecha) +
    ceros(String(control ?? '').trim() || '0', 12, 'Campo de control') +
    ' '.repeat(12) + '000000',
    LARGO_LINEA,
  );

  const lineas = [
    encabezado,
    lineaDetalle({
      iban: debito.iban, codigo: COD_DEBITO, monto: total,
      medio: ' '.repeat(12) + padDer(clienteTxt, 4) + ' '.repeat(16),
      descripcion: debito.descripcion,
    }),
    ...creditos.map((c) => lineaDetalle({
      iban: c.iban, codigo: COD_CREDITO, monto: c.monto, medio: ' '.repeat(32), descripcion: c.descripcion,
    })),
  ];
  return lineas.join('\r\n') + '\r\n';
}
