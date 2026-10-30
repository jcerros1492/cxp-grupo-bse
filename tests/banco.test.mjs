// Pruebas del generador del archivo del banco y de la importación.
// Ejecutar con:  node --test
// Usan cuentas ficticias (con dígitos de control válidos): no hay datos reales en el repositorio.

import test from 'node:test';
import assert from 'node:assert/strict';
import { construirArchivo, LARGO_LINEA } from '../js/banco.js';
import { errorIBAN, parseNum, fmtNum } from '../js/util.js';
import { importarFilas } from '../js/importer.js';

/** Crea un IBAN de Costa Rica ficticio con dígitos de control correctos. */
function ibanFicticio(cuerpo18) {
  const num = (cuerpo18 + 'CR00').replace(/[A-Z]/g, (c) => String(c.charCodeAt(0) - 55));
  let r = 0;
  for (const ch of num) r = (r * 10 + Number(ch)) % 97;
  return 'CR' + String(98 - r).padStart(2, '0') + cuerpo18;
}

const DEB = ibanFicticio('015200000000000001');
const P1 = ibanFicticio('015100000000000002');
const P2 = ibanFicticio('015200000000000003');

test('los IBAN ficticios son válidos y uno alterado no', () => {
  assert.equal(errorIBAN(DEB), null);
  assert.notEqual(errorIBAN(DEB.slice(0, -1) + (DEB.endsWith('1') ? '2' : '1')), null);
  assert.notEqual(errorIBAN('CR123'), null);
});

test('estructura del archivo', () => {
  const txt = construirArchivo({
    cliente: '1234', consecutivo: 5, moneda: 'USD', fecha: '2026-10-29', control: '205470116',
    debito: { iban: DEB, descripcion: 'PROVEE PRUEBA' },
    creditos: [
      { iban: P1, monto: 863.13, descripcion: 'Pago Fact 1 2 Uno' },
      { iban: P2, monto: 3529.92, descripcion: 'Pago Fact 3 mas Intereses Dos' },
    ],
  });
  assert.ok(txt.endsWith('\r\n'));
  const lineas = txt.split('\r\n').slice(0, -1);
  assert.equal(lineas.length, 4);
  for (const l of lineas) assert.equal(l.length, LARGO_LINEA);

  const [enc, deb, c1] = lineas;
  assert.equal(enc.slice(0, 15), '000000000001234');
  assert.equal(enc.slice(15, 18), '005'); // consecutivo del archivo
  assert.equal(enc.slice(18, 26), '29102026');
  assert.equal(enc.slice(26, 38), '000205470116');
  assert.equal(enc.slice(38, 50), ' '.repeat(12));
  assert.equal(enc.slice(50, 56), '000000');

  assert.equal(deb.slice(0, 28), '000' + DEB + '024');
  assert.equal(deb.slice(32, 52), '00000000000000439305'); // 863,13 + 3.529,92
  assert.equal(deb.slice(52, 84), ' '.repeat(12) + '1234'.padEnd(20));
  assert.equal(deb.slice(84).trimEnd(), 'PROVEE PRUEBA');

  assert.equal(c1.slice(0, 28), '000' + P1 + '022');
  assert.equal(c1.slice(32, 52), '00000000000000086313');
  assert.equal(c1.slice(52, 84), ' '.repeat(32));
  assert.equal(c1.slice(84).trimEnd(), 'Pago Fact 1 2 Uno');
});

test('quita tildes y rechaza IBAN inválido', () => {
  const txt = construirArchivo({
    cliente: '1', consecutivo: 1, moneda: 'CRC', fecha: '2026-01-02', debito: { iban: DEB, descripcion: 'Débito' },
    creditos: [{ iban: P1, monto: 1, descripcion: 'Pagó Ñandú' }],
  });
  assert.ok(txt.includes('Pago Nandu'));
  assert.throws(() => construirArchivo({
    cliente: '1', consecutivo: 1, moneda: 'CRC', fecha: '2026-01-02', debito: { iban: DEB, descripcion: 'x' },
    creditos: [{ iban: 'CR00123', monto: 1, descripcion: 'x' }],
  }), /IBAN inválido/);
});

test('colones usa 014/012, dólares 024/022; cédula larga completa en el débito', () => {
  const base = { consecutivo: 2, fecha: '2026-05-28', control: '205470116', debito: { iban: DEB, descripcion: 'LOMA - PROVEE' },
    creditos: [{ iban: P1, monto: 21089.5, descripcion: 'Pago Fact 1' }] };
  const crc = construirArchivo({ ...base, cliente: '3101032890', moneda: 'CRC' }).split('\r\n');
  assert.equal(crc[0].slice(0, 18), '000003101032890002');
  assert.equal(crc[1].slice(25, 28), '014');
  assert.equal(crc[1].slice(64, 84).trimEnd(), '3101032890');
  assert.equal(crc[2].slice(25, 28), '012');
  const usd = construirArchivo({ ...base, cliente: '37251', moneda: 'USD' }).split('\r\n');
  assert.equal(usd[1].slice(25, 28), '024');
  assert.equal(usd[2].slice(25, 28), '022');
  assert.throws(() => construirArchivo({ ...base, cliente: '1', moneda: 'CRC', consecutivo: 0 }), /consecutivo/);
  assert.throws(() => construirArchivo({ ...base, cliente: '1', moneda: 'CRC', consecutivo: 1000 }), /consecutivo/);
});

test('montos escritos a mano', () => {
  assert.equal(parseNum('1.234,56'), 1234.56);
  assert.equal(parseNum('1234.56'), 1234.56);
  assert.equal(parseNum('10000'), 10000);
  assert.equal(parseNum('1,234.56'), 1234.56);
  assert.equal(fmtNum(18598.65), '18.598,65');
});

test('importación: abonos negativos se unen a su factura aunque estén antes', () => {
  const enc = ['Origen de la Factura', 'Código', 'Nombre del Proveedor', 'Fec/Factura', 'Fec/Vence', 'Mes de vencimiento',
    'Número Documento', 'Colones', 'Dólares', 'Estado', 'Moneda', 'Fecha de Pago', 'N° Transferencia', 'Observaciones'];
  const f = (serial) => serial; // fechas como número de serie de Excel
  const filas = [enc,
    ['SOC A S,A', 89, 'PROV UNO S,A,', f(45203), f(45293), 'enero', 8715, -4000000, null, 'Cancelado', 'COL', f(45504), 111, null],
    ['SOC A S,A', 89, 'PROV UNO S,A,', f(45203), f(45293), 'enero', 8715, 10000000, null, 'Cancelado', 'COL', f(45576), 222, null],
    ['SOC A S,A', 7, 'PROV DOS', f(46200), null, null, 55, null, 100.5, 'Pendiente', 'USD', null, null, null],
    [null, null, null, null, null, null, 'TOTALES', 1, 1, null, null, null, null, null],
  ];
  const r = importarFilas(filas);
  assert.equal(r.facturas.length, 2);
  const a = r.facturas.find((x) => x.documento === '8715');
  assert.deepEqual(a.movimientos.map((m) => [m.tipo, m.monto]), [['abono', 4000000], ['pago', 6000000]]);
  assert.equal(r.proveedores['89'].nombre, 'PROV UNO S.A.');
  const b = r.facturas.find((x) => x.documento === '55');
  assert.equal(b.moneda, 'USD');
  assert.equal(b.venceEstimado, true);
  assert.equal(b.fechaVence, '2026-07-27'); // 2026-06-27 + 30 días
});
