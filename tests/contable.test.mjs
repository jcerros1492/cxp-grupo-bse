// Archivo para el sistema contable: mismo formato del reporte de GTI. Datos ficticios.
import test from 'node:test';
import assert from 'node:assert/strict';
import { lineaContable, archivoContable, numero } from '../js/contable.js';

test('números como en GTI: punto decimal, sin ceros de más, 0 en vez de vacío', () => {
  assert.equal(numero('12294.4000'), '12294.4');
  assert.equal(numero('240.00000'), '240');
  assert.equal(numero(''), '0');
  assert.equal(numero('0.00000'), '0');
  assert.equal(numero('457.80'), '457.8');
});

test('línea con las 13 columnas, nombres sin espacios y colones con 2 decimales', () => {
  const l = lineaContable({ tipo: 'FE', fecha: '2026-10-07', consecutivo: '00100001010000000001', cedulaReceptor: '3101000001',
    nombreReceptor: 'EMPRESA DE PRUEBA S.A.', cedulaEmisor: '3101000002', nombreEmisor: 'Proveedor Ficticio S.A.',
    neto: '240.00000', impuesto: '2.40000', total: '242.40000', moneda: 'USD', tipoCambio: '457.80' });
  assert.equal(l, 'FacturaElectrónica;7/10/2026;00100001010000000001;3101000001;EMPRESADEPRUEBAS.A.;3101000002;ProveedorFicticioS.A.;240;2.4;242.4;USD;457.8;110970.72');
  assert.equal(l.split(';').length, 13);
});

test('archivo en Windows-1252 con CRLF', () => {
  const b = archivoContable([{ tipo: 'FE', fecha: '2026-11-02', consecutivo: '1', cedulaReceptor: '1', nombreReceptor: 'A', cedulaEmisor: '2', nombreEmisor: 'CAÑA', neto: '1', impuesto: '0', total: '1', moneda: 'CRC', tipoCambio: '1' }]);
  const t = Buffer.from(b).toString('latin1');
  assert.ok(t.startsWith('Facturaó'.replace('Facturaó', 'FacturaElectrónica')));
  assert.ok(t.includes('CAÑA'));
  assert.ok(t.endsWith('\r\n'));
  assert.equal(b[13], 0xf3); // «ó» de Electrónica en un solo byte
});
