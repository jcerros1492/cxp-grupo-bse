// Importación de facturas desde CSV. Ejecutar con: node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { leerFilas, filaEncabezado, sugerirMapeo, prepararFacturas } from '../js/importar-facturas.js';
const XLSX = createRequire(import.meta.url)('../js/vendor/xlsx.full.min.js');

const csv = [
  'Reporte de compras,,,,',
  'Fecha Emisión,Consecutivo,Nombre Emisor,Moneda,Total Comprobante',
  '29/10/2026,00100001010000012345,DISTRIBUIDORA AGRO COMERCIAL S.A.,USD,"1,234.50"',
  '30/10/2026,00100001010000012346,Proveedor Nuevo Ltda,CRC,"250.000,00"',
  '30/10/2026,999,DISTRIBUIDORA AGRO COMERCIAL S.A.,USD,10',
  '30/10/2026,999,DISTRIBUIDORA AGRO COMERCIAL S.A.,USD,10',
  ',,TOTAL,,1500',
].join('\r\n');

const datos = {
  sociedades: { 'NOSAVAR S,A,': { clave: 'NOSAVAR S,A,', corto: 'Nosavar' } },
  proveedores: { 7: { codigo: '7', nombre: 'DISTRIBUIDORA AGRO COMERCIAL S.A.', plazo: 30 } },
  facturas: { a: { sociedad: 'NOSAVAR S,A,', codigo: '7', documento: '999', moneda: 'USD' } },
};

test('reconoce columnas y prepara facturas', () => {
  const filas = leerFilas(csv, XLSX, true);
  const inicio = filaEncabezado(filas);
  assert.equal(inicio, 1);
  const mapeo = sugerirMapeo(filas[inicio]);
  assert.equal(mapeo.fechaFactura, 0);
  assert.equal(mapeo.documento, 1);
  assert.equal(mapeo.proveedor, 2);
  assert.equal(mapeo.moneda, 3);
  assert.equal(mapeo.monto, 4);
  const r = prepararFacturas(filas, inicio, mapeo, { sociedad: 'NOSAVAR S,A,' }, datos);
  const [a, b, c, d] = r.filas;
  assert.equal(a.estado, 'nueva');
  assert.equal(a.factura.codigo, '7');
  assert.equal(a.factura.monto, 1234.5);
  assert.equal(a.factura.fechaFactura, '2026-10-29');
  assert.equal(a.factura.fechaVence, '2026-11-28');
  assert.equal(b.estado, 'nueva');
  assert.equal(b.factura.monto, 250000);
  assert.equal(b.factura.moneda, 'CRC');
  assert.ok(b.proveedorNuevo);
  assert.equal(b.factura.codigo, '8');
  assert.equal(c.estado, 'existe');
  assert.equal(d.estado, 'existe');
  assert.equal(r.filas.length, 4); // la fila de TOTAL se ignora
  assert.equal(r.resumen.nuevas, 2);
  assert.equal(r.resumen.proveedoresNuevos.length, 1);
});
