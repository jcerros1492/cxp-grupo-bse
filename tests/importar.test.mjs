// Importación de facturas desde CSV. Ejecutar con: node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { leerFilas, filaEncabezado, sugerirMapeo, prepararFacturas, detectarGTI, numeroDocumento } from '../js/importar-facturas.js';
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

// Reporte de GTI: CSV sin encabezado, separado por punto y coma (como el que se ajusta para el sistema contable)
const gtiCsv = [
  'Factura Electrónica;7/10/2026;00100001010000107091;3101006133;SOCIEDADAGRICOLACOMERCIALLAHILDAS.A.;3101010782;DistribuidoraAgroComercialS.A.;240;2.4;242.4;USD;457.8;110970.72',
  'Factura Electrónica;8/10/2026;00100001010000107099;3101006133;SOCIEDADAGRICOLACOMERCIALLAHILDAS.A.;3101010782;DistribuidoraAgroComercialS.A.;100;1;101;USD;457.8;46237.8',
  'Factura Electrónica;9/10/2026;00100001010000000701;3101006133;SOCIEDADAGRICOLACOMERCIALLAHILDAS.A.;602530318;CESAR ALBERTO CARRANZA GARITA;50;0;50;CRC;1;50',
  'Nota Crédito;9/10/2026;00100001030000000077;3101006133;SOCIEDADAGRICOLACOMERCIALLAHILDAS.A.;3101999999;TRANSPORTES MUFLA SOCIEDAD ANONIMA;-10;0;-10;CRC;1;-10',
].join('\r\n');
const datosGti = {
  sociedades: { 'SOC, AGRICOLA COMERCIAL LA HILDA S,A': { clave: 'SOC, AGRICOLA COMERCIAL LA HILDA S,A', corto: 'La Hilda' } },
  proveedores: {
    7: { codigo: '7', nombre: 'DISTRIBUIDORA AGRO COMERCIAL S.A.', alias: 'DAC', plazo: 30 },
    279: { codigo: '279', nombre: 'CARRANZA GARITA CESAR ALBERTO', plazo: 15 },
  },
  facturas: { a: { sociedad: 'SOC, AGRICOLA COMERCIAL LA HILDA S,A', codigo: '7', documento: '107091', moneda: 'USD', monto: 242.4 } },
};

test('número de factura desde el consecutivo electrónico', () => {
  assert.equal(numeroDocumento('00100001010000107091'), '107091');
  assert.equal(numeroDocumento('00200001010000046738'), '46738');
  assert.equal(numeroDocumento('FAC-12'), 'FAC-12');
});

test('reconoce el CSV de GTI sin encabezado y evita duplicados', () => {
  const g = detectarGTI(leerFilas(gtiCsv, XLSX, true));
  assert.ok(g);
  const r = prepararFacturas(g.filas, g.inicio, g.mapeo, {}, datosGti);
  const [a, b, c, d] = r.filas;
  assert.equal(a.estado, 'existe'); // 107091 ya está registrada
  assert.equal(b.estado, 'nueva');
  assert.equal(b.factura.codigo, '7');
  assert.equal(b.factura.sociedad, 'SOC, AGRICOLA COMERCIAL LA HILDA S,A');
  assert.equal(b.factura.monto, 101);
  assert.equal(b.factura.fechaFactura, '2026-10-08');
  assert.match(b.factura.observaciones, /Céd\. 3101010782/);
  assert.equal(c.factura.codigo, '279'); // mismo nombre con las palabras en otro orden
  assert.equal(c.factura.fechaVence, '2026-10-24'); // plazo del proveedor
  assert.equal(d.estado, 'omitida'); // nota de crédito
  assert.equal(r.resumen.nuevas, 2);
  assert.equal(r.resumen.proveedoresNuevos.length, 0);
});

test('la cédula aprendida y la elección del usuario mandan', () => {
  const g = detectarGTI(leerFilas(gtiCsv, XLSX, true));
  const r = prepararFacturas(g.filas, g.inicio, g.mapeo, {}, datosGti, { asignados: { 602530318: '' }, memoria: {} });
  assert.ok(r.filas[2].proveedorNuevo); // eligió "crear nuevo"
  const r2 = prepararFacturas(g.filas, g.inicio, g.mapeo, {}, datosGti, { asignados: { 3101010782: '279' } });
  assert.equal(r2.filas[1].factura.codigo, '279');
});

test('reconoce el Excel de GTI con encabezados', () => {
  const enc = ['IdEmpresa', 'TipoDocRecibido', 'ConsecutivoAceptación', 'FechaEmisiónAceptación', 'ConsecutivoDocumento', 'TipoDocAceptación', 'IdentificaciónReceptor', 'NombreReceptor', 'ActividadEconómicaReceptor', 'IdentificaciónEmisor', 'NombreEmisor', 'ActividadEconómicaEmisor', 'DetalleAceptación', 'FechaEmisiónDocumento', 'EstadoHacienda', 'MensajeHacienda', 'TotalDeGastoAplicable', 'TotalImpuestoAcreditar', 'TotalNeto', 'TotalImpuesto', 'TotalComprobante', 'CódigoMoneda'];
  const g = detectarGTI([enc]);
  assert.equal(g.mapeo.documento, 4);
  assert.equal(g.mapeo.fechaFactura, 13);
  assert.equal(g.mapeo.monto, 20);
  assert.equal(g.mapeo.moneda, 21);
  assert.equal(g.mapeo.cedulaEmisor, 9);
});
