// Reglas de negocio compartidas con Apps Script. Ejecutar con: node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
const require = createRequire(import.meta.url);
const L = require('../js/logica.js');

function base() {
  const d = L.vacio();
  d.facturas.f1 = { id: 'f1', sociedad: 'S', codigo: '1', documento: '100', moneda: 'USD', monto: 1000, revision: false };
  d.proveedores['1'] = { codigo: '1', nombre: 'Prov' };
  return L.prepararIndices(d);
}
const ctx = (nombre) => ({ usuario: { nombre, rol: 'aplica' }, ahora: 1 });

test('programar un abono y aplicarlo deja saldo', () => {
  const d = base();
  assert.ok(L.aplicarOp(d, { tipo: 'programar', fechaPago: '2026-11-06', items: [{ progId: 'p1', facturaId: 'f1', monto: 400 }] }, ctx('Norman')).ok);
  assert.equal(L.estado(d, d.facturas.f1), 'Programada');
  const r = L.aplicarOp(d, { tipo: 'aplicarLote', loteId: 'l1', progIds: ['p1'], comprobante: '123', fecha: '2026-11-06' }, ctx('Jonathan'));
  assert.ok(r.ok);
  assert.equal(L.saldo(d, d.facturas.f1), 600);
  assert.equal(d.facturas.f1.estado, 'Abonada');
  assert.equal(d.movimientos['l1-p1'].tipo, 'abono');
});

test('no se puede programar dos veces ni aplicar dos veces', () => {
  const d = base();
  L.aplicarOp(d, { tipo: 'programar', fechaPago: '2026-11-06', items: [{ progId: 'p1', facturaId: 'f1', monto: 1000 }] }, ctx('Norman'));
  const dos = L.aplicarOp(d, { tipo: 'programar', fechaPago: '2026-11-07', items: [{ progId: 'p2', facturaId: 'f1', monto: 10 }] }, ctx('Norman'));
  assert.equal(dos.ok, false);
  assert.ok(L.aplicarOp(d, { tipo: 'aplicarLote', loteId: 'a', progIds: ['p1'], comprobante: '1', fecha: '2026-11-06' }, ctx('Jonathan')).ok);
  const otra = L.aplicarOp(d, { tipo: 'aplicarLote', loteId: 'b', progIds: ['p1'], comprobante: '2', fecha: '2026-11-06' }, ctx('Daniel'));
  assert.equal(otra.ok, false);
  assert.match(otra.error, /ya había sido aplicada por Jonathan/);
  assert.equal(L.saldo(d, d.facturas.f1), 0);
  assert.equal(d.facturas.f1.estado, 'Cancelada');
});

test('monto mayor que el saldo se rechaza sin cambiar nada', () => {
  const d = base();
  const r = L.aplicarOp(d, { tipo: 'movimiento', movId: 'm', facturaId: 'f1', tipo: 'abono', monto: 2000, fecha: '2026-01-01' }, ctx('J'));
  assert.equal(r.ok, false);
  assert.equal(Object.keys(d.movimientos).length, 0);
});

test('la copia para Apps Script es idéntica', () => {
  assert.equal(readFileSync(new URL('../apps-script/Logica.gs', import.meta.url), 'utf8'), readFileSync(new URL('../js/logica.js', import.meta.url), 'utf8'));
});
