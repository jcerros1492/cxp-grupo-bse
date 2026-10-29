// Crea el archivo de Excel listo para subir a Google Drive y abrir como Google Sheets,
// con todas las pestañas que usa Codigo.gs, a partir del Excel de cuentas por pagar.
//
//   node tools/crear-hoja.mjs <excel-de-cuentas.xlsx> <salida.xlsx> [datos-extra.json]
//
// datos-extra.json (opcional) puede traer: { config: {clienteBanco, controlEncabezado},
//   sociedades: {"<nombre completo>": {cuentaCRC, cuentaUSD, descripcionBanco}},
//   proveedores: {"<codigo>": {alias, ibanCRC, ibanUSD}} }
// El resultado contiene datos reales: NO lo suba al repositorio.

import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { randomBytes } from 'node:crypto';
import vm from 'node:vm';

const require = createRequire(import.meta.url);
const raiz = new URL('..', import.meta.url).pathname;
const XLSX = require(raiz + 'js/vendor/xlsx.full.min.js');
const L = require(raiz + 'js/logica.js');
const { importarExcel } = await import(raiz + 'js/importer.js');

const [entrada, salida, extraRuta] = process.argv.slice(2);
if (!entrada || !salida) { console.error('Uso: node tools/crear-hoja.mjs <excel.xlsx> <salida.xlsx> [extra.json]'); process.exit(1); }
const extra = extraRuta ? JSON.parse(readFileSync(extraRuta, 'utf8')) : {};

// El esquema se lee del mismo Codigo.gs para que las columnas coincidan siempre.
const ctx = vm.createContext({});
vm.runInContext(readFileSync(raiz + 'apps-script/Codigo.gs', 'utf8') + '\n;this.__ = { ESQUEMA, COLS_USUARIOS, COLS_SESIONES, COLS_BITACORA };', ctx);
const { ESQUEMA, COLS_USUARIOS, COLS_SESIONES, COLS_BITACORA } = ctx.__;

const res = importarExcel(readFileSync(entrada), XLSX);
const ahora = Date.now();
const d = L.vacio();
for (const f of res.facturas) {
  const { movimientos, proveedorExcel, ...resto } = f;
  d.facturas[f.id] = { ...resto, proveedor: res.proveedores[f.codigo]?.nombre || proveedorExcel, creadaPor: 'Excel', v: ahora };
  for (const m of movimientos) d.movimientos[m.id] = { ...m, facturaId: f.id, programadoPor: '', loteId: '', v: ahora };
}
for (const p of Object.values(res.proveedores)) d.proveedores[p.codigo] = { ...p, ...(extra.proveedores?.[p.codigo] || {}), v: ahora };
for (const s of Object.values(res.sociedades)) {
  d.sociedades[s.nombre] = { clave: s.nombre, corto: s.corto, cuentaCRC: '', cuentaUSD: '', descripcionBanco: '', ...(extra.sociedades?.[s.nombre] || {}), v: ahora };
}
const cfg = { epoca: String(ahora), importado: new Date(ahora).toISOString().slice(0, 16).replace('T', ' ') + ' · Excel inicial', ...(extra.config || {}) };
for (const [k, v] of Object.entries(cfg)) d.config[k] = { clave: k, valor: String(v), v: ahora };
L.recalcularTodo(d);

const libro = XLSX.utils.book_new();
function pestaña(nombre, cols, filas, tipos) {
  const ws = XLSX.utils.aoa_to_sheet([cols, ...filas]);
  // Columnas de texto con formato "texto" para que Google no las convierta en fechas o números.
  const rango = XLSX.utils.decode_range(ws['!ref']);
  for (let r = 1; r <= rango.e.r; r++) {
    tipos.forEach((t, c) => {
      const cel = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cel) return;
      if (t === 't') { cel.t = 's'; cel.v = String(cel.v); cel.z = '@'; }
      if (t === 'n' && cols[c] !== 'v') cel.z = '#,##0.00';
    });
  }
  ws['!cols'] = cols.map((c) => ({ wch: Math.max(10, Math.min(40, c.length + 4)) }));
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  XLSX.utils.book_append_sheet(libro, ws, nombre);
}

for (const tabla of L.TABLAS) {
  const esq = ESQUEMA[tabla];
  const cols = esq.cols.map((c) => c[0]);
  const filas = Object.values(d[tabla]).map((e) => esq.cols.map(([k, t]) => (t === 'n' ? Number(e[k]) || 0 : t === 'b' ? !!e[k] : e[k] == null ? '' : String(e[k]))));
  pestaña(esq.hoja, cols, filas, esq.cols.map((c) => c[1]));
}
// Clave temporal «acceso»: la aplicación entra tocando el nombre y la reemplaza sola por una interna.
const temporal = () => 'acceso';
const usuarios = [
  ['jonathan', 'Jonathan Cerros', 'admin', 'TRUE', temporal(), '', ''],
  ['fernando', 'Fernando', 'aplica', 'TRUE', temporal(), '', ''],
  ['norman', 'Norman Alpízar', 'aprueba', 'TRUE', temporal(), '', ''],
  ['daniel', 'Daniel Barquero', 'aplica', 'TRUE', temporal(), '', ''],
];
pestaña('Usuarios', COLS_USUARIOS, usuarios, COLS_USUARIOS.map(() => 't'));
pestaña('Sesiones', COLS_SESIONES, [], COLS_SESIONES.map(() => 't'));
pestaña('Bitacora', COLS_BITACORA, [['importar-' + ahora, new Date(ahora).toISOString(), 'Excel inicial', 'importar', 'ok',
  `${res.facturas.length} facturas, ${Object.keys(d.movimientos).length} movimientos`]], COLS_BITACORA.map(() => 't'));

writeFileSync(salida, XLSX.write(libro, { type: 'buffer', bookType: 'xlsx', compression: true }));
const t = Object.values(d.facturas).reduce((a, f) => { if (f.saldo > 0.004) { a[f.moneda] += f.saldo; a.n++; } return a; }, { CRC: 0, USD: 0, n: 0 });
console.log(`Listo: ${salida}`);
console.log(`${res.facturas.length} facturas, ${Object.keys(d.movimientos).length} movimientos, ${Object.keys(d.proveedores).length} proveedores`);
console.log(`Con saldo: ${t.n} · CRC ${t.CRC.toFixed(2)} · USD ${t.USD.toFixed(2)}`);
if (res.avisos.length) console.log('Avisos:', res.avisos);
