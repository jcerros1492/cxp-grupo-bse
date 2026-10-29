// Simulación local de Google Sheets + Apps Script para probar la aplicación sin Google.
// Ejecuta el mismo Codigo.gs y Logica.gs que se pegan en Apps Script.
//
//   node tests/hoja-simulada.mjs [puerto] [hoja-inicial.xlsx]
//   POST /           → doPost (igual que la aplicación web de Apps Script)
//   GET  /claves     → usuarios y contraseñas temporales creados por configurarInicio() (solo pruebas)
//   GET  /hoja/<Pestaña> → contenido de una pestaña

import http from 'node:http';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { randomUUID, createHash } from 'node:crypto';

const raiz = new URL('..', import.meta.url).pathname;

class Rango {
  constructor(hoja, r, c, nr, nc) { Object.assign(this, { hoja, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const fila = this.hoja.datos[this.r - 1 + i] || [];
      const v = [];
      for (let j = 0; j < this.nc; j++) v.push(fila[this.c - 1 + j] ?? '');
      out.push(v);
    }
    return out;
  }
  setValues(vals) {
    if (vals.length !== this.nr || vals.some((f) => f.length !== this.nc)) throw new Error('Dimensiones incorrectas en setValues');
    vals.forEach((f, i) => {
      const fila = (this.hoja.datos[this.r - 1 + i] ||= []);
      f.forEach((v, j) => { fila[this.c - 1 + j] = v instanceof Date ? v : v; });
    });
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  clearContent() { this.setValues(Array.from({ length: this.nr }, () => Array(this.nc).fill(''))); return this; }
  setNumberFormat() { return this; }
  setNumberFormats() { return this; }
  setFontWeight() { return this; }
}

class Hoja {
  constructor(nombre) { this.nombre = nombre; this.datos = []; }
  getName() { return this.nombre; }
  getLastRow() {
    for (let i = this.datos.length - 1; i >= 0; i--) if ((this.datos[i] || []).some((v) => v !== '' && v != null)) return i + 1;
    return 0;
  }
  getLastColumn() { return Math.max(0, ...this.datos.map((f) => (f || []).length)); }
  getRange(r, c, nr = 1, nc = 1) { return new Rango(this, r, c, nr, nc); }
  setFrozenRows() {}
}

const libro = {
  hojas: [new Hoja('Hoja 1')],
  getSheetByName(n) { return this.hojas.find((h) => h.nombre === n) || null; },
  insertSheet(n) { const h = new Hoja(n); this.hojas.push(h); return h; },
  getSheets() { return this.hojas; },
  deleteSheet(h) { this.hojas = this.hojas.filter((x) => x !== h); },
  getUrl() { return 'https://docs.google.com/spreadsheets/d/SIMULADA/edit'; },
  getSpreadsheetTimeZone() { return 'America/Costa_Rica'; },
};

const propiedades = {};
const contexto = vm.createContext({
  SpreadsheetApp: { getActive: () => libro, flush() {} },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => propiedades[k] ?? null, setProperty: (k, v) => { propiedades[k] = v; } }) },
  ContentService: {
    MimeType: { JSON: 'json' },
    createTextOutput: (t) => ({ texto: t, setMimeType() { return this; } }),
  },
  Utilities: {
    getUuid: () => randomUUID(),
    DigestAlgorithm: { SHA_256: 'sha256' },
    Charset: { UTF_8: 'utf8' },
    computeDigest: (alg, texto) => [...createHash(alg).update(texto, 'utf8').digest()].map((b) => (b > 127 ? b - 256 : b)),
    formatDate: (d) => d.toISOString().slice(0, 10),
  },
  Logger: { log: (m) => console.log('[Apps Script]', m) },
  console,
});
vm.runInContext(readFileSync(raiz + 'js/logica.js', 'utf8'), contexto, { filename: 'Logica.gs' });
vm.runInContext(readFileSync(raiz + 'apps-script/Codigo.gs', 'utf8'), contexto, { filename: 'Codigo.gs' });
// Con un .xlsx como segundo argumento, la hoja simulada arranca con ese contenido
// (como si se hubiera subido a Drive y abierto con Google Sheets).
if (process.argv[3]) {
  const { createRequire } = await import('node:module');
  const XLSX = createRequire(import.meta.url)(raiz + 'js/vendor/xlsx.full.min.js');
  const wb = XLSX.read(readFileSync(process.argv[3]), { type: 'buffer' });
  libro.hojas = wb.SheetNames.map((n) => {
    const h = new Hoja(n);
    h.datos = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
    return h;
  });
}
vm.runInContext('configurarInicio()', contexto);

const puerto = Number(process.argv[2] || 8766);
let cola = Promise.resolve();

http.createServer((req, res) => {
  const cors = { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' };
  if (req.method === 'GET' && req.url === '/servicio') {
    res.writeHead(200, cors).end(JSON.stringify(propiedades.CLAVE_SERVICIO));
    return;
  }
  if (req.method === 'GET' && req.url === '/claves') {
    const filas = libro.getSheetByName('Usuarios').datos.slice(1);
    res.writeHead(200, cors).end(JSON.stringify(filas));
    return;
  }
  if (req.method === 'GET' && req.url.startsWith('/hoja/')) {
    const h = libro.getSheetByName(decodeURIComponent(req.url.slice(6)));
    res.writeHead(h ? 200 : 404, cors).end(JSON.stringify(h ? h.datos : null));
    return;
  }
  if (req.method !== 'POST') { res.writeHead(200, cors).end(JSON.stringify({ ok: true })); return; }
  let cuerpo = '';
  req.on('data', (c) => { cuerpo += c; });
  req.on('end', () => {
    // Las solicitudes se atienden de una en una, como el bloqueo de Apps Script.
    cola = cola.then(async () => {
      await new Promise((r) => setTimeout(r, 150));
      const salida = contexto.doPost({ postData: { contents: cuerpo } });
      res.writeHead(200, cors).end(salida.texto);
    });
  });
}).listen(puerto, () => console.log('Hoja simulada en http://localhost:' + puerto));
