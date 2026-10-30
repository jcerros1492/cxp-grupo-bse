// Guarda la aplicación en la computadora para que abra sin internet.
// Estrategia: primero la red (para recibir actualizaciones) y, si no hay, la copia guardada.
// Las consultas a Google Sheets (POST) no pasan por aquí.

const CACHE = 'cxp-1.18.2';
const ARCHIVOS = [
  './', 'index.html', 'css/app.css', 'manifest.webmanifest',
  'js/app.js', 'js/store.js', 'js/sync.js', 'js/config.js', 'js/logica.js', 'js/util.js',
  'js/banco.js', 'js/importer.js', 'js/importar-facturas.js', 'js/vendor/xlsx.full.min.js', 'js/vendor/jspdf.umd.min.js', 'js/vendor/jspdf.plugin.autotable.min.js', 'js/reporte-pdf.js', 'js/aviso-pdf.js', 'js/vendor/pdf-lib.min.js', 'img/emblema-oscuro.png',
  'img/icono-180.png', 'img/icono-192.png', 'img/icono-512.png', 'img/icono-maskable-512.png',
  'img/emblema-blanco.png', 'img/logo.png', 'img/favicon-64.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const propio = url.origin === self.location.origin;
  const fuentes = url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com';
  if (!propio && !fuentes) return;
  e.respondWith(
    fetch(req, { cache: 'no-cache' })
      .then((resp) => {
        if (resp.ok || resp.type === 'opaque') {
          const copia = resp.clone();
          caches.open(CACHE).then((c) => c.put(req, copia));
        }
        return resp;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || (req.mode === 'navigate' ? caches.match('index.html') : undefined))),
  );
});
