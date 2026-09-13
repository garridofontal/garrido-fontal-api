// Ficheiro de Clientes — service worker
// Sube la versión cada vez que cambien los datos para forzar la actualización.
const VERSION = 'clientes-v1';
const ARCHIVOS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Red primero para el HTML (así ve los datos nuevos si hay cobertura),
// caché como respaldo inmediato cuando no hay conexión.
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const esPagina = req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html');
  if (esPagina) {
    e.respondWith(
      fetch(req)
        .then(r => { const copia = r.clone(); caches.open(VERSION).then(c => c.put(req, copia)); return r; })
        .catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
    );
  } else {
    e.respondWith(
      caches.match(req).then(r => r || fetch(req).then(resp => {
        if (resp.ok && new URL(req.url).origin === location.origin) {
          const copia = resp.clone(); caches.open(VERSION).then(c => c.put(req, copia));
        }
        return resp;
      }).catch(() => r))
    );
  }
});
