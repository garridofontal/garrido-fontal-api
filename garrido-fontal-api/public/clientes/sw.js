// Service worker do Ficheiro de Clientes — Garrido Fontal SLU
const VERSION = 'clientes-v4';
const ARCHIVOS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './favicon.png'
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

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const esHTML = req.mode === 'navigate' ||
    (req.headers.get('accept') || '').includes('text/html');

  if (esHTML) {
    // rede primeiro: sempre os datos máis recentes cando hai conexión
    e.respondWith(
      fetch(req)
        .then(r => {
          const copia = r.clone();
          caches.open(VERSION).then(c => c.put(req, copia));
          return r;
        })
        .catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
    );
    return;
  }

  // o resto: caché primeiro
  e.respondWith(
    caches.match(req).then(r => r || fetch(req).then(resp => {
      const copia = resp.clone();
      caches.open(VERSION).then(c => c.put(req, copia));
      return resp;
    }).catch(() => r))
  );
});
