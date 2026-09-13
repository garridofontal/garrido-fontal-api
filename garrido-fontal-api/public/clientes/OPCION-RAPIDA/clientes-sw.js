// Service worker do Ficheiro de Clientes — só actúa sobre /clientes*
const VERSION = 'clientes-v1';
const PAXINA = 'clientes.html';
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.add(PAXINA)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // non tocar o resto da aplicación (xerador de facturas, API...)
  if (e.request.method !== 'GET' || !url.pathname.includes('clientes')) return;
  e.respondWith(
    fetch(e.request)
      .then(r => { const c = r.clone(); caches.open(VERSION).then(x => x.put(e.request, c)); return r; })
      .catch(() => caches.match(e.request).then(r => r || caches.match(PAXINA)))
  );
});
