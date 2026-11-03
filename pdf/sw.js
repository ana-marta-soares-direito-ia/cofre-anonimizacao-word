// Cofre PDF: guarda os ficheiros da aplicação para funcionar sem ligação. Não guarda documentos.
const CACHE = 'cofre-pdf-v3';
const FILES = ['./', 'index.html', 'app.css', 'app.js', 'lib/pdf.min.mjs', 'lib/pdf.worker.min.mjs', 'lib/pdf-lib.min.js',
  '../engine.js', '../docxutil.js', '../licenca.js', '../nacl-fast.min.js', 'icones/icone-192.png',
  '../fonts/fraunces-600.woff2', '../fonts/inter-400.woff2', '../fonts/inter-500.woff2', '../fonts/inter-600.woff2'];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  if (url.pathname.endsWith('revogadas.json')) return; // sempre da rede
  // rede primeiro, para receber actualizações; sem ligação, usa a cópia guardada
  e.respondWith(fetch(e.request).then((r) => { const c = r.clone(); caches.open(CACHE).then((ca) => ca.put(e.request, c)); return r; }).catch(() => caches.match(e.request)));
});
