// Cofre Transcrição: activa o isolamento de origem (necessário para usar vários núcleos do processador)
// e guarda os ficheiros da aplicação para funcionar sem ligação. Não guarda áudio nem transcrições.
const CACHE = 'cofre-audio-v3';
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', (e) => {
  const req = e.request;
  const url = new URL(req.url);
  if (url.origin !== location.origin || req.method !== 'GET') return;
  if (req.cache === 'only-if-cached' && req.mode !== 'same-origin') return;
  if (url.pathname.endsWith('revogadas.json')) return;
  e.respondWith((async () => {
    let r;
    try {
      r = await fetch(req);
      if (r.ok && !/\.(m4a|mp3|wav)$/.test(url.pathname)) { const c = r.clone(); caches.open(CACHE).then((ca) => ca.put(req, c)); }
    } catch (err) {
      r = await caches.match(req);
      if (!r) throw err;
    }
    if (r.status === 0) return r;
    const h = new Headers(r.headers);
    h.set('Cross-Origin-Opener-Policy', 'same-origin');
    h.set('Cross-Origin-Embedder-Policy', 'require-corp');
    h.set('Cross-Origin-Resource-Policy', 'same-origin');
    return new Response(r.body, { status: r.status, statusText: r.statusText, headers: h });
  })());
});
