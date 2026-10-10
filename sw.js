const V = 'v24';
self.addEventListener('install', e => { self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x)))).then(() => self.clients.claim()));
});
// Always ask the network first (bypassing the browser's own cache) so updates show up immediately; fall back to the saved copy when offline.
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (u.pathname.includes('/img/')) {
    // Pictures rarely change: show the saved copy instantly and refresh it quietly in the background.
    e.respondWith(caches.open(V).then(ca => ca.match(e.request).then(m => {
      const f = fetch(e.request).then(r => { if (r.ok) ca.put(e.request, r.clone()); return r; }).catch(() => m);
      return m || f;
    })));
    return;
  }
  e.respondWith(fetch(e.request, { cache: 'no-store' }).then(r => { const c = r.clone(); caches.open(V).then(ca => ca.put(e.request, c)); return r; }).catch(() => caches.match(e.request)));
});
