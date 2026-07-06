/* Wes's Castle service worker — NETWORK-FIRST so a new index.html always
   wins; the cache only serves when offline. Cross-origin requests (the
   Apps Script backend) pass through untouched. */
const CACHE = 'wes-castle-v1';

self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    }).catch(() => caches.match(e.request))
  );
});
