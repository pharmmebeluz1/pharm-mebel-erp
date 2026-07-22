const CACHE = 'mebel360-pwa-v1.0.0';
const CORE = [
  '/', '/offline.html', '/manifest.webmanifest',
  '/static/app.css?v=1.0.0', '/static/app.js?v=1.0.0',
  '/static/assets/mebel360-logo.png', '/static/assets/icon-192.png',
  '/static/assets/icon-512.png', '/static/assets/favicon.png'
];
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(CORE)));
});
self.addEventListener('activate', event => {
  event.waitUntil(Promise.all([
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))),
    self.clients.claim()
  ]));
});
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  if (event.request.mode === 'navigate') {
    event.respondWith(fetch(event.request).then(response => {
      const copy = response.clone(); caches.open(CACHE).then(c => c.put(event.request, copy)); return response;
    }).catch(() => caches.match(event.request).then(r => r || caches.match('/offline.html'))));
    return;
  }
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
    if (response.ok) { const copy=response.clone(); caches.open(CACHE).then(c => c.put(event.request,copy)); }
    return response;
  })));
});
