const CACHE = "mebel360-support-v5";
const ASSETS = [
  "/",
  "/static/css/style.css",
  "/static/js/app.js",
  "/static/img/mebel360-logo.jpg",
  "/static/img/icon-192.png",
  "/static/img/icon-512.png"
];
self.addEventListener("install", event => event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS))));
self.addEventListener("activate", event => event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))));
self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request)));
});
