const CACHE_NAME = "aryan-export-v3";
const FILES = [
  "./",
  "./index.html",
  "./app.js",
  "./manifest.json"
];

self.addEventListener("install", evt => {
  evt.waitUntil(
    caches.open(CACHE_NAME).then(c => c.addAll(FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", evt => {
  evt.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", evt => {
  evt.respondWith(
    fetch(evt.request).then(res => {
      const clone = res.clone();
      caches.open(CACHE_NAME).then(c => c.put(evt.request, clone));
      return res;
    }).catch(() => caches.match(evt.request))
  );
});
