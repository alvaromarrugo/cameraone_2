const CACHE = "cam-onedrive-v1";
const SHELL = [
  "./",
  "./index.html",
  "./app.js",
  "./config.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Solo cacheamos el shell estático. Las llamadas a login.microsoftonline.com
// y graph.microsoft.com siempre van directo a la red (nunca a caché).
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (url.origin.includes("microsoftonline.com") || url.origin.includes("graph.microsoft.com")) {
    return; // deja pasar sin interceptar
  }
  e.respondWith(
    caches.match(e.request).then((cached) => cached || fetch(e.request))
  );
});
