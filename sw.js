const CACHE = "cam-onedrive-v13";
const SHELL = [
  "./",
  "./index.html",
  "./app.js",
  "./config.js",
  "./manifest.json",
  "./msal-browser.min.js",
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

// login.microsoftonline.com y graph.microsoft.com siempre van directo a la red.
// index.html/app.js/config.js: red primero (para no quedar pegado con una
// versión vieja cuando se actualizan), con caché como respaldo sin conexión.
// Íconos: caché primero (no cambian).
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (url.origin.includes("microsoftonline.com") || url.origin.includes("graph.microsoft.com")) {
    return;
  }

  const isCoreFile = /\.(html)$|\/(app|config)\.js$|\/$/.test(url.pathname);

  if (isCoreFile) {
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          caches.open(CACHE).then((c) => c.put(e.request, res.clone()));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
  } else {
    // Estáticos (msal, ffmpeg-vendor, íconos): caché primero; si no está
    // en caché, se descarga de la red y se guarda para la próxima vez.
    e.respondWith(
      caches.match(e.request).then((cached) => {
        if (cached) return cached;
        return fetch(e.request).then((res) => {
          if (res && res.ok) {
            caches.open(CACHE).then((c) => c.put(e.request, res.clone()));
          }
          return res;
        });
      })
    );
  }
});
