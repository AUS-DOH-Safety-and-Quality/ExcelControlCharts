const cacheName = "excel-control-charts-v6";
const appShell = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./assets/spc_excel_add_in_logo_192_192.png",
  "./assets/spc_excel_add_in_logo_512_512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(cacheName).then((cache) => cache.addAll(appShell)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== cacheName).map((key) => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const network = fetch(event.request)
        .then((response) => {
          if (response.ok && new URL(event.request.url).origin === self.location.origin) {
            caches.open(cacheName).then((cache) => cache.put(event.request, response.clone()));
          }
          return response;
        })
        .catch(() => cached ?? caches.match("./index.html"));
      return cached ?? network;
    })
  );
});