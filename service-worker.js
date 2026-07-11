const CACHE_NAME = "realstock-mobile-v8";

const URLS_TO_CACHE = [
  "/",
  "/index.html",
  "/contagem-mobile.html",
  "/manifest.json",
  "/service-worker.js",

  "/logo-realstock.png",
  "/icone-192.png",
  "/icone-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(URLS_TO_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const request = event.request;

  if (request.method !== "GET") return;

  const url = new URL(request.url);

  if (url.pathname === "/coleta-mobile") {
    event.respondWith(
      fetch(request).catch(() => caches.match("/coleta-mobile"))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then((cached) => {
      return cached || fetch(request);
    })
  );
});