const CACHE_NAME = "realstock-offline-v6";

const URLS_TO_CACHE = [
  "/manifest.json",
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

  if (request.method !== "GET") {
    return;
  }

  const url = new URL(request.url);

  if (url.pathname === "/coleta-mobile") {
    event.respondWith(
      fetch(request).catch(() => caches.match("/coleta-mobile-offline"))
    );
    return;
  }

  if (
    url.pathname === "/manifest.json" ||
    url.pathname === "/icone-192.png" ||
    url.pathname === "/icone-512.png"
  ) {
    event.respondWith(
      caches.match(request).then((cached) => cached || fetch(request))
    );
    return;
  }

  event.respondWith(fetch(request));
});