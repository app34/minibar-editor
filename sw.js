const CACHE = "minibar-xl-3.0.0";

self.addEventListener("install", (e) => {
  e.waitUntil(self.skipWaiting());
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.map((k) => caches.delete(k)))).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  const live = /version\.json$|\/index\.html$|\/sw\.js$/.test(url.pathname) || url.searchParams.has("v");
  if (live) {
    e.respondWith(fetch(e.request, { cache: "no-store" }));
    return;
  }
  e.respondWith(
    fetch(e.request, { cache: "no-store" }).catch(() => caches.match(e.request))
  );
});
