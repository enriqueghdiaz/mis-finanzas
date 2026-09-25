// Service worker: guarda la app para que abra al instante y sin conexión.
// Los datos del Excel NO pasan por aquí (van directos a Microsoft Graph).
const CACHE = "mis-finanzas-v1";
const SHELL = ["./", "index.html", "style.css", "config.js", "manifest.webmanifest",
  "js/auth.js", "js/excel.js", "js/demo.js", "js/charts.js", "js/app.js",
  "icons/icon-192.png", "icons/icon-512.png", "icons/maskable-512.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  // Red primero (para recibir actualizaciones), caché si no hay conexión
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok && !url.search) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match("index.html"))));
});
