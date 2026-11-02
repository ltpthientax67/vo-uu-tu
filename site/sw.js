/* Vô Ưu Tự · service worker: trang luôn lấy bản mới khi có mạng; âm thanh, thư viện 3D, phông chữ lưu sẵn để mở nhanh. */
const CACHE = "vut-1.3";
const CORE = ["./", "index.html", "config.js", "version.json", "manifest.webmanifest", "icon-192.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).catch(() => {})); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request; if (r.method !== "GET") return;
  const u = new URL(r.url);
  if (u.hostname.endsWith("script.google.com") || u.hostname.endsWith("googleusercontent.com")) return;
  const heavy = u.pathname.includes("/sounds/") || u.hostname === "cdnjs.cloudflare.com" || u.hostname.endsWith("gstatic.com") || u.hostname === "fonts.googleapis.com" || /\.png$/.test(u.pathname);
  if (heavy) {
    e.respondWith(caches.open(CACHE).then(c => c.match(r).then(hit => hit || fetch(r).then(res => { if (res.ok || res.type === "opaque") c.put(r, res.clone()); return res; }))));
    return;
  }
  if (u.origin === location.origin) {
    e.respondWith(fetch(r).then(res => { if (res.ok) caches.open(CACHE).then(c => c.put(r, res.clone())); return res; }).catch(() => caches.match(r, { ignoreSearch: true }).then(h => h || caches.match("index.html"))));
  }
});
