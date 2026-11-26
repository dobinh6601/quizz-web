// Cho phép mở Quizz khi không có mạng. Ưu tiên bản mới trên mạng, mất mạng thì dùng bản đã lưu.
const CACHE = "quizz-web-1795681204";
const FILES = ["./", "index.html", "importer.js", "manifest.webmanifest", "icons/icon-192.png", "icons/icon-512.png",
               "icons/apple-touch-icon.png"];
// Bộ đọc PDF (pdf.js) tải từ cdnjs khi nhập PDF lần đầu; lưu lại để lần sau đọc được cả khi mất mạng
const PDFJS = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/";
const PDFJS_CACHE = "quizz-pdfjs";

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE && k !== PDFJS_CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  if (e.request.method === "GET" && e.request.url.startsWith(PDFJS)) {
    e.respondWith(caches.open(PDFJS_CACHE).then(c => c.match(e.request).then(hit => hit || fetch(e.request).then(res => {
      if (res.ok) c.put(e.request, res.clone());
      return res;
    }))));
    return;
  }
  if (e.request.method !== "GET" || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return res;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match("index.html"))));
});
