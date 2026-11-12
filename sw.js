// Lưu sẵn các file của trang để dùng được khi mất mạng.
// Khi sửa trang, tăng số phiên bản ở đây VÀ số ?v= trong thẻ <script> của index.html.
const CACHE = "so-dang-bo-v6";
const FILES = ["./", "index.html", "core.js", "word.js", "exceljs.min.js", "jszip.min.js", "xlsx.full.min.js", "manifest.webmanifest", "icon.svg"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Có mạng: lấy bản mới nhất và cập nhật bộ nhớ. Mất mạng: dùng bản đã lưu.
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET" || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    fetch(e.request)
      .then(res => {
        if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
