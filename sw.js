/* 環島打卡 Service Worker — BUILD R3.0
   目的只有一個：沒訊號時打卡頁還打得開，打卡資料由頁面自己排隊補送。 */
const CACHE = 'digua-R3.0';
const SHELL = ['checkin.html', 'manifest.json', 'icon-192.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function(){ return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; })
                           .map(function (k) { return caches.delete(k); }));
  }).then(function(){ return self.clients.claim(); }));
});

self.addEventListener('fetch', function (e) {
  const url = new URL(e.request.url);
  // API 一律走網路，不快取
  if (e.request.method !== 'GET' || url.hostname === 'script.google.com') return;
  e.respondWith(
    fetch(e.request)
      .then(function (res) {
        const copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
        return res;
      })
      .catch(function () { return caches.match(e.request); })
  );
});
