// 수학교실 PWA 셸 서비스워커 — 네트워크 우선(항상 최신), 오프라인 시 캐시 폴백
// 앱 내용물은 GAS iframe에서 오므로 여기서는 셸(껍데기)만 다룬다
var CACHE = 'math-shell-v1';
var SHELL = ['./', './manifest.json', '../icon-192.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }));
  self.skipWaiting(); // 새 버전 즉시 적용 (수정 반영 지연 방지)
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  // 같은 폴더(셸 자원)만 처리 — GAS·파이어베이스 요청은 건드리지 않음
  if (e.request.method !== 'GET' || e.request.url.indexOf(self.registration.scope) !== 0) return;
  e.respondWith(
    fetch(e.request).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      return res;
    }).catch(function () { return caches.match(e.request); })
  );
});
