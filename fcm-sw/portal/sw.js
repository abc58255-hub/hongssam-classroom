// 교사포털 PWA 서비스워커 — 네트워크 우선(항상 최신), 오프라인 시 캐시 폴백
// 푸시 알림은 상위 경로의 firebase-messaging-sw.js가 전담 (여기서는 캐싱만)
var CACHE = 'portal-app-v1';
var SHELL = ['./', './manifest.json'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }));
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET' || e.request.url.indexOf(self.registration.scope) !== 0) return;
  e.respondWith(
    // cache:'no-cache' = Pages의 10분 HTTP 캐시를 건너뛰고 서버에 재검증 (수정 즉시 반영)
    fetch(e.request, { cache: 'no-cache' }).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      return res;
    }).catch(function () { return caches.match(e.request); })
  );
});
