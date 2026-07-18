// 수학교실 PWA 서비스워커 — 네트워크 우선(항상 최신), 오프라인 시 캐시 폴백
// 푸시 알림은 상위 경로의 firebase-messaging-sw.js가 전담 (여기서는 캐싱만)
var CACHE = 'math-app-v3';
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
  // 같은 폴더(앱 자원)만 처리 — GAS·파이어베이스 요청은 건드리지 않음
  if (e.request.method !== 'GET' || e.request.url.indexOf(self.registration.scope) !== 0) return;
  e.respondWith(
    // cache:'no-cache' = 브라우저 HTTP 캐시(Pages max-age=600)를 건너뛰고 서버에 재검증
    // → 수정 배포가 10분 캐시에 걸려 늦게 반영되던 문제 해결 (304면 응답 몇 바이트뿐)
    fetch(e.request, { cache: 'no-cache' }).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      return res;
    }).catch(function () { return caches.match(e.request); })
  );
});
