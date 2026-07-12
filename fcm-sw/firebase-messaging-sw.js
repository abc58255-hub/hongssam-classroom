importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging-compat.js');
importScripts('./config.js');

firebase.initializeApp(FIREBASE_CONFIG);

const messaging = firebase.messaging();

// 새 서비스워커 즉시 활성화
self.addEventListener('install', function() { self.skipWaiting(); });
self.addEventListener('activate', function(e) { e.waitUntil(self.clients.claim()); });

// 알림 표시는 FCM SDK가 notification 페이로드로 자동 1회 수행 — 여기서 또 표시하면 2번 뜸.
// (data-only + 수동 표시 방식은 iOS가 서비스워커를 늦게 깨우면 빈 알림이 떠서 폐기)
// 클릭 동작도 SDK가 fcm_options.link로 처리하고, 아래 핸들러는 구형 알림용 폴백.

// 알림 클릭 시 앱 열기
self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  var url = (event.notification.data && event.notification.data.url) || '/';

  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(list) {
      for (var i = 0; i < list.length; i++) {
        if (list[i].url.indexOf('script.google.com') > -1 && 'focus' in list[i]) {
          return list[i].focus();
        }
      }
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});
