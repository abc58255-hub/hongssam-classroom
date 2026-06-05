importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging-compat.js');
importScripts('./config.js');

firebase.initializeApp(FIREBASE_CONFIG);

const messaging = firebase.messaging();

// 새 서비스워커 즉시 활성화 (옛 버전 캐시로 내용 비는 문제 방지)
self.addEventListener('install', function() { self.skipWaiting(); });
self.addEventListener('activate', function(e) { e.waitUntil(self.clients.claim()); });

// ⚠️ 중복 표시 방지: onBackgroundMessage에서 showNotification 하지 않음.
// notification 필드가 있으면 브라우저/iOS가 자동으로 1번 표시하므로 여기선 아무것도 안 함.
messaging.onBackgroundMessage(function() { /* no-op: 브라우저 자동표시에 맡김 */ });

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
