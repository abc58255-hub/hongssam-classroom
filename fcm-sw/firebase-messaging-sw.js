importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging-compat.js');
importScripts('./config.js');

firebase.initializeApp(FIREBASE_CONFIG);

const messaging = firebase.messaging();

// 새 서비스워커 즉시 활성화
self.addEventListener('install', function() { self.skipWaiting(); });
self.addEventListener('activate', function(e) { e.waitUntil(self.clients.claim()); });

// 알림 표시: 서버가 data-only로 보내므로 표시는 여기서 딱 1번
// (notification 필드를 함께 보내면 브라우저 자동표시와 겹쳐 2번 뜸 — iOS는 tag 병합이 안 됨)
messaging.onBackgroundMessage(function(payload) {
  const d = payload.data || {};
  const base = self.location.origin + self.location.pathname.replace('firebase-messaging-sw.js', '');
  return self.registration.showNotification(d.title || '홍쌤 교실', {
    body: d.body || '',
    icon: d.icon || (base + 'icon-192.png'),
    tag: d.tag || 'hongssam',
    renotify: false,
    data: d
  });
});

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
