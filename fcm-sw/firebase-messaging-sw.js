importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging-compat.js');
importScripts('./config.js');

firebase.initializeApp(FIREBASE_CONFIG);

const messaging = firebase.messaging();

// 앱이 닫혀 있을 때 백그라운드 알림 처리
messaging.onBackgroundMessage(function(payload) {
  const title   = (payload.notification && payload.notification.title) || '홍쌤 교실';
  const options = {
    body:  (payload.notification && payload.notification.body) || '',
    icon:  'https://abc58255-hub.github.io/fcm-sw/icon-192.png',
    badge: 'https://abc58255-hub.github.io/fcm-sw/badge.png',
    tag:   (payload.data && payload.data.tag) || 'hongssam',
    data:  payload.data || {}
  };
  return self.registration.showNotification(title, options);
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
