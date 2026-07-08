// Cosétika Service Worker — Push Notifications
self.addEventListener('install', e => { self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(clients.claim()); });

self.addEventListener('push', e => {
  if (!e.data) return;
  let data = {};
  try { data = e.data.json(); } catch(err) { data = { title: 'Cosétika', body: e.data.text() }; }
  const options = {
    body: data.body || '',
    icon: '/icon-192.png',
    badge: '/badge-72.png',
    tag: data.tag || 'pedido-' + Date.now(),
    renotify: true,
    data: { url: data.url || '/' },
    vibrate: [200, 100, 200]
  };
  if (data.badge !== undefined && self.navigator?.setAppBadge) {
    self.navigator.setAppBadge(data.badge).catch(() => {});
  }
  e.waitUntil(self.registration.showNotification(data.title || 'Cosétika', options));
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clientList => {
      for (const client of clientList) {
        if (client.url.includes(self.location.origin) && 'focus' in client) {
          client.focus();
          // Decirle a la app que vaya a Pedidos
          client.postMessage({ type: 'GO_PEDIDOS' });
          return;
        }
      }
      if (clients.openWindow) return clients.openWindow('/');
    })
  );
});

self.addEventListener('message', e => {
  if (e.data?.type === 'CLEAR_BADGE' && self.navigator?.clearAppBadge) {
    self.navigator.clearAppBadge().catch(() => {});
  }
});
