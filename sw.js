// Cosétika Service Worker — Push Notifications
const CACHE_NAME = 'cosetika-v1';

self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(clients.claim());
});

// Recibir notificación push
self.addEventListener('push', e => {
  if (!e.data) return;
  let data = {};
  try { data = e.data.json(); } catch(err) { data = { title: 'Cosétika', body: e.data.text() }; }

  const title = data.title || 'Cosétika — Nuevo pedido';
  const options = {
    body: data.body || '',
    icon: data.icon || '/icon-192.png',
    badge: '/badge-72.png',
    tag: data.tag || 'pedido-' + Date.now(),
    renotify: true,
    data: { url: data.url || '/' },
    vibrate: [200, 100, 200],
    actions: [
      { action: 'ver', title: 'Ver pedido' }
    ]
  };

  // Actualizar badge con número de pedidos pendientes
  if (data.badge !== undefined && navigator.setAppBadge) {
    navigator.setAppBadge(data.badge).catch(() => {});
  }

  e.waitUntil(self.registration.showNotification(title, options));
});

// Click en la notificación
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = e.notification.data?.url || '/';
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clientList => {
      // Si la app ya está abierta, enfocarla
      for (const client of clientList) {
        if (client.url.includes(self.location.origin) && 'focus' in client) {
          client.focus();
          client.postMessage({ type: 'NAVIGATE', url });
          return;
        }
      }
      // Si no está abierta, abrirla
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});

// Limpiar badge cuando se abre la app
self.addEventListener('message', e => {
  if (e.data?.type === 'CLEAR_BADGE' && navigator.clearAppBadge) {
    navigator.clearAppBadge().catch(() => {});
  }
});
