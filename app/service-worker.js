const CACHE_NAME = 'cafc-2026-agenda-v6';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './images/sponsors/rosenbauer.png',
  './images/sponsors/indocom.jpg',
  './images/sponsors/geomatica.jpg',
  './images/sponsors/esri.png',
  './images/sponsors/allen-co.gif',
  './images/sponsors/angloco.jpg',
  './images/sponsors/paladius.png',
  './images/sponsors/irp.png',
  './images/sponsors/holmatro.png',
  './images/sponsors/eagle.png',
  './audio/welcome-en.mp3',
  './audio/welcome-sp.mp3'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Network-first for HTML (so agenda updates appear immediately),
// cache-first for everything else (images/audio/manifest/icons).
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const isHTML = event.request.mode === 'navigate' ||
                 (event.request.headers.get('accept') || '').includes('text/html');
  if (isHTML) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => caches.match(event.request).then((cached) => cached || caches.match('./index.html')))
    );
    return;
  }
  event.respondWith(
    caches.match(event.request).then((cached) => {
      const fetchPromise = fetch(event.request)
        .then((response) => {
          if (response && response.status === 200 && response.type === 'basic') {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => cached);
      return cached || fetchPromise;
    })
  );
});
