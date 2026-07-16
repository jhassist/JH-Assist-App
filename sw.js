const CACHE_NAME = 'jh-assist-1.1.0-v3';
const APP_SHELL = [
  './',
  './index.html',
  './app.js',
  './sync.js',
  './vendor/msal-browser.min.js',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const requestUrl = new URL(event.request.url);

  // Microsoft-Anmeldung, Graph-Aufrufe und die kurzlebigen OneDrive-
  // Downloadadressen dürfen weder abgefangen noch im App-Cache gespeichert
  // werden. Der Service Worker verwaltet ausschließlich eigene App-Dateien.
  if (requestUrl.origin !== self.location.origin) return;

  if (event.request.mode === 'navigate' || requestUrl.pathname.endsWith('/index.html')) {
    event.respondWith(
      fetch(event.request, { cache: 'no-store' })
        .then(response => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put('./index.html', copy));
          return response;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
      if (response.ok && response.type === 'basic') {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
      }
      return response;
    }))
  );
});
