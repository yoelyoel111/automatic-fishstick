// Service Worker for הגרלומט
const CACHE_NAME = 'hagralot-v1';
const urlsToCache = [
  '/automatic-fishstick/',
  '/automatic-fishstick/index.html',
  '/automatic-fishstick/styles/makor.css',
  '/automatic-fishstick/js/makor-core.js',
  '/automatic-fishstick/js/makor-init.js',
  '/automatic-fishstick/js/translations.js',
  '/automatic-fishstick/manifest.json'
];

// Install event
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        return cache.addAll(urlsToCache);
      })
  );
});

// Fetch event
self.addEventListener('fetch', function(event) {
  event.respondWith(
    caches.match(event.request)
      .then(function(response) {
        // Return cached version or fetch from network
        return response || fetch(event.request);
      }
    )
  );
});

// Activate event
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.map(function(cacheName) {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
});