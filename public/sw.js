const CACHE_NAME = 'soulwinning-pwa-cache-v1';

// 1. Pre-cache your core SPA file so the fallback actually exists
const PRECACHE_ASSETS = [
  '/',
  '/index.html'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(clients.claim());
});

self.addEventListener('fetch', (event) => {
  // Don't cache non-GET requests
  if (event.request.method !== 'GET') return;
  
  const url = new URL(event.request.url);
  // Don't cache external API calls (e.g., Supabase)
  if (url.origin !== location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Only cache valid, successful responses
        if (!response || response.status !== 200 || response.type !== 'basic') {
          return response;
        }

        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return response;
      })
      .catch(() => {
        return caches.match(event.request).then((cachedResponse) => {
          // Return the cached version if we have it
          if (cachedResponse) return cachedResponse;
          
          // Fallback to index.html for React Router / SPA navigation
          if (event.request.mode === 'navigate') {
            return caches.match('/index.html').then((htmlResponse) => {
              // Ensure we actually return a Response, even if index.html is missing
              return htmlResponse || new Response('Offline', { status: 503, statusText: 'Offline' });
            });
          }

          // FINAL FALLBACK: For images/scripts/css when offline to prevent the TypeError
          return new Response('', { status: 503, statusText: 'Offline' });
        });
      })
  );
});