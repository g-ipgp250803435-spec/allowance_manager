const CACHE_NAME = 'allowance-manager-v1';
const ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icon.svg',
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap',
  'https://cdn.tailwindcss.com',
  'https://unpkg.com/lucide@latest',
  'https://cdn.jsdelivr.net/npm/chart.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    })
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);

  // For API dashboard requests, handle specifically to cache data
  if (url.pathname.includes('/api') && url.searchParams.get('action') === 'getDashboard') {
    const cacheKey = new Request('/api?action=getDashboard-cached', { method: 'GET' });
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          if (res.status === 200) {
            const resClone = res.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(cacheKey, resClone);
            });
          }
          return res;
        })
        .catch(() => {
          return caches.match(cacheKey);
        })
    );
  } else {
    // Standard assets
    e.respondWith(
      caches.match(e.request).then((cachedRes) => {
        return cachedRes || fetch(e.request);
      })
    );
  }
});
