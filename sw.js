const CACHE_NAME = 'leave-system-v1';
const ASSETS = [
  '/leave-system/',
  '/leave-system/index.html',
  '/leave-system/manifest.json',
  '/leave-system/icon-192x192.png',
  '/leave-system/icon-512x512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // API 請求（Supabase）永遠走網路，不快取
  if (e.request.url.includes('supabase.co')) {
    e.respondWith(fetch(e.request));
    return;
  }
  // 其他資源：網路優先，失敗才用快取
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const clone = res.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
