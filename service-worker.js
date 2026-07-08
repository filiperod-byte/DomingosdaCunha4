const DC4_CACHE = 'dc4-pwa-v3-20260708';
const APP_SHELL = [
  '/DomingosdaCunha4/',
  '/DomingosdaCunha4/V2/index.html',
  '/DomingosdaCunha4/V2/app.css',
  '/DomingosdaCunha4/V2/app.js',
  '/DomingosdaCunha4/V2/config.json',
  '/DomingosdaCunha4/V2/extintores-direct.js',
  '/DomingosdaCunha4/V2/admin.html',
  '/DomingosdaCunha4/assets/icons/app-icon.svg',
  '/DomingosdaCunha4/manifest.webmanifest',
  '/DomingosdaCunha4/pwa-install.js'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(caches.open(DC4_CACHE).then(cache => cache.addAll(APP_SHELL).catch(() => undefined)));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== DC4_CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;

  if (url.pathname.includes('/macros/') || url.searchParams.has('action')) return;

  event.respondWith(
    caches.match(req).then(cached => {
      const network = fetch(req).then(res => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(DC4_CACHE).then(cache => cache.put(req, copy));
        }
        return res;
      }).catch(() => cached || caches.match('/DomingosdaCunha4/V2/index.html'));
      return cached || network;
    })
  );
});
