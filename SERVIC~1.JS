/* Habit Tracker service worker — caches the app shell for offline use.
   Bump CACHE when you change any shell file so clients update. */
const CACHE = 'habits-v1';
const SHELL = [
  './', './index.html', './app.js', './config.js',
  './manifest.webmanifest', './icon-192.png', './icon-512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(()=>{}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Never cache Microsoft Graph / login — always hit network.
  if (url.hostname.endsWith('microsoft.com') || url.hostname.endsWith('microsoftonline.com') ||
      url.hostname.endsWith('sharepoint.com') || url.hostname.endsWith('msauth.net')) {
    return; // default network handling
  }
  // App shell: cache-first, fall back to network.
  if (url.origin === location.origin) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
  }
});
