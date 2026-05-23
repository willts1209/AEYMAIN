// Minimal service worker — just enough to make the PWA installable on iOS.
// No caching for the MVP; all requests go through to the network so the user
// always gets fresh code while iterating.
self.addEventListener('install', (e) => self.skipWaiting());
self.addEventListener('activate', (e) => self.clients.claim());
self.addEventListener('fetch', (e) => {});
