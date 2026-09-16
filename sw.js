const CACHE_NAME = 'dmf-pwa-v1';

self.addEventListener('install', event => {
    self.skipWaiting();
});

self.addEventListener('activate', event => {
    event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event => {
    // Modo seguro: Siempre busca en internet para que las entradas del reloj nunca fallen ni se queden guardadas en caché viejo.
    event.respondWith(fetch(event.request));
});
