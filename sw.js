// Service worker del Life OS: solo cachea el "shell" de la app para que abra rapido
// y funcione aunque no haya internet. Nunca cachea llamadas a Microsoft Graph / login,
// para no interferir jamas con la sincronizacion de OneDrive.

const CACHE_NAME = 'life-os-shell-v1';
const SHELL_FILES = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
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

self.addEventListener('fetch', (event) => {
  const url = event.request.url;

  // Nunca tocar llamadas a Microsoft Graph, login de Microsoft, ni ninguna libreria externa (MSAL).
  if (
    url.includes('graph.microsoft.com') ||
    url.includes('login.microsoftonline.com') ||
    url.includes('microsoftonline.com') ||
    url.includes('cdn.jsdelivr.net') ||
    event.request.method !== 'GET'
  ) {
    return; // dejar pasar directo a la red, sin intervenir
  }

  // El HTML se sirve "network-first": asi, cuando se sube una actualizacion,
  // se ve enseguida en la proxima carga en vez de quedar pegada a una copia vieja
  // guardada en el cache del navegador. Si no hay internet, cae a la copia cacheada.
  const isShellHTML = event.request.mode === 'navigate' || url.endsWith('/') || url.endsWith('index.html');

  if (isShellHTML) {
    event.respondWith(
      fetch(event.request)
        .then((resp) => {
          if (resp && resp.ok) {
            const clone = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return resp;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((resp) => {
          if (resp && resp.ok) {
            const clone = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return resp;
        })
        .catch(() => cached);
    })
  );
});
