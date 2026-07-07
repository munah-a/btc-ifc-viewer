/**
 * Service worker (PWA, W5.5 / C6 offline). Precaches the entire app shell —
 * HTML, hashed JS/CSS chunks, self-hosted web-ifc wasm + fragments worker,
 * fonts, icons — so the app boots and previously-cached models open with NO
 * network. C1: every precached URL is same-origin (no CDN); the SW never
 * fetches an external host.
 *
 * The precache list + version are injected at build time by
 * scripts/pwa-plugin.mjs so they always match the content-hashed asset names of
 * the current build.
 *
 * Strategy:
 *  - install: precache the shell, then skipWaiting.
 *  - activate: drop old caches, claim clients.
 *  - fetch (navigations): network-first with a cache fallback (so an online
 *    user gets fresh HTML, an offline user gets the cached shell).
 *  - fetch (same-origin assets): cache-first (hashed/immutable — serve instantly,
 *    fall back to network + cache on a miss).
 *  - cross-origin / non-GET: passthrough (never cached).
 */
const SW_VERSION = '__SW_VERSION__';
const CACHE_NAME = `btc-viewer-shell-${SW_VERSION}`;
const PRECACHE = __PRECACHE_MANIFEST__;

self.addEventListener('install', (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      // Add individually so one 404 doesn't abort the whole precache.
      await Promise.all(
        PRECACHE.map((url) =>
          cache.add(new Request(url, { cache: 'reload' })).catch(() => {
            /* ignore a single asset that fails to precache */
          }),
        ),
      );
      await self.skipWaiting();
    })(),
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(
        keys.filter((key) => key.startsWith('btc-viewer-shell-') && key !== CACHE_NAME).map((key) => caches.delete(key)),
      );
      await self.clients.claim();
    })(),
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return; // C1: never touch other hosts

  // Navigations: network-first (fresh when online), cached shell when offline.
  if (request.mode === 'navigate') {
    event.respondWith(
      (async () => {
        try {
          const response = await fetch(request);
          const cache = await caches.open(CACHE_NAME);
          cache.put(request, response.clone());
          return response;
        } catch {
          const cached = await caches.match(request);
          if (cached) return cached;
          // Fall back to the app shell entry (index.html) for any navigation.
          const shell = await caches.match('./index.html') || await caches.match('index.html');
          if (shell) return shell;
          throw new Error('offline and no cached shell');
        }
      })(),
    );
    return;
  }

  // Same-origin assets: cache-first (hashed assets are immutable). Match by URL
  // (ignoreVary/ignoreSearch) so a precached entry serves regardless of the
  // request's mode/headers — module/preload fetches vary and would otherwise
  // miss the cache and fail offline.
  event.respondWith(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(request, { ignoreVary: true, ignoreSearch: false });
      if (cached) return cached;
      const response = await fetch(request);
      if (response.ok) cache.put(request, response.clone());
      return response;
    })(),
  );
});
