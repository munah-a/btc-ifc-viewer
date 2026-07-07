/**
 * Service-worker registration (PWA, W5.5 / C6). Called from bundled app code
 * (not an inline script — CSP script-src 'self' blocks those). Registers the
 * build-time-generated `/sw.js` (see scripts/pwa-plugin.mjs) at the app scope so
 * the shell + self-hosted wasm/worker/fonts precache and the app works offline.
 *
 * No-ops when the SW isn't available (unsupported browser, non-secure context,
 * or a dev/e2e build where sw.js was not emitted — the plugin only runs on
 * `vite build`, and a 404 is swallowed). Registration failures are non-fatal.
 */
export function registerServiceWorker(): void {
  if (typeof navigator === 'undefined' || !('serviceWorker' in navigator)) return;
  // Resolve against BASE_URL so the Pages mirror (base=/btc-ifc-viewer/) works.
  const swUrl = `${import.meta.env.BASE_URL}sw.js`;
  const register = (): void => {
    navigator.serviceWorker.register(swUrl, { scope: import.meta.env.BASE_URL }).catch((error: unknown) => {
      // Non-fatal: offline support simply won't be available this session.
      console.debug('[pwa] service worker registration skipped', error);
    });
  };
  // registerServiceWorker() is called from async app init, so `load` has usually
  // already fired — register now in that case, else wait for load.
  if (document.readyState === 'complete') register();
  else window.addEventListener('load', register, { once: true });
}
