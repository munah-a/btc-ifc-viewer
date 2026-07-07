/**
 * Dependency-free PWA plugin (W5.5 / C6). At build time it emits `sw.js` from
 * src/sw-template.js with a precache manifest of the ACTUAL built assets (all
 * hashed JS/CSS chunks + HTML + the public/ shell files: web-ifc.wasm,
 * worker.mjs, fonts, favicon, manifest). C1: every precached URL is same-origin.
 *
 * We hand-roll this instead of pulling in vite-plugin-pwa/Workbox to keep the
 * dependency surface minimal and the SW auditable — the whole SW is ~90 lines.
 *
 * Only applies to `vite build` (apply:'build'); dev serving is unaffected.
 */
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { resolve, join, relative } from 'node:path';

const PUBLIC_DIR = resolve(process.cwd(), 'public');

/** Recursively lists files under a dir as base-relative POSIX paths. */
function listFiles(dir, root = dir) {
  const out = [];
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) {
      out.push(...listFiles(full, root));
    } else {
      out.push(relative(root, full).split('\\').join('/'));
    }
  }
  return out;
}

export function pwaPlugin() {
  let base = '/';
  return {
    name: 'btc-pwa',
    apply: 'build',
    configResolved(config) {
      base = config.base || '/';
    },
    generateBundle(_options, bundle) {
      // 1. Every emitted chunk/asset (hashed JS, CSS, HTML, copied public assets).
      const emitted = new Set(Object.keys(bundle).map((name) => `${base}${name}`));

      // 2. public/ files that Vite copies verbatim (wasm, worker, favicon,
      //    manifest) — enumerate them so offline boot has the runtime assets.
      let publicFiles;
      try {
        publicFiles = listFiles(PUBLIC_DIR);
      } catch {
        publicFiles = [];
      }
      for (const file of publicFiles) emitted.add(`${base}${file}`);

      // Always include the app entry HTML for the navigation fallback.
      emitted.add(`${base}index.html`);
      emitted.add(base); // the start_url itself

      // Do NOT precache the SW itself or source maps.
      const precache = [...emitted].filter((url) => !url.endsWith('/sw.js') && !url.endsWith('.map'));
      precache.sort();

      const version = Date.now().toString(36);
      const template = readFileSync(resolve(process.cwd(), 'src/sw-template.js'), 'utf8');
      const source = template
        .replaceAll('__SW_VERSION__', version)
        .replaceAll('__PRECACHE_MANIFEST__', JSON.stringify(precache));

      this.emitFile({ type: 'asset', fileName: 'sw.js', source });
    },
  };
}
