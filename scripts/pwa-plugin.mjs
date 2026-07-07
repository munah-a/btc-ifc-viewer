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
import { resolve, join, relative, basename } from 'node:path';

const PUBLIC_DIR = resolve(process.cwd(), 'public');

/**
 * P1 (W5-fixups): the HTML entry file names (e.g. `index.html`, `embed.html`)
 * from the resolved rollupOptions.input, so the precache manifest enumerates
 * EVERY navigable shell instead of hard-coding `index.html`. Vite emits the HTML
 * assets AFTER generateBundle, so we cannot read them off the bundle — derive
 * them from the configured inputs. Falls back to ['index.html'].
 */
export function htmlEntriesFromInput(input) {
  const values = Array.isArray(input)
    ? input
    : input && typeof input === 'object'
      ? Object.values(input)
      : input
        ? [input]
        : [];
  const htmls = values
    .filter((value) => typeof value === 'string' && value.endsWith('.html'))
    .map((value) => basename(value));
  return htmls.length > 0 ? [...new Set(htmls)] : ['index.html'];
}

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
  let htmlEntries = ['index.html'];
  return {
    name: 'btc-pwa',
    apply: 'build',
    configResolved(config) {
      base = config.base || '/';
      htmlEntries = htmlEntriesFromInput(config.build?.rollupOptions?.input);
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

      // P1 (W5-fixups): precache EVERY configured HTML entry (index.html AND
      // embed.html) so an offline navigation to /embed falls back to the real
      // chromeless embed shell — not the full-app index.html. Enumerated from the
      // resolved rollupOptions.input rather than hard-coded.
      for (const html of htmlEntries) emitted.add(`${base}${html}`);
      emitted.add(base); // the start_url itself

      // Do NOT precache the SW itself or source maps.
      const precache = [...emitted].filter((url) => !url.endsWith('/sw.js') && !url.endsWith('.map'));
      precache.sort();

      const version = Date.now().toString(36);
      const template = readFileSync(resolve(process.cwd(), 'src/sw-template.js'), 'utf8');
      // P1/P2 (W5-fixups): inline the pure SW decision helpers so the SW and the
      // unit tests share one source of truth. `sw-logic.js` uses `export` for the
      // test import; strip the keyword so the injected SW stays a classic worker
      // (no module semantics). Only leading `export ` on declarations is removed.
      const swLogic = readFileSync(resolve(process.cwd(), 'src/sw-logic.js'), 'utf8').replace(
        /^export\s+/gm,
        '',
      );
      const source = template
        .replace('// __SW_LOGIC__', swLogic)
        .replaceAll('__SW_VERSION__', version)
        .replaceAll('__PRECACHE_MANIFEST__', JSON.stringify(precache));

      this.emitFile({ type: 'asset', fileName: 'sw.js', source });
    },
  };
}
