/**
 * Vendor runtime assets out of node_modules into public/ (AUDIT A2/P2).
 *
 * The viewer must not depend on any CDN at runtime (constraint C1: offline /
 * field usage). web-ifc's WASM and the fragments worker are copied from the
 * exact installed package versions, so they always track package-lock.json.
 * Runs automatically via the prebuild/predev npm hooks.
 */
import { copyFileSync, mkdirSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const ASSETS = [
  {
    from: join(repoRoot, 'node_modules', 'web-ifc', 'web-ifc.wasm'),
    to: join(repoRoot, 'public', 'web-ifc.wasm'),
  },
  {
    from: join(repoRoot, 'node_modules', '@thatopen', 'fragments', 'dist', 'Worker', 'worker.mjs'),
    to: join(repoRoot, 'public', 'worker.mjs'),
  },
];

mkdirSync(join(repoRoot, 'public'), { recursive: true });

for (const { from, to } of ASSETS) {
  try {
    copyFileSync(from, to);
    const { size } = statSync(to);
    console.log(`[vendor-assets] ${from} -> ${to} (${(size / 1024).toFixed(0)} kB)`);
  } catch (error) {
    console.error(`[vendor-assets] FAILED copying ${from}: ${String(error)}`);
    console.error('[vendor-assets] Did you run npm install?');
    process.exit(1);
  }
}
