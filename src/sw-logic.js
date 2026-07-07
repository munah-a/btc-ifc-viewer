/**
 * Pure service-worker decision helpers (W5-fixups P1/P2). Kept in their own
 * side-effect-free module so they are unit-testable in Node (the SW template
 * itself references `self`/`caches` and can't be imported directly). At build
 * time scripts/pwa-plugin.mjs inlines this file's body into the emitted sw.js
 * (with the `export ` keywords stripped), so the SW and the tests share ONE
 * source of truth — no drift.
 */

/**
 * P1: the precached shell name for a navigation, chosen by PATH. A navigation
 * whose path is (or is under) `/embed` falls back to the chromeless `embed.html`;
 * everything else falls back to the full-app `index.html`. Before this fix every
 * offline navigation — including /embed — served index.html, so an offline embed
 * rendered the full-app chrome.
 */
export function pickShellName(pathname) {
  return /(^|\/)embed(\.html)?(\/|$)/.test(pathname) ? 'embed.html' : 'index.html';
}

/**
 * P2: whether a navigation response may be written to the shell cache. Only a
 * good, same-origin (basic, non-opaque) response is cacheable — an unconditional
 * cache.put let a transient 500/503 overwrite the good cached `/` and then be
 * served offline (poisoned shell).
 */
export function shouldCacheNavigation(response) {
  return Boolean(response) && response.ok === true && response.type === 'basic';
}
