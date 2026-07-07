import { describe, expect, it } from 'vitest';

// Pure SW decision helpers (build-time .js inlined into sw.js; typed via
// src/sw-logic.d.ts). Kept in their own module so they're unit-testable in Node.
import { pickShellName, shouldCacheNavigation } from '../../src/sw-logic.js';

/**
 * P1 (W5-fixups): the SW navigate fallback must be PATH-AWARE — a /embed
 * navigation falls back to the chromeless embed.html, everything else to the
 * full-app index.html. Before the fix every offline navigation served
 * index.html, so an offline embed rendered the full-app chrome.
 */
describe('pickShellName (P1 — W5-fixups path-aware navigate fallback)', () => {
  it('selects embed.html for the embed entry paths', () => {
    // A URL pathname never carries the query string, so /embed with ?m=… still
    // has pathname '/embed'.
    for (const p of ['/embed', '/embed.html', '/embed/', '/base/embed.html']) {
      expect(pickShellName(p)).toBe('embed.html');
    }
  });

  it('selects index.html for the app and every other path', () => {
    for (const p of ['/', '/index.html', '/viewer', '/base/', '/embedded', '/my-embedder']) {
      expect(pickShellName(p)).toBe('index.html');
    }
  });

  it('does not match a substring like /embedded as embed', () => {
    // Regression: the word boundary must not treat /embedded (a different route)
    // as the embed shell.
    expect(pickShellName('/embedded')).toBe('index.html');
  });
});

/**
 * P2 (W5-fixups): only a good, same-origin (basic) navigation may be cached — an
 * unconditional cache.put let a transient 5xx overwrite the good cached shell and
 * then be served offline (poisoned shell).
 */
describe('shouldCacheNavigation (P2 — W5-fixups cache-poisoning guard)', () => {
  it('caches a good same-origin (ok + basic) navigation', () => {
    expect(shouldCacheNavigation({ ok: true, type: 'basic' })).toBe(true);
  });

  it('does NOT cache a non-ok response (500/503 must not poison the shell)', () => {
    expect(shouldCacheNavigation({ ok: false, type: 'basic' })).toBe(false);
  });

  it('does NOT cache an opaque/cors (non-basic) response', () => {
    expect(shouldCacheNavigation({ ok: true, type: 'opaque' })).toBe(false);
    expect(shouldCacheNavigation({ ok: true, type: 'cors' })).toBe(false);
  });

  it('does NOT cache a null/undefined response (thrown fetch)', () => {
    expect(shouldCacheNavigation(null)).toBe(false);
    expect(shouldCacheNavigation(undefined)).toBe(false);
  });
});
