import { describe, expect, it } from 'vitest';

import {
  hashFragBytes,
  InMemoryFragCache,
  shouldEvictFragKey,
  type FragCacheEntry,
} from '../../src/core/frag-cache';

describe('hashFragBytes (C8 — W5.2)', () => {
  it('is stable for identical byte buffers', () => {
    const a = new Uint8Array([1, 2, 3, 4, 5]);
    const b = new Uint8Array([1, 2, 3, 4, 5]);
    expect(hashFragBytes(a)).toBe(hashFragBytes(b));
  });

  it('differs for different content of the same length', () => {
    const a = new Uint8Array([1, 2, 3, 4]);
    const b = new Uint8Array([1, 2, 3, 5]);
    expect(hashFragBytes(a)).not.toBe(hashFragBytes(b));
  });

  it('differs for different lengths (length is part of the key)', () => {
    const a = new Uint8Array([1, 2, 3]);
    const b = new Uint8Array([1, 2, 3, 0]);
    expect(hashFragBytes(a)).not.toBe(hashFragBytes(b));
  });

  it('encodes the byte length in hex as a suffix', () => {
    const key = hashFragBytes(new Uint8Array(256));
    expect(key.endsWith('-100')).toBe(true); // 256 === 0x100
  });

  it('produces a deterministic key across many transpositions', () => {
    // A weak (order-insensitive) hash would collide on transpositions.
    const a = new Uint8Array([10, 20, 30, 40]);
    const b = new Uint8Array([40, 30, 20, 10]);
    expect(hashFragBytes(a)).not.toBe(hashFragBytes(b));
  });
});

describe('InMemoryFragCache adapter (C8 — W5.2)', () => {
  const entry = (key: string, bytes: number[]) => ({
    key,
    fileName: `${key}.frag`,
    sizeBytes: bytes.length,
    storedAt: 123,
    bytes: new Uint8Array(bytes),
  });

  it('put/get round-trips an entry with its bytes', async () => {
    const cache = new InMemoryFragCache();
    await cache.put(entry('k1', [1, 2, 3]));
    const got = await cache.get('k1');
    expect(got?.fileName).toBe('k1.frag');
    expect(Array.from(got?.bytes ?? [])).toEqual([1, 2, 3]);
  });

  it('returns null for a missing key', async () => {
    const cache = new InMemoryFragCache();
    expect(await cache.get('nope')).toBeNull();
  });

  it('lists keys and deletes/clears', async () => {
    const cache = new InMemoryFragCache();
    await cache.put(entry('k1', [1]));
    await cache.put(entry('k2', [2]));
    expect((await cache.keys()).sort()).toEqual(['k1', 'k2']);
    await cache.delete('k1');
    expect(await cache.get('k1')).toBeNull();
    expect(await cache.keys()).toEqual(['k2']);
    await cache.clear();
    expect(await cache.keys()).toEqual([]);
  });

  it('does not alias stored bytes with the caller (defensive copy)', async () => {
    const cache = new InMemoryFragCache();
    const src = entry('k1', [1, 2, 3]);
    await cache.put(src);
    src.bytes[0] = 99;
    const got = await cache.get('k1');
    expect(got?.bytes[0]).toBe(1);
  });
});

describe('shouldEvictFragKey (C2-race guard — W5-fixups review)', () => {
  const START = 1_000;

  it('evicts an unreferenced, not-newer key', () => {
    expect(shouldEvictFragKey('orphan', new Set(), START - 100, START)).toBe(true);
  });

  it('never evicts a referenced key (live model or persisted session)', () => {
    expect(shouldEvictFragKey('live', new Set(['live']), START - 100, START)).toBe(false);
  });

  it('never evicts a frag written AFTER the prune began (concurrent load, guard 2)', () => {
    // This is the unload-A-then-load-B race: B's put lands mid-prune with a
    // storedAt newer than pruneStartedAt, and B's record fragKey isn't stamped
    // yet (so it is not in `referenced`). It must survive.
    expect(shouldEvictFragKey('freshB', new Set(), START + 50, START)).toBe(false);
  });

  it('evicts an unreferenced key whose storedAt equals the prune start (boundary)', () => {
    expect(shouldEvictFragKey('edge', new Set(), START, START)).toBe(true);
  });

  it('evicts when the entry is missing (storedAt null) and unreferenced', () => {
    expect(shouldEvictFragKey('gone', new Set(), null, START)).toBe(true);
  });
});

describe('unload-A-then-load-B eviction ordering (C2-race — W5-fixups review)', () => {
  // Reproduces the headline race against the real adapter contract: mid-prune,
  // model B's freshly-converted frag is cached (put) with a NEW content hash and
  // a storedAt newer than the prune start. The guarded prune must NOT delete it.
  const mk = (key: string, storedAt: number): FragCacheEntry => ({
    key,
    fileName: `${key}.frag`,
    sizeBytes: 1,
    storedAt,
    bytes: new Uint8Array([1]),
  });

  it('leaves B\'s frag intact when its put lands after the prune snapshot', async () => {
    const cache = new InMemoryFragCache();
    const pruneStartedAt = 5_000;
    // A's frag existed before the prune and is now unreferenced (A was unloaded).
    await cache.put(mk('A', pruneStartedAt - 1_000));
    // B's frag is written AFTER the prune began (concurrent different-model load),
    // and B's record fragKey is not stamped yet, so `referenced` is empty.
    await cache.put(mk('B', pruneStartedAt + 500));

    const referenced = new Set<string>(); // nothing referenced yet
    const stored = await cache.keys();
    for (const key of stored) {
      const entry = await cache.get(key);
      if (shouldEvictFragKey(key, referenced, entry?.storedAt ?? null, pruneStartedAt)) {
        await cache.delete(key);
      }
    }

    // A (stale, unreferenced, older) is evicted; B (fresh, newer) survives so a
    // future restore does NOT silently re-convert it.
    expect(await cache.get('A')).toBeNull();
    expect(await cache.get('B')).not.toBeNull();
  });

  it('still evicts a genuinely orphaned older key', async () => {
    const cache = new InMemoryFragCache();
    const pruneStartedAt = 5_000;
    await cache.put(mk('old-orphan', pruneStartedAt - 2_000));
    const referenced = new Set<string>();
    for (const key of await cache.keys()) {
      const entry = await cache.get(key);
      if (shouldEvictFragKey(key, referenced, entry?.storedAt ?? null, pruneStartedAt)) {
        await cache.delete(key);
      }
    }
    expect(await cache.get('old-orphan')).toBeNull();
  });
});
