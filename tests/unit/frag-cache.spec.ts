import { describe, expect, it } from 'vitest';

import { hashFragBytes, InMemoryFragCache } from '../../src/core/frag-cache';

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
