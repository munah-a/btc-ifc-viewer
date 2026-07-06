import { describe, expect, it } from 'vitest';

import { InMemoryStorage, createRealStorage, type UploadMeta } from '../../api/_lib/storage';

function makeClock(start = 1_000_000): { now: () => number; advance: (ms: number) => void } {
  let t = start;
  return { now: () => t, advance: (ms) => (t += ms) };
}

function meta(id: string, ownerId = 'owner-1'): UploadMeta {
  return {
    id,
    fragUrl: `https://blob.test.local/frags/${id}.frag`,
    blobPath: `frags/${id}.frag`,
    fileName: `${id}.frag`,
    sizeBytes: 100,
    createdAt: new Date(0).toISOString(),
    expiresAt: new Date(0).toISOString(),
    ownerId,
    tier: 'anon',
    deleteTokenHash: 'hash',
  };
}

describe('InMemoryStorage · blobs', () => {
  it('puts and serves opaque bytes with a stored max-age; defensive-copies the input', async () => {
    const s = new InMemoryStorage();
    const bytes = new Uint8Array([9, 8, 7]);
    const { url, path } = await s.putFrag('frags/x.frag', bytes, {
      contentType: 'application/octet-stream',
      cacheControlMaxAgeSeconds: 31536000,
    });
    expect(url).toContain('frags/x.frag');
    expect(path).toBe('frags/x.frag');
    // Mutating the caller's buffer must NOT corrupt stored bytes (defensive copy).
    bytes[0] = 0;
    expect([...(s.getBlobBytes('frags/x.frag') ?? [])]).toEqual([9, 8, 7]);
    expect(s.getBlobMaxAge('frags/x.frag')).toBe(31536000);

    // S3: deleteFrag takes the full Blob-CDN URL (what meta.fragUrl holds).
    await s.deleteFrag(url);
    expect(s.hasBlob('frags/x.frag')).toBe(false);
  });

  it('deleteFrag also tolerates a bare pathname (idempotent)', async () => {
    const s = new InMemoryStorage();
    await s.putFrag('frags/y.frag', new Uint8Array([1]), {
      contentType: 'application/octet-stream',
      cacheControlMaxAgeSeconds: 60,
    });
    await s.deleteFrag('frags/y.frag');
    expect(s.hasBlob('frags/y.frag')).toBe(false);
  });
});

describe('InMemoryStorage · metadata TTL', () => {
  it('expires a record after its TTL elapses (getMeta returns null)', async () => {
    const clock = makeClock();
    const s = new InMemoryStorage(clock.now);
    await s.setMeta(meta('a'), 60); // 60s TTL
    expect(await s.getMeta('a')).not.toBeNull();
    clock.advance(59_000);
    expect(await s.getMeta('a')).not.toBeNull();
    clock.advance(2_000);
    expect(await s.getMeta('a')).toBeNull();
  });

  it('listExpired returns and removes only expired records', async () => {
    const clock = makeClock();
    const s = new InMemoryStorage(clock.now);
    await s.setMeta(meta('short'), 10);
    await s.setMeta(meta('long'), 1000);
    clock.advance(20_000);
    const expired = await s.listExpired(clock.now());
    expect(expired.map((m) => m.id)).toEqual(['short']);
    expect(await s.getMeta('long')).not.toBeNull();
  });
});

describe('InMemoryStorage · owner reservation (quota, S1)', () => {
  it('reserves slots up to maxActive, then rejects; delete frees a slot', async () => {
    const clock = makeClock();
    const s = new InMemoryStorage(clock.now);
    expect(await s.reserveOwnerSlot('owner-1', 'a', 3, 1000)).toBe(true);
    await s.setMeta(meta('a'), 1000);
    expect(await s.reserveOwnerSlot('owner-1', 'b', 3, 1000)).toBe(true);
    await s.setMeta(meta('b'), 1000);
    expect(await s.reserveOwnerSlot('owner-1', 'c', 3, 1000)).toBe(true);
    await s.setMeta(meta('c'), 1000);
    expect(await s.countActiveByOwner('owner-1')).toBe(3);
    // 4th over the cap is rejected (and not counted).
    expect(await s.reserveOwnerSlot('owner-1', 'd', 3, 1000)).toBe(false);
    expect(await s.countActiveByOwner('owner-1')).toBe(3);

    // Deleting one frees a slot.
    await s.deleteMeta('a');
    await s.untrackOwnerUpload('owner-1', 'a');
    expect(await s.reserveOwnerSlot('owner-1', 'e', 3, 1000)).toBe(true);
  });

  it('an in-flight reservation (no meta yet) still occupies a slot', async () => {
    const s = new InMemoryStorage();
    expect(await s.reserveOwnerSlot('o', 'a', 1, 1000)).toBe(true); // reserved, no setMeta
    // A second reserve for the same owner must fail while the first is in flight.
    expect(await s.reserveOwnerSlot('o', 'b', 1, 1000)).toBe(false);
  });

  it('expired reservations free up over time', async () => {
    const clock = makeClock();
    const s = new InMemoryStorage(clock.now);
    expect(await s.reserveOwnerSlot('o', 'a', 1, 10)).toBe(true);
    expect(await s.reserveOwnerSlot('o', 'b', 1, 10)).toBe(false);
    clock.advance(20_000); // first reservation TTL elapses
    expect(await s.reserveOwnerSlot('o', 'c', 1, 10)).toBe(true);
  });
});

describe('InMemoryStorage · rate limiting (fixed window)', () => {
  it('allows up to the limit, blocks beyond, then resets after the window', async () => {
    const clock = makeClock();
    const s = new InMemoryStorage(clock.now);
    const hit = () => s.rateLimit('k', 3, 60);
    expect((await hit()).allowed).toBe(true); // 1
    expect((await hit()).allowed).toBe(true); // 2
    expect((await hit()).allowed).toBe(true); // 3
    const blocked = await hit(); // 4
    expect(blocked.allowed).toBe(false);
    expect(blocked.count).toBe(4);

    // Advance past the window → counter resets.
    clock.advance(61_000);
    expect((await hit()).allowed).toBe(true);
  });

  it('keys are independent', async () => {
    const s = new InMemoryStorage();
    expect((await s.rateLimit('a', 1, 60)).allowed).toBe(true);
    expect((await s.rateLimit('a', 1, 60)).allowed).toBe(false);
    expect((await s.rateLimit('b', 1, 60)).allowed).toBe(true); // different key
  });
});

describe('createRealStorage · adapter mapping (injected fakes, no live Vercel)', () => {
  it('maps putFrag/getMeta/rateLimit onto the injected blob + redis deps', async () => {
    const kv = new Map<string, unknown>();
    const sets = new Map<string, Set<string>>();
    const counters = new Map<string, number>();

    const delCalls: string[] = [];
    const fakeBlob = {
      put: (path: string, _bytes: Uint8Array, opts: { cacheControlMaxAge?: number }) =>
        Promise.resolve({ url: `https://blob.cdn/${path}`, pathname: path, _maxAge: opts.cacheControlMaxAge }),
      del: (url: string) => {
        delCalls.push(url);
        return Promise.resolve();
      },
    };
    const fakeRedis = {
      get: <T>(key: string) => Promise.resolve((kv.get(key) as T) ?? null),
      set: (key: string, value: unknown) => {
        kv.set(key, value);
        return Promise.resolve('OK');
      },
      del: (...keys: string[]) => {
        keys.forEach((k) => kv.delete(k));
        return Promise.resolve(keys.length);
      },
      incr: (key: string) => {
        const n = (counters.get(key) ?? 0) + 1;
        counters.set(key, n);
        return Promise.resolve(n);
      },
      expire: (_key: string, _s: number) => Promise.resolve(1),
      sadd: (key: string, ...members: string[]) => {
        const set = sets.get(key) ?? new Set<string>();
        members.forEach((m) => set.add(m));
        sets.set(key, set);
        return Promise.resolve(members.length);
      },
      srem: (key: string, ...members: string[]) => {
        const set = sets.get(key);
        members.forEach((m) => set?.delete(m));
        return Promise.resolve(members.length);
      },
      scard: (key: string) => Promise.resolve(sets.get(key)?.size ?? 0),
      smembers: (key: string) => Promise.resolve([...(sets.get(key) ?? [])]),
    };

    const storage = await createRealStorage({ blob: fakeBlob, redis: fakeRedis });

    const put = await storage.putFrag('frags/z.frag', new Uint8Array([1]), {
      contentType: 'application/octet-stream',
      cacheControlMaxAgeSeconds: 42,
    });
    expect(put.url).toBe('https://blob.cdn/frags/z.frag');
    expect(put.path).toBe('frags/z.frag');

    await storage.setMeta(meta('z'), 100);
    expect((await storage.getMeta('z'))?.id).toBe('z');

    // S1: reserveOwnerSlot maps onto sadd+scard (atomic bounded set).
    expect(await storage.reserveOwnerSlot('owner-x', 'z', 3, 100)).toBe(true);
    expect(await storage.countActiveByOwner('owner-x')).toBe(1);
    // Over the cap → rolled back via srem, returns false.
    expect(await storage.reserveOwnerSlot('owner-x', 'z2', 1, 100)).toBe(false);
    expect(await storage.countActiveByOwner('owner-x')).toBe(1);

    // S3: deleteFrag calls blob.del with the FULL URL (not the pathname).
    await storage.deleteFrag('https://blob.cdn/frags/z.frag');
    expect(delCalls).toContain('https://blob.cdn/frags/z.frag');

    const r1 = await storage.rateLimit('rk', 1, 60);
    expect(r1.allowed).toBe(true);
    const r2 = await storage.rateLimit('rk', 1, 60);
    expect(r2.allowed).toBe(false);
  });
});
