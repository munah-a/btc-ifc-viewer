import { afterEach, describe, expect, it } from 'vitest';

import { IndexedDbFragCache, type FragCacheEntry } from '../../src/core/frag-cache';

/**
 * C1 + M3 (W5-fixups): IndexedDbFragCache error/commit-path coverage with a
 * hand-rolled fake `indexedDB` (NO external dep — fake-indexeddb would violate
 * C1/dep discipline). Asserts:
 *  - a failed open degrades to a caught rejection;
 *  - a readwrite `put` resolves on `tx.oncomplete` (durable commit), NOT the
 *    earlier `request.onsuccess`, and REJECTS on `tx.onabort` (guards C1's fix —
 *    a commit-time abort must not be swallowed as a silent success);
 *  - `get` returns null on a missing key.
 *
 * The fake models the async event ordering that trips C1: the request fires
 * onsuccess synchronously-ish, and the transaction fires oncomplete/onabort a
 * tick later.
 */

// ---- minimal fake IndexedDB ----------------------------------------------

type Handler = (() => void) | null;

class FakeRequest<T = unknown> {
  onsuccess: Handler = null;
  onerror: Handler = null;
  result: T | undefined;
  error: Error | null = null;
}

class FakeOpenRequest extends FakeRequest<FakeDb> {
  onupgradeneeded: Handler = null;
}

interface FakeTxConfig {
  /** If set, the tx aborts on commit with this error instead of completing. */
  abortError?: Error;
}

class FakeObjectStore {
  constructor(
    private readonly data: Map<string, FragCacheEntry>,
    private readonly tx: FakeTransaction,
  ) {}

  private request<T>(compute: () => T): FakeRequest<T> {
    const req = new FakeRequest<T>();
    // Fire the request success on a microtask, then let the tx settle after.
    queueMicrotask(() => {
      try {
        req.result = compute();
        req.onsuccess?.();
        this.tx.scheduleSettle();
      } catch (error) {
        req.error = error as Error;
        req.onerror?.();
        this.tx.scheduleAbort(error as Error);
      }
    });
    return req;
  }

  get(key: string): FakeRequest<FragCacheEntry | undefined> {
    return this.request(() => this.data.get(key));
  }

  put(entry: FragCacheEntry): FakeRequest<string> {
    return this.request(() => {
      // The tx may be configured to abort on commit (e.g. quota on flush).
      this.data.set(entry.key, entry);
      return entry.key;
    });
  }

  delete(key: string): FakeRequest<undefined> {
    return this.request(() => {
      this.data.delete(key);
      return undefined;
    });
  }

  getAllKeys(): FakeRequest<string[]> {
    return this.request(() => [...this.data.keys()]);
  }

  clear(): FakeRequest<undefined> {
    return this.request(() => {
      this.data.clear();
      return undefined;
    });
  }
}

class FakeTransaction {
  oncomplete: Handler = null;
  onabort: Handler = null;
  onerror: Handler = null;
  error: Error | null = null;
  private settled = false;

  constructor(
    private readonly data: Map<string, FragCacheEntry>,
    private readonly config: FakeTxConfig,
  ) {}

  objectStore(): FakeObjectStore {
    return new FakeObjectStore(this.data, this);
  }

  scheduleSettle(): void {
    if (this.settled) return;
    this.settled = true;
    queueMicrotask(() => {
      if (this.config.abortError) {
        this.error = this.config.abortError;
        this.onabort?.();
      } else {
        this.oncomplete?.();
      }
    });
  }

  scheduleAbort(error: Error): void {
    if (this.settled) return;
    this.settled = true;
    queueMicrotask(() => {
      this.error = error;
      this.onabort?.();
    });
  }
}

class FakeDb {
  objectStoreNames = { contains: () => true };
  constructor(private readonly txConfig: FakeTxConfig = {}) {}
  private readonly data = new Map<string, FragCacheEntry>();
  createObjectStore(): void {}
  transaction(): FakeTransaction {
    return new FakeTransaction(this.data, this.txConfig);
  }
  close(): void {}
}

interface FakeFactoryOptions {
  failOpen?: boolean;
  txConfig?: FakeTxConfig;
}

const installFakeIndexedDb = (opts: FakeFactoryOptions = {}): FakeDb => {
  const db = new FakeDb(opts.txConfig);
  (globalThis as unknown as { indexedDB: unknown }).indexedDB = {
    open: () => {
      const req = new FakeOpenRequest();
      queueMicrotask(() => {
        if (opts.failOpen) {
          req.error = new Error('open blocked');
          req.onerror?.();
          return;
        }
        req.result = db;
        req.onupgradeneeded?.();
        req.onsuccess?.();
      });
      return req;
    },
  };
  return db;
};

const entry = (key: string): FragCacheEntry => ({
  key,
  fileName: `${key}.frag`,
  sizeBytes: 3,
  storedAt: 1,
  bytes: new Uint8Array([1, 2, 3]),
});

describe('IndexedDbFragCache error/commit paths (C1 + M3 — W5-fixups)', () => {
  const original = (globalThis as unknown as { indexedDB?: unknown }).indexedDB;

  afterEach(() => {
    (globalThis as unknown as { indexedDB?: unknown }).indexedDB = original;
  });

  it('degrades to a caught rejection when the DB fails to open', async () => {
    installFakeIndexedDb({ failOpen: true });
    const cache = new IndexedDbFragCache();
    await expect(cache.get('k')).rejects.toThrow('open blocked');
  });

  it('get returns null on a missing key', async () => {
    installFakeIndexedDb();
    const cache = new IndexedDbFragCache();
    expect(await cache.get('missing')).toBeNull();
  });

  it('put resolves on the durable transaction commit (round-trips)', async () => {
    installFakeIndexedDb();
    const cache = new IndexedDbFragCache();
    await cache.put(entry('k1'));
    const got = await cache.get('k1');
    expect(got?.key).toBe('k1');
  });

  it('put REJECTS when the transaction aborts on commit (guards C1)', async () => {
    // The request onsuccess fires, but the transaction aborts at commit time
    // (e.g. quota on flush). C1: this must reject, not resolve as a false success.
    installFakeIndexedDb({ txConfig: { abortError: new Error('quota exceeded on flush') } });
    const cache = new IndexedDbFragCache();
    await expect(cache.put(entry('k1'))).rejects.toThrow('quota exceeded on flush');
  });

  it('keys() lists stored keys after a durable put', async () => {
    installFakeIndexedDb();
    const cache = new IndexedDbFragCache();
    await cache.put(entry('a'));
    await cache.put(entry('b'));
    expect((await cache.keys()).sort()).toEqual(['a', 'b']);
  });

  it('delete removes a key durably', async () => {
    installFakeIndexedDb();
    const cache = new IndexedDbFragCache();
    await cache.put(entry('a'));
    await cache.delete('a');
    expect(await cache.get('a')).toBeNull();
  });
});
