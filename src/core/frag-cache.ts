/**
 * Fragments byte cache (W5.2 / C8). Loaded models are cached as their converted
 * `.frag` bytes so a full session can be RESTORED without re-converting the IFC.
 * localStorage is far too small (a `.frag` is MBs), so the store is IndexedDB.
 *
 * The store is behind a small `FragCacheAdapter` seam:
 *  - `IndexedDbFragCache` — the production adapter (browser IndexedDB).
 *  - `InMemoryFragCache` — a pure, dependency-free adapter used by unit tests
 *    (node has no IndexedDB and we do not want a fake-indexeddb dependency, C1).
 *
 * Keys are a stable content hash of the `.frag` bytes (`hashFragBytes`) so the
 * same model reused across sessions maps to one cache entry, and stale entries
 * can be pruned by key.
 */

/** A stored fragments entry: the bytes plus light metadata for pruning/UX. */
export interface FragCacheEntry {
  /** Stable content-hash key. */
  key: string;
  /** Original file name (for display on restore). */
  fileName: string;
  /** `.frag` byte length. */
  sizeBytes: number;
  /** Epoch ms the entry was written (LRU-ish pruning). */
  storedAt: number;
  /** The converted fragments bytes. */
  bytes: Uint8Array;
}

/**
 * C2-race guard (W5-fixups review): the pure decision for whether a candidate
 * frag key may be evicted during a prune. A key is evictable ONLY when it is not
 * referenced (by any live model or the persisted session) AND its stored entry is
 * not newer than the moment the prune began. The `storedAt` guard protects a frag
 * written by a load that raced the eviction — it isn't referenced yet only because
 * its model record's fragKey hasn't been stamped. `referenced` is passed freshly
 * re-computed at decision time (not a stale start-of-prune snapshot).
 */
export function shouldEvictFragKey(
  key: string,
  referenced: ReadonlySet<string>,
  entryStoredAt: number | null,
  pruneStartedAt: number,
): boolean {
  if (referenced.has(key)) return false;
  if (entryStoredAt !== null && entryStoredAt > pruneStartedAt) return false;
  return true;
}

/** Minimal async KV contract over `.frag` byte blobs. */
export interface FragCacheAdapter {
  get(key: string): Promise<FragCacheEntry | null>;
  put(entry: FragCacheEntry): Promise<void>;
  delete(key: string): Promise<void>;
  /** All stored keys (used to prune entries no longer referenced by a session). */
  keys(): Promise<string[]>;
  clear(): Promise<void>;
}

/**
 * Stable, dependency-free content hash of the `.frag` bytes (FNV-1a, 64-bit as
 * two 32-bit lanes rendered hex). Not cryptographic — it only needs to be a
 * stable, low-collision key for identical byte buffers. C1: no external dep.
 */
export function hashFragBytes(bytes: Uint8Array): string {
  // Two independent FNV-1a lanes widen the effective space to ~64 bits.
  let h1 = 0x811c9dc5;
  let h2 = 0x811c9dc5 ^ 0x9e3779b9;
  for (let i = 0; i < bytes.length; i += 1) {
    const b = bytes[i];
    h1 ^= b;
    // 32-bit FNV prime multiply via shifts (avoids float precision loss).
    h1 = (h1 + ((h1 << 1) + (h1 << 4) + (h1 << 7) + (h1 << 8) + (h1 << 24))) >>> 0;
    h2 ^= b;
    h2 = (h2 + ((h2 << 1) + (h2 << 4) + (h2 << 7) + (h2 << 8) + (h2 << 24))) >>> 0;
  }
  const hex = (n: number): string => (n >>> 0).toString(16).padStart(8, '0');
  return `${hex(h1)}${hex(h2)}-${bytes.length.toString(16)}`;
}

/** Pure in-memory adapter (unit tests + a safe fallback when IDB is absent). */
export class InMemoryFragCache implements FragCacheAdapter {
  private readonly store = new Map<string, FragCacheEntry>();

  get(key: string): Promise<FragCacheEntry | null> {
    const entry = this.store.get(key);
    // Copy bytes out so callers cannot mutate the stored buffer (mirrors the
    // structured-clone semantics of the real IndexedDB adapter).
    return Promise.resolve(entry ? { ...entry, bytes: new Uint8Array(entry.bytes) } : null);
  }

  put(entry: FragCacheEntry): Promise<void> {
    // Copy bytes in so a later mutation of the caller's buffer doesn't corrupt
    // the cache (IndexedDB structured-clones on write).
    this.store.set(entry.key, { ...entry, bytes: new Uint8Array(entry.bytes) });
    return Promise.resolve();
  }

  delete(key: string): Promise<void> {
    this.store.delete(key);
    return Promise.resolve();
  }

  keys(): Promise<string[]> {
    return Promise.resolve([...this.store.keys()]);
  }

  clear(): Promise<void> {
    this.store.clear();
    return Promise.resolve();
  }
}

const DB_NAME = 'btc-viewer-frag-cache';
const DB_VERSION = 1;
const STORE = 'frags';

/**
 * IndexedDB-backed adapter. Opens a single object store keyed by `key`. Values
 * are the full FragCacheEntry (IndexedDB structured-clones the Uint8Array).
 * A failed open (private mode, quota, unsupported) surfaces as a rejected
 * promise the caller can degrade from (session save/restore becomes a no-op).
 */
export class IndexedDbFragCache implements FragCacheAdapter {
  private dbPromise: Promise<IDBDatabase> | null = null;

  private openDb(): Promise<IDBDatabase> {
    if (this.dbPromise) return this.dbPromise;
    this.dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
      if (typeof indexedDB === 'undefined') {
        reject(new Error('IndexedDB is not available in this environment'));
        return;
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: 'key' });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error('Failed to open frag cache'));
    });
    return this.dbPromise;
  }

  private async withStore(
    mode: IDBTransactionMode,
    run: (store: IDBObjectStore) => IDBRequest,
  ): Promise<unknown> {
    const db = await this.openDb();
    return new Promise<unknown>((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const request = run(tx.objectStore(STORE));
      // C1 (W5-fixups): a readwrite `put`/`delete`/`clear` request fires
      // `onsuccess` BEFORE the transaction commits, so resolving there swallows a
      // commit-time abort (e.g. quota on flush, tab unload) — a "saved" frag never
      // persists and restore silently drops the model. So for readwrite, resolve
      // on `tx.oncomplete` (the real durable point) and reject on
      // `tx.onabort`/`tx.onerror`, capturing `request.result` for get-like reads.
      // Reads (readonly) can safely resolve on `request.onsuccess`.
      if (mode === 'readonly') {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error ?? new Error('Frag cache request failed'));
        return;
      }
      let result: unknown;
      request.onsuccess = () => {
        result = request.result;
      };
      request.onerror = () => reject(request.error ?? new Error('Frag cache request failed'));
      tx.oncomplete = () => resolve(result);
      tx.onabort = () => reject(tx.error ?? new Error('Frag cache transaction aborted'));
      tx.onerror = () => reject(tx.error ?? new Error('Frag cache transaction failed'));
    });
  }

  async get(key: string): Promise<FragCacheEntry | null> {
    const result = await this.withStore('readonly', (store) => store.get(key));
    return (result as FragCacheEntry | undefined) ?? null;
  }

  async put(entry: FragCacheEntry): Promise<void> {
    await this.withStore('readwrite', (store) => store.put(entry));
  }

  async delete(key: string): Promise<void> {
    await this.withStore('readwrite', (store) => store.delete(key));
  }

  async keys(): Promise<string[]> {
    const result = (await this.withStore('readonly', (store) => store.getAllKeys())) as IDBValidKey[];
    // The store's keyPath is the string `key`, so every stored key is a string.
    return result.map((key) => (typeof key === 'string' ? key : JSON.stringify(key)));
  }

  async clear(): Promise<void> {
    await this.withStore('readwrite', (store) => store.clear());
  }
}
