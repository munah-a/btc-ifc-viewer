/**
 * Storage adapter (W4.3) — the thin seam between the hosting API and the actual
 * bytes/metadata stores. Two implementations share one interface:
 *
 *   • InMemoryStorage  — used by unit tests (tests/unit/api-*.spec.ts). No
 *     network, fully deterministic; TTL/rate-limit windows are driven by an
 *     injectable clock so time-dependent logic is testable.
 *   • createRealStorage — used in production. Lazily `import()`s @vercel/blob
 *     (frag bytes) and @upstash/redis (metadata + rate-limit counters). These
 *     packages are NOT installed in this repo yet (the PO provisions Vercel Blob
 *     + a Redis store and adds the deps at deploy time — see the wave handoff),
 *     so the import is dynamic and only runs on Vercel. Nothing in the build or
 *     the test suite ever touches live storage.
 *
 * CONSTRAINT C2: the server stores and serves BYTES ONLY. This module never
 * parses, renders, or inspects the contents of a `.frag` — it puts/gets opaque
 * Uint8Array blobs and small JSON metadata records. No IFC/fragments code here.
 *
 * @vercel/kv note: Vercel KV is sunset; the current KV-style store is a
 * Marketplace Redis (Upstash) accessed via @upstash/redis. The interface below
 * is store-agnostic so swapping Redis for another KV needs no handler changes.
 */

/** Metadata stored per upload (KV/Redis). Carries the C4 entitlement seam fields. */
export interface UploadMeta {
  /** Public id (also the blob path stem and the /e/:id route param). */
  id: string;
  /** Direct Blob-CDN URL the browser fetches the `.frag` from (no function egress). */
  fragUrl: string;
  /** Blob pathname (needed to delete the object). */
  blobPath: string;
  /** Original file name (display only; never trusted for logic). */
  fileName: string;
  /** Byte size of the stored `.frag`. */
  sizeBytes: number;
  /** ISO timestamp the record was created. */
  createdAt: string;
  /** ISO timestamp after which the upload is expired and eligible for cleanup. */
  expiresAt: string;
  /**
   * C4 entitlement seam: today an anon salted-IP surrogate; a future paid tier
   * swaps this for a real account id with NO schema change.
   */
  ownerId: string;
  /** C4: 'anon' today; 'free'/'pro'/… later. */
  tier: string;
  /** SHA-256 hash of the delete token (never store the raw token). */
  deleteTokenHash: string;
}

/** Result of a rate-limit check: whether the hit is allowed + the current count. */
export interface RateLimitResult {
  allowed: boolean;
  count: number;
  limit: number;
}

/**
 * The storage seam. All methods are async so the real (network-backed) impl and
 * the in-memory test impl are interchangeable.
 */
export interface StorageAdapter {
  /**
   * Stores `.frag` bytes and returns the public Blob-CDN URL + the pathname.
   * `cacheControlMaxAgeSeconds` is applied at put-time (the Blob CDN host does
   * NOT read vercel.json headers — W4.3 requirement).
   */
  putFrag(
    path: string,
    bytes: Uint8Array,
    opts: { contentType: string; cacheControlMaxAgeSeconds: number },
  ): Promise<{ url: string; path: string }>;

  /**
   * Deletes a stored blob (idempotent). S3: `@vercel/blob` deletes BY URL, so
   * callers pass `meta.fragUrl` (the full Blob-CDN URL), not the pathname.
   */
  deleteFrag(fragUrl: string): Promise<void>;

  /** Reads an upload's metadata, or null if absent/expired-and-swept. */
  getMeta(id: string): Promise<UploadMeta | null>;

  /**
   * Writes an upload's metadata with a TTL (seconds). The store auto-expires the
   * record; cron cleanup also sweeps the blob (the blob store has no TTL).
   */
  setMeta(meta: UploadMeta, ttlSeconds: number): Promise<void>;

  /** Deletes an upload's metadata record (idempotent). */
  deleteMeta(id: string): Promise<void>;

  /** Number of non-expired uploads currently owned by `ownerId` (read-only view). */
  countActiveByOwner(ownerId: string): Promise<number>;

  /**
   * S1 (quota TOCTOU fix): ATOMICALLY reserve a quota slot for `ownerId`. Adds
   * `id` to the owner set and returns true only if the resulting active count is
   * ≤ `maxActive`; on overflow it removes `id` again and returns false. This is
   * the single check-and-act step the upload path uses instead of the racy
   * count-then-track pair (N concurrent POSTs can no longer all pass the check).
   * The reservation TTL matches the meta so it self-cleans. Caller rolls back
   * (untrackOwnerUpload) if a later step fails.
   */
  reserveOwnerSlot(ownerId: string, id: string, maxActive: number, ttlSeconds: number): Promise<boolean>;

  /** Removes an id from its owner's active set (on delete). */
  untrackOwnerUpload(ownerId: string, id: string): Promise<void>;

  /**
   * Fixed-window rate limit. Increments the counter for `key` and returns
   * whether the hit is within `limit` for the `windowSeconds` window.
   */
  rateLimit(key: string, limit: number, windowSeconds: number): Promise<RateLimitResult>;

  /** Lists ids whose metadata has expired (cron cleanup). May be a no-op if the store self-expires and cron only sweeps blobs via a separate index. */
  listExpired(now: number): Promise<UploadMeta[]>;
}

/** A clock function (ms since epoch). Injected so tests can control time. */
export type Clock = () => number;

// ─────────────────────────────────────────────────────────────────────────────
// In-memory implementation (tests). Deterministic; time driven by an injected
// clock so TTL expiry and rate-limit windows are exercisable without waiting.
// ─────────────────────────────────────────────────────────────────────────────

interface StoredMeta {
  meta: UploadMeta;
  /** ms-epoch expiry for the metadata record (mirrors the store TTL). */
  expiresAtMs: number;
}

interface RateWindow {
  count: number;
  windowStartMs: number;
}

export class InMemoryStorage implements StorageAdapter {
  private readonly blobs = new Map<string, { bytes: Uint8Array; contentType: string; maxAge: number }>();
  private readonly metas = new Map<string, StoredMeta>();
  private readonly ownerSets = new Map<string, Map<string, number>>(); // owner -> (id -> expiresAtMs)
  private readonly rateWindows = new Map<string, RateWindow>();
  private readonly host: string;

  constructor(
    private readonly now: Clock = () => Date.now(),
    host = 'https://blob.test.local',
  ) {
    this.host = host;
  }

  putFrag(
    path: string,
    bytes: Uint8Array,
    opts: { contentType: string; cacheControlMaxAgeSeconds: number },
  ): Promise<{ url: string; path: string }> {
    // Defensive copy so a caller mutating its buffer can't corrupt stored bytes.
    this.blobs.set(path, {
      bytes: bytes.slice(),
      contentType: opts.contentType,
      maxAge: opts.cacheControlMaxAgeSeconds,
    });
    return Promise.resolve({ url: `${this.host}/${path}`, path });
  }

  deleteFrag(fragUrl: string): Promise<void> {
    // S3: callers delete BY URL (`${host}/${path}`). Strip the host prefix to
    // recover the pathname key; tolerate a bare pathname too (idempotent).
    const path = fragUrl.startsWith(`${this.host}/`) ? fragUrl.slice(this.host.length + 1) : fragUrl;
    this.blobs.delete(path);
    return Promise.resolve();
  }

  getMeta(id: string): Promise<UploadMeta | null> {
    const stored = this.metas.get(id);
    if (!stored) return Promise.resolve(null);
    if (stored.expiresAtMs <= this.now()) {
      // Emulate the store auto-expiring the record.
      this.metas.delete(id);
      return Promise.resolve(null);
    }
    return Promise.resolve(stored.meta);
  }

  setMeta(meta: UploadMeta, ttlSeconds: number): Promise<void> {
    this.metas.set(meta.id, { meta, expiresAtMs: this.now() + ttlSeconds * 1000 });
    return Promise.resolve();
  }

  deleteMeta(id: string): Promise<void> {
    this.metas.delete(id);
    return Promise.resolve();
  }

  countActiveByOwner(ownerId: string): Promise<number> {
    const set = this.ownerSets.get(ownerId);
    if (!set) return Promise.resolve(0);
    const nowMs = this.now();
    let count = 0;
    for (const [id, expiresAtMs] of set) {
      if (expiresAtMs <= nowMs) {
        set.delete(id);
        continue;
      }
      // Count live reservations: an entry backed by a live meta, OR an in-flight
      // reservation whose meta hasn't been written yet (not yet expired). A meta
      // that was created then DELETED leaves no set entry (untrackOwnerUpload).
      const stored = this.metas.get(id);
      if (!stored || stored.expiresAtMs > nowMs) count += 1;
    }
    return Promise.resolve(count);
  }

  reserveOwnerSlot(ownerId: string, id: string, maxActive: number, ttlSeconds: number): Promise<boolean> {
    let set = this.ownerSets.get(ownerId);
    if (!set) {
      set = new Map();
      this.ownerSets.set(ownerId, set);
    }
    // Prune only EXPIRED entries so the count reflects live reservations. We do
    // NOT drop no-meta entries — those are in-flight reservations that must still
    // occupy a slot (that is the whole point of reserving before the meta write).
    const nowMs = this.now();
    for (const [entryId, expiresAtMs] of set) {
      if (expiresAtMs <= nowMs) set.delete(entryId);
    }
    // Atomic (single synchronous turn — no await between add and check): add the
    // id, then reject if that pushes the owner over the limit, rolling back.
    set.set(id, nowMs + ttlSeconds * 1000);
    if (set.size > maxActive) {
      set.delete(id);
      return Promise.resolve(false);
    }
    return Promise.resolve(true);
  }

  untrackOwnerUpload(ownerId: string, id: string): Promise<void> {
    this.ownerSets.get(ownerId)?.delete(id);
    return Promise.resolve();
  }

  rateLimit(key: string, limit: number, windowSeconds: number): Promise<RateLimitResult> {
    const nowMs = this.now();
    const windowMs = windowSeconds * 1000;
    const existing = this.rateWindows.get(key);
    if (!existing || nowMs - existing.windowStartMs >= windowMs) {
      this.rateWindows.set(key, { count: 1, windowStartMs: nowMs });
      return Promise.resolve({ allowed: 1 <= limit, count: 1, limit });
    }
    existing.count += 1;
    return Promise.resolve({ allowed: existing.count <= limit, count: existing.count, limit });
  }

  listExpired(now: number): Promise<UploadMeta[]> {
    const expired: UploadMeta[] = [];
    for (const [id, stored] of this.metas) {
      if (stored.expiresAtMs <= now) {
        expired.push(stored.meta);
        this.metas.delete(id);
      }
    }
    return Promise.resolve(expired);
  }

  // ── test-only helpers (not part of the interface) ──
  /** True if a blob exists at `path` (used by delete/cleanup tests). */
  hasBlob(path: string): boolean {
    return this.blobs.has(path);
  }
  /** Raw stored bytes (test assertions). */
  getBlobBytes(path: string): Uint8Array | undefined {
    return this.blobs.get(path)?.bytes;
  }
  /** Cache-control max-age recorded at put-time (test assertion). */
  getBlobMaxAge(path: string): number | undefined {
    return this.blobs.get(path)?.maxAge;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Real implementation (production, Vercel). Lazily imports the storage SDKs so
// this file compiles and the tests run without those packages installed. The PO
// adds `@vercel/blob` + `@upstash/redis` and provisions the stores at deploy.
// ─────────────────────────────────────────────────────────────────────────────

/** Minimal shapes we use from the SDKs (kept local so we don't depend on their types at build time). */
interface VercelBlobModule {
  put: (
    path: string,
    body: Uint8Array,
    opts: {
      access: 'public';
      contentType?: string;
      cacheControlMaxAge?: number;
      addRandomSuffix?: boolean;
      token?: string;
    },
  ) => Promise<{ url: string; pathname: string }>;
  del: (url: string, opts?: { token?: string }) => Promise<void>;
}

interface RedisLike {
  get: <T>(key: string) => Promise<T | null>;
  set: (key: string, value: unknown, opts?: { ex?: number }) => Promise<unknown>;
  del: (...keys: string[]) => Promise<number>;
  incr: (key: string) => Promise<number>;
  expire: (key: string, seconds: number) => Promise<unknown>;
  sadd: (key: string, ...members: string[]) => Promise<number>;
  srem: (key: string, ...members: string[]) => Promise<number>;
  scard: (key: string) => Promise<number>;
  smembers: (key: string) => Promise<string[]>;
}

const META_KEY = (id: string): string => `upload:${id}`;
const OWNER_KEY = (ownerId: string): string => `owner:${ownerId}`;
const RATE_KEY = (key: string): string => `rate:${key}`;

/**
 * Builds the production adapter. `deps` are injected in tests of the real
 * adapter's mapping logic; in production they default to the lazily-imported
 * SDKs. Throws if the SDKs/env are missing (fail fast at request time, never at
 * build time).
 */
export async function createRealStorage(deps?: {
  blob?: VercelBlobModule;
  redis?: RedisLike;
  blobToken?: string;
}): Promise<StorageAdapter> {
  // The storage SDKs are lazily imported so tests can inject fakes and the
  // function bundle only loads them on first use. Provisioned 2026-07-13: the
  // real packages are now dependencies (optional-deps.d.ts stubs removed). The
  // casts pin the REAL SDK surfaces to the minimal local shapes this adapter
  // calls — the SDK types are wider (extra body types / option unions), but
  // every call here is valid against both (put/del incl. cacheControlMaxAge +
  // addRandomSuffix; set{ex}/get/del/incr/expire/sadd/srem/scard/smembers).
  const blob = (deps?.blob ?? ((await import('@vercel/blob')) as unknown)) as VercelBlobModule;
  const redis = (deps?.redis ?? ((await import('@upstash/redis')).Redis.fromEnv() as unknown)) as RedisLike;
  const blobToken = deps?.blobToken ?? process.env.BLOB_READ_WRITE_TOKEN;

  return {
    async putFrag(path, bytes, opts) {
      const result = await blob.put(path, bytes, {
        access: 'public',
        contentType: opts.contentType,
        cacheControlMaxAge: opts.cacheControlMaxAgeSeconds,
        addRandomSuffix: false,
        token: blobToken,
      });
      return { url: result.url, path: result.pathname };
    },
    async deleteFrag(fragUrl) {
      // S3: @vercel/blob `del` deletes BY URL — callers pass meta.fragUrl.
      await blob.del(fragUrl, { token: blobToken });
    },
    async getMeta(id) {
      const meta = await redis.get<UploadMeta>(META_KEY(id));
      return meta ?? null;
    },
    async setMeta(meta, ttlSeconds) {
      await redis.set(META_KEY(meta.id), meta, { ex: ttlSeconds });
    },
    async deleteMeta(id) {
      await redis.del(META_KEY(id));
    },
    async countActiveByOwner(ownerId) {
      // Read-only view (SCARD). Stale entries self-clean via the set TTL and the
      // delete path (untrackOwnerUpload); we do NOT srem in-flight reservations
      // here (that would defeat reserveOwnerSlot's atomic slot accounting).
      return redis.scard(OWNER_KEY(ownerId));
    },
    async reserveOwnerSlot(ownerId, id, maxActive, ttlSeconds) {
      // S1: atomic reserve. SADD is atomic; SCARD reflects all concurrent adds.
      // If our add pushed the owner over the limit, roll it back and reject —
      // so N concurrent POSTs can store at most maxActive.
      const key = OWNER_KEY(ownerId);
      await redis.sadd(key, id);
      await redis.expire(key, ttlSeconds); // set TTL so idle owner sets self-clean
      const size = await redis.scard(key);
      if (size > maxActive) {
        await redis.srem(key, id);
        return false;
      }
      return true;
    },
    async untrackOwnerUpload(ownerId, id) {
      await redis.srem(OWNER_KEY(ownerId), id);
    },
    async rateLimit(key, limit, windowSeconds) {
      const redisKey = RATE_KEY(key);
      const count = await redis.incr(redisKey);
      if (count === 1) await redis.expire(redisKey, windowSeconds);
      return { allowed: count <= limit, count, limit };
    },
    listExpired() {
      // Redis auto-expires metadata; blob sweep for orphans is handled by the
      // cron reading the owner sets. Nothing to enumerate here (avoids a costly
      // full keyspace scan on the hot path).
      return Promise.resolve([]);
    },
  };
}
