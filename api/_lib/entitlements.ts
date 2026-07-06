/**
 * Entitlements (W4.3, constraint C4) — the ONE seam a future paid tier swaps.
 *
 * Today every request is anonymous and gets the anon defaults below. When a
 * paid tier arrives, `resolveEntitlements` takes a real account/tier instead of
 * always returning anon — and NOTHING else in the API changes, because every
 * quota/TTL/size decision flows through the returned `Entitlements` object and
 * upload metadata already carries `ownerId`/`tier`/`expiresAt`.
 *
 * Anon defaults (from the plan §5 W4.3): 50 MB per upload, 7-day TTL, 3 active
 * uploads per anon key. All overridable via env so the PO can tune without a
 * code change (kept within the ≤$20/mo cost envelope).
 */

export interface Entitlements {
  /** Tier label stored on the upload metadata (C4). */
  tier: string;
  /** Max bytes accepted for a single `.frag` upload. */
  maxUploadBytes: number;
  /** Time-to-live for a hosted upload, in days (C3). */
  ttlDays: number;
  /** Max simultaneously-active uploads per owner (quota). */
  maxActiveUploads: number;
}

const MB = 1024 * 1024;

/** Anon-tier defaults. Env overrides let the PO tune the cost envelope. */
export const ANON_DEFAULTS: Entitlements = {
  tier: 'anon',
  maxUploadBytes: readIntEnv('BTC_ANON_MAX_UPLOAD_MB', 50) * MB,
  ttlDays: readIntEnv('BTC_ANON_TTL_DAYS', 7),
  maxActiveUploads: readIntEnv('BTC_ANON_MAX_ACTIVE', 3),
};

/**
 * Resolves the entitlements for a request. Today: always anon. The signature
 * takes an optional owner/tier context so the paid-tier swap is additive.
 */
export function resolveEntitlements(_context?: { tier?: string }): Entitlements {
  // Future: look up `_context.tier` → per-tier limits. For now, anon only.
  return ANON_DEFAULTS;
}

/** TTL in seconds derived from an entitlement's `ttlDays`. */
export function ttlSeconds(entitlements: Entitlements): number {
  return entitlements.ttlDays * 24 * 60 * 60;
}

/** Reads a positive integer env var, falling back to `fallback` when unset/invalid. */
function readIntEnv(name: string, fallback: number): number {
  const raw = typeof process !== 'undefined' ? process.env[name] : undefined;
  if (!raw) return fallback;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

// ── Rate-limit policy (per anon key) — separate from per-upload quota. ──

/** Max uploads accepted per rate-limit window (defends the POST endpoint). */
export const RATE_LIMIT_MAX = readIntEnv('BTC_RATE_LIMIT_MAX', 10);
/** Rate-limit window, in seconds. */
export const RATE_LIMIT_WINDOW_SECONDS = readIntEnv('BTC_RATE_LIMIT_WINDOW_S', 60 * 60);
