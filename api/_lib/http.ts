/**
 * Shared HTTP + crypto helpers for the hosting API (W4.3).
 *
 * Security-sensitive surface (upload/quota/delete-token/rate-limit), so these
 * are written defensively:
 *   • `ownerIdFromRequest` derives a STABLE, NON-reversible surrogate owner id
 *     from the client IP + a server-side salt (C4). We never store a raw IP.
 *   • delete tokens are random 256-bit values returned ONCE to the uploader;
 *     only their SHA-256 hash is stored, and comparison is constant-time.
 *   • `json`/`error` build Web-standard Responses with a strict content type and
 *     no-store caching for the JSON API responses (the `.frag` bytes are cached
 *     separately at Blob put-time).
 *
 * Web-standard (Request/Response/crypto.subtle) plus `node:crypto` for the
 * constant-time compare — runs on the Vercel Node.js runtime and is
 * unit-testable in Node without any Vercel packages.
 */
import { timingSafeEqual } from 'node:crypto';

/**
 * True on a real Vercel deployment (Vercel sets `VERCEL=1`). Used to FAIL CLOSED
 * on missing security config (CRON_SECRET / BTC_OWNER_SALT) in production while
 * still letting local tests exercise the code paths.
 */
export function isProduction(): boolean {
  return typeof process !== 'undefined' && Boolean(process.env.VERCEL);
}

/** SHA-256 of a string → lowercase hex. */
export async function sha256Hex(input: string): Promise<string> {
  const data = new TextEncoder().encode(input);
  const digest = await crypto.subtle.digest('SHA-256', data);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Derives a non-reversible owner surrogate from the TRUSTED client IP + a server
 * salt. This is the C4 seam value stored as `ownerId`. S7: the salt is required
 * in production — a missing BTC_OWNER_SALT there is a fatal misconfig (throws),
 * so we never silently fall back to a predictable dev salt on a live deployment.
 */
export async function ownerIdFromRequest(request: Request): Promise<string> {
  const ip = clientIp(request) ?? 'no-ip';
  const salt = ownerSalt();
  const hash = await sha256Hex(`${salt}:${ip}`);
  return `anon_${hash.slice(0, 24)}`;
}

/** The owner-id salt. Fails closed in production (S7); dev/test gets a fixed salt. */
function ownerSalt(): string {
  const configured = typeof process !== 'undefined' ? process.env.BTC_OWNER_SALT : undefined;
  if (configured) return configured;
  if (isProduction()) {
    throw new Error('BTC_OWNER_SALT must be set in production (owner-id salting).');
  }
  return 'btc-ifc-viewer-dev-salt';
}

/**
 * The TRUSTED client IP (S2 — anti-spoof). `x-forwarded-for` is a hop list that
 * the CLIENT controls at the LEFT; anything the client sends there is untrusted
 * and would let an attacker mint a fresh quota/rate bucket per request. On
 * Vercel the platform-set `x-real-ip` (and `x-vercel-forwarded-for`) carry the
 * actual connecting IP, so we prefer those. If only `x-forwarded-for` is
 * present we take the RIGHT-MOST hop (the one appended closest to our
 * infrastructure), never the client-supplied left. Used solely to bucket
 * quota/rate-limit and immediately salted+hashed — never stored raw.
 */
export function clientIp(request: Request): string | null {
  // Platform-trusted single-value headers first.
  const real = request.headers.get('x-real-ip')?.trim();
  if (real) return real;
  const vercelXff = request.headers.get('x-vercel-forwarded-for')?.trim();
  if (vercelXff) {
    const hops = vercelXff.split(',').map((h) => h.trim()).filter(Boolean);
    if (hops.length > 0) return hops[hops.length - 1];
  }
  // Fall back to the RIGHT-MOST x-forwarded-for hop (closest to us), NOT the
  // left-most client-controlled value.
  const xff = request.headers.get('x-forwarded-for');
  if (xff) {
    const hops = xff.split(',').map((h) => h.trim()).filter(Boolean);
    if (hops.length > 0) return hops[hops.length - 1];
  }
  return null;
}

/** Generates a URL-safe random token (default 256-bit) for delete auth. */
export function generateToken(bytes = 32): string {
  const buf = new Uint8Array(bytes);
  crypto.getRandomValues(buf);
  let binary = '';
  for (const b of buf) binary += String.fromCharCode(b);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/**
 * Constant-time comparison of a presented delete token against the stored hash.
 * S6: both sides are hashed to fixed-length hex, then compared with Node's
 * `crypto.timingSafeEqual` (a true constant-time primitive) instead of a
 * hand-rolled XOR loop. Length is equal by construction (SHA-256 hex = 64
 * chars); a defensive length guard avoids timingSafeEqual throwing.
 */
export async function verifyToken(presentedToken: string, storedHash: string): Promise<boolean> {
  if (!presentedToken || !storedHash) return false;
  const presentedHash = await sha256Hex(presentedToken);
  if (presentedHash.length !== storedHash.length) return false;
  const a = new TextEncoder().encode(presentedHash);
  const b = new TextEncoder().encode(storedHash);
  return timingSafeEqualBytes(a, b);
}

/**
 * Constant-time byte comparison via Node's `crypto.timingSafeEqual` (a true
 * constant-time primitive). Both inputs must be equal length (guarded here).
 */
function timingSafeEqualBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

/** Generates a short, URL-safe public upload id (unpredictable). */
export function generateId(): string {
  return generateToken(12); // 96-bit → ~16 base64url chars
}

const JSON_HEADERS: Record<string, string> = {
  'content-type': 'application/json; charset=utf-8',
  // API JSON is never cached (the .frag bytes are cached at Blob put-time).
  'cache-control': 'no-store',
};

/** Builds a JSON Response with the given status and merged headers. */
export function json(body: unknown, status = 200, extraHeaders?: Record<string, string>): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...JSON_HEADERS, ...extraHeaders },
  });
}

/** Builds a JSON error Response `{ error: code, message }`. */
export function error(status: number, code: string, message?: string): Response {
  return json({ error: code, message: message ?? code }, status);
}

/** True when the method is not allowed; returns a 405 with an Allow header. */
export function methodNotAllowed(allowed: string[]): Response {
  return json({ error: 'method_not_allowed' }, 405, { allow: allowed.join(', ') });
}
