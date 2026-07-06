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
 * Web-standard only (Request/Response/crypto.subtle) — runs on the Vercel
 * Node.js runtime and is unit-testable in Node without any Vercel packages.
 */

/** SHA-256 of a string → lowercase hex. */
export async function sha256Hex(input: string): Promise<string> {
  const data = new TextEncoder().encode(input);
  const digest = await crypto.subtle.digest('SHA-256', data);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Derives a non-reversible owner surrogate from the request IP + a server salt.
 * This is the C4 seam value stored as `ownerId`; a real account id replaces it
 * later with no schema change. Falls back to a constant bucket when no IP is
 * available (dev/local), which is safe — it just shares one quota bucket.
 */
export async function ownerIdFromRequest(request: Request): Promise<string> {
  const ip = clientIp(request) ?? 'no-ip';
  const salt = (typeof process !== 'undefined' && process.env.BTC_OWNER_SALT) || 'btc-ifc-viewer-dev-salt';
  const hash = await sha256Hex(`${salt}:${ip}`);
  return `anon_${hash.slice(0, 24)}`;
}

/**
 * Best-effort client IP from proxy headers Vercel sets. `x-forwarded-for` is a
 * comma list (client first). We take only the first hop; it is used solely to
 * bucket quota/rate-limit and is immediately salted+hashed, never stored raw.
 */
export function clientIp(request: Request): string | null {
  const xff = request.headers.get('x-forwarded-for');
  if (xff) {
    const first = xff.split(',')[0]?.trim();
    if (first) return first;
  }
  const real = request.headers.get('x-real-ip');
  return real?.trim() || null;
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
 * Constant-time-ish comparison of a presented delete token against the stored
 * hash. Hashing the presented token first means both sides are fixed-length
 * hex, so the char-by-char compare doesn't leak length and never short-circuits.
 */
export async function verifyToken(presentedToken: string, storedHash: string): Promise<boolean> {
  if (!presentedToken || !storedHash) return false;
  const presentedHash = await sha256Hex(presentedToken);
  if (presentedHash.length !== storedHash.length) return false;
  let diff = 0;
  for (let i = 0; i < presentedHash.length; i += 1) {
    diff |= presentedHash.charCodeAt(i) ^ storedHash.charCodeAt(i);
  }
  return diff === 0;
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
