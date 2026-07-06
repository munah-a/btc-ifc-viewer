/**
 * Hosting API business logic (W4.3), separated from the Vercel route wrappers so
 * every path is unit-testable with the in-memory storage adapter and an
 * injectable clock — NO live Vercel storage, NO framework, just
 * `(Request, StorageAdapter) => Response`.
 *
 * CONSTRAINT C2 (server never processes models): these functions validate size,
 * rate-limit and quota, then store/serve OPAQUE bytes + small metadata. They do
 * NOT parse, render or inspect `.frag` contents. The `.frag` is produced by the
 * uploader's browser (IFC→fragments) and fetched by embeds DIRECTLY from the
 * Blob CDN — functions never stream model bytes on the read path (no egress).
 *
 * Endpoints:
 *   POST  /api/uploads      → store a browser-converted .frag, return
 *                             { id, embedUrl, viewerUrl, deleteToken, expiresAt }
 *   GET   /api/e/:id        → { id, fragUrl (direct Blob-CDN URL), expiresAt, … }
 *   DELETE /api/e/:id       → delete blob+meta (requires the delete token)
 *   GET   /api/cron-cleanup → sweep expired uploads (CRON_SECRET-guarded)
 */

import {
  ANON_DEFAULTS,
  RATE_LIMIT_MAX,
  RATE_LIMIT_WINDOW_SECONDS,
  resolveEntitlements,
  ttlSeconds,
  type Entitlements,
} from './entitlements';
import { error, generateId, generateToken, json, methodNotAllowed, ownerIdFromRequest, sha256Hex, verifyToken } from './http';
import type { Clock, StorageAdapter, UploadMeta } from './storage';

/** How embeds/deep-links are built. Config-driven so the host domain is easy to change. */
export interface HostConfig {
  /** Public origin of the deployment, e.g. https://btc-ifc-viewer-2.vercel.app (no trailing slash). */
  origin: string;
}

/** Long cache for immutable `.frag` bytes at Blob put-time (1 year). */
const FRAG_CACHE_MAX_AGE_SECONDS = 31536000;
const FRAG_CONTENT_TYPE = 'application/octet-stream';

/** Resolves the host origin from env (BTC_EMBED_ORIGIN) or a request, defaulting to the Vercel project. */
export function resolveHostConfig(request?: Request): HostConfig {
  const fromEnv = typeof process !== 'undefined' ? process.env.BTC_EMBED_ORIGIN : undefined;
  if (fromEnv) return { origin: fromEnv.replace(/\/+$/, '') };
  if (request) {
    try {
      const url = new URL(request.url);
      // Prefer the forwarded host (Vercel sets it) so embed URLs match the public domain.
      const host = request.headers.get('x-forwarded-host') ?? url.host;
      const proto = request.headers.get('x-forwarded-proto') ?? url.protocol.replace(':', '');
      if (host) return { origin: `${proto}://${host}` };
    } catch {
      // fall through
    }
  }
  return { origin: 'https://btc-ifc-viewer-2.vercel.app' };
}

/** Builds the embed iframe URL for an upload id. */
export function embedUrl(host: HostConfig, id: string, fragUrl: string): string {
  // The embed loads by model URL (?m=), so it works even if the meta record is
  // gone-but-blob-lives edge case; the id is carried for "open in viewer".
  const params = new URLSearchParams({ m: fragUrl, id });
  return `${host.origin}/embed.html?${params.toString()}`;
}

/** Builds the full-app "open in viewer" URL for an upload. */
export function viewerUrl(host: HostConfig, fragUrl: string): string {
  return `${host.origin}/?m=${encodeURIComponent(fragUrl)}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// POST /api/uploads
// ─────────────────────────────────────────────────────────────────────────────

export async function handleUpload(
  request: Request,
  storage: StorageAdapter,
  opts: { clock?: Clock; host?: HostConfig } = {},
): Promise<Response> {
  if (request.method !== 'POST') return methodNotAllowed(['POST']);

  const now = opts.clock ?? (() => Date.now());
  const host = opts.host ?? resolveHostConfig(request);
  const entitlements: Entitlements = resolveEntitlements();

  const ownerId = await ownerIdFromRequest(request);

  // 1) Rate limit the endpoint per owner surrogate (defends against abuse).
  const rate = await storage.rateLimit(`upload:${ownerId}`, RATE_LIMIT_MAX, RATE_LIMIT_WINDOW_SECONDS);
  if (!rate.allowed) {
    return error(429, 'rate_limited', `Too many uploads. Limit is ${rate.limit} per ${RATE_LIMIT_WINDOW_SECONDS / 60} minutes.`);
  }

  // 2) Read the raw bytes. The `.frag` is sent as the request body (octet-stream)
  //    with the file name in a header — we NEVER parse the bytes (C2).
  //    ASSUMPTION (PO): the entitlement size cap (50 MB) can exceed Vercel's
  //    default function request-body limit (~4.5 MB). Before production, either
  //    raise the body limit for api/uploads OR switch large uploads to the
  //    @vercel/blob CLIENT-upload flow (browser → Blob directly via a token
  //    endpoint), which also avoids buffering big bodies in the function. The
  //    early Content-Length check below rejects oversize bodies before buffering.
  const declaredLength = Number(request.headers.get('content-length') ?? 0);
  if (declaredLength > entitlements.maxUploadBytes) {
    const maxMb = Math.floor(entitlements.maxUploadBytes / (1024 * 1024));
    return error(413, 'payload_too_large', `Upload exceeds the ${maxMb} MB limit for the ${entitlements.tier} tier.`);
  }

  const contentType = request.headers.get('content-type') ?? '';
  if (!contentType.includes('application/octet-stream')) {
    return error(415, 'unsupported_media_type', 'Upload the converted .frag as application/octet-stream.');
  }

  const buffer = await request.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  // 3) Size cap from entitlements (before any storage write).
  if (bytes.byteLength === 0) {
    return error(400, 'empty_body', 'No .frag bytes received.');
  }
  if (bytes.byteLength > entitlements.maxUploadBytes) {
    const maxMb = Math.floor(entitlements.maxUploadBytes / (1024 * 1024));
    return error(413, 'payload_too_large', `Upload exceeds the ${maxMb} MB limit for the ${entitlements.tier} tier.`);
  }

  // 4) Per-owner active-upload quota (C4). Checked AFTER size so a rejected
  //    oversize upload doesn't consume a quota slot check needlessly.
  const activeCount = await storage.countActiveByOwner(ownerId);
  if (activeCount >= entitlements.maxActiveUploads) {
    return error(
      409,
      'quota_exceeded',
      `You have ${activeCount}/${entitlements.maxActiveUploads} active embeds. Delete one before creating another.`,
    );
  }

  // 5) Sanitize the display file name (never trusted for logic/paths).
  const rawName = request.headers.get('x-file-name') ?? 'model.frag';
  const fileName = sanitizeFileName(rawName);

  // 6) Store bytes → Blob CDN, with long immutable cache at put-time.
  const id = generateId();
  const blobPath = `frags/${id}.frag`;
  const { url: fragUrl, path: storedPath } = await storage.putFrag(blobPath, bytes, {
    contentType: FRAG_CONTENT_TYPE,
    cacheControlMaxAgeSeconds: FRAG_CACHE_MAX_AGE_SECONDS,
  });

  // 7) Mint a delete token (returned once); store only its hash.
  const deleteToken = generateToken(32);
  const deleteTokenHash = await sha256Hex(deleteToken);

  const ttl = ttlSeconds(entitlements);
  const createdAtMs = now();
  const expiresAtMs = createdAtMs + ttl * 1000;
  const meta: UploadMeta = {
    id,
    fragUrl,
    blobPath: storedPath,
    fileName,
    sizeBytes: bytes.byteLength,
    createdAt: new Date(createdAtMs).toISOString(),
    expiresAt: new Date(expiresAtMs).toISOString(),
    ownerId,
    tier: entitlements.tier,
    deleteTokenHash,
  };

  await storage.setMeta(meta, ttl);
  await storage.trackOwnerUpload(ownerId, id, ttl);

  return json(
    {
      id,
      embedUrl: embedUrl(host, id, fragUrl),
      viewerUrl: viewerUrl(host, fragUrl),
      fragUrl,
      deleteToken, // shown to the uploader ONCE; not recoverable
      expiresAt: meta.expiresAt,
    },
    201,
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// GET/DELETE /api/e/:id
// ─────────────────────────────────────────────────────────────────────────────

export async function handleEmbedMeta(request: Request, id: string, storage: StorageAdapter): Promise<Response> {
  if (request.method === 'GET') return getEmbedMeta(id, storage);
  if (request.method === 'DELETE') return deleteEmbed(request, id, storage);
  return methodNotAllowed(['GET', 'DELETE']);
}

async function getEmbedMeta(id: string, storage: StorageAdapter): Promise<Response> {
  if (!isValidId(id)) return error(400, 'bad_id', 'Invalid upload id.');
  const meta = await storage.getMeta(id);
  if (!meta) return error(404, 'not_found', 'This embed has expired or does not exist.');
  // Return ONLY what the embed needs; never leak ownerId / deleteTokenHash.
  return json({
    id: meta.id,
    fragUrl: meta.fragUrl,
    fileName: meta.fileName,
    sizeBytes: meta.sizeBytes,
    expiresAt: meta.expiresAt,
  });
}

async function deleteEmbed(request: Request, id: string, storage: StorageAdapter): Promise<Response> {
  if (!isValidId(id)) return error(400, 'bad_id', 'Invalid upload id.');
  const token = deleteTokenFromRequest(request);
  if (!token) return error(401, 'missing_token', 'A delete token is required.');

  const meta = await storage.getMeta(id);
  if (!meta) return error(404, 'not_found', 'This embed has expired or does not exist.');

  const ok = await verifyToken(token, meta.deleteTokenHash);
  if (!ok) return error(403, 'forbidden', 'Invalid delete token.');

  await storage.deleteFrag(meta.blobPath);
  await storage.deleteMeta(id);
  await storage.untrackOwnerUpload(meta.ownerId, id);

  return json({ id, deleted: true });
}

// ─────────────────────────────────────────────────────────────────────────────
// GET /api/cron-cleanup  (CRON_SECRET-guarded)
// ─────────────────────────────────────────────────────────────────────────────

export async function handleCronCleanup(
  request: Request,
  storage: StorageAdapter,
  opts: { clock?: Clock } = {},
): Promise<Response> {
  if (request.method !== 'GET') return methodNotAllowed(['GET']);

  // Vercel Cron sends `Authorization: Bearer $CRON_SECRET`. Reject anything else.
  const secret = typeof process !== 'undefined' ? process.env.CRON_SECRET : undefined;
  if (secret) {
    const auth = request.headers.get('authorization');
    if (auth !== `Bearer ${secret}`) return error(401, 'unauthorized', 'Invalid cron secret.');
  }
  // If no secret is configured (local test), allow — the real deploy MUST set it.

  const now = (opts.clock ?? (() => Date.now()))();
  const expired = await storage.listExpired(now);
  let deleted = 0;
  for (const meta of expired) {
    await storage.deleteFrag(meta.blobPath);
    await storage.deleteMeta(meta.id);
    await storage.untrackOwnerUpload(meta.ownerId, meta.id);
    deleted += 1;
  }
  return json({ sweptAt: new Date(now).toISOString(), deleted });
}

// ── small helpers ──

/** Public ids are our own base64url tokens: [-_A-Za-z0-9], reasonable length. */
export function isValidId(id: string): boolean {
  return typeof id === 'string' && /^[A-Za-z0-9_-]{8,64}$/.test(id);
}

/** Reads the delete token from `Authorization: Bearer` or an `x-delete-token` header. */
function deleteTokenFromRequest(request: Request): string | null {
  const auth = request.headers.get('authorization');
  if (auth?.startsWith('Bearer ')) {
    const t = auth.slice('Bearer '.length).trim();
    if (t) return t;
  }
  const header = request.headers.get('x-delete-token');
  return header?.trim() || null;
}

/** Strips path separators + C0/DEL control chars from a display file name; caps length. */
export function sanitizeFileName(name: string): string {
  let base = name.split('/').join('_').split(String.fromCharCode(92)).join('_');
  // Drop C0 controls (0x00-0x1F) and DEL (0x7F); ASCII-only construction avoids
  // embedding literal control bytes in the source.
  base = base
    .split('')
    .filter((ch) => {
      const code = ch.charCodeAt(0);
      return code > 0x1f && code !== 0x7f;
    })
    .join('')
    .trim();
  const clean = base.slice(0, 120);
  return clean || 'model.frag';
}

// Re-export the anon defaults for callers that want to advertise limits (e.g. oEmbed / share dialog copy).
export { ANON_DEFAULTS };
