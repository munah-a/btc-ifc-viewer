import { beforeEach, describe, expect, it } from 'vitest';

import { ANON_DEFAULTS, RATE_LIMIT_MAX } from '../../api/_lib/entitlements';
import {
  handleCronCleanup,
  handleEmbedMeta,
  handleUpload,
  isValidId,
  sanitizeFileName,
  type HostConfig,
} from '../../api/_lib/hosting';
import { InMemoryStorage } from '../../api/_lib/storage';
import { ownerIdFromRequest } from '../../api/_lib/http';

const HOST: HostConfig = { origin: 'https://btc-ifc-viewer-2.vercel.app' };

/** A controllable clock (ms) for TTL/rate-window tests. */
function makeClock(start = 1_700_000_000_000): { now: () => number; advance: (ms: number) => void } {
  let t = start;
  return { now: () => t, advance: (ms) => (t += ms) };
}

function uploadRequest(bytes: Uint8Array, headers: Record<string, string> = {}): Request {
  return new Request('https://btc-ifc-viewer-2.vercel.app/api/uploads', {
    method: 'POST',
    headers: {
      'content-type': 'application/octet-stream',
      'x-file-name': 'tower.frag',
      'x-forwarded-for': '203.0.113.7',
      ...headers,
    },
    // Send the underlying ArrayBuffer (a valid BodyInit under strict DOM types).
    body: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
  });
}

const FRAG = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);

describe('api · upload happy path', () => {
  it('stores bytes, returns embed/viewer/delete/expiry, applies long blob cache', async () => {
    const clock = makeClock();
    const storage = new InMemoryStorage(clock.now);
    const res = await handleUpload(uploadRequest(FRAG), storage, { clock: clock.now, host: HOST });
    expect(res.status).toBe(201);
    const body = (await res.json()) as {
      id: string;
      embedUrl: string;
      viewerUrl: string;
      fragUrl: string;
      deleteToken: string;
      expiresAt: string;
    };

    expect(isValidId(body.id)).toBe(true);
    expect(body.embedUrl).toContain('https://btc-ifc-viewer-2.vercel.app/embed.html?');
    expect(body.embedUrl).toContain('m=');
    expect(body.viewerUrl).toContain('https://btc-ifc-viewer-2.vercel.app/?m=');
    expect(body.deleteToken).toBeTruthy();
    // 7-day TTL from the anon default.
    const expiresMs = new Date(body.expiresAt).getTime();
    expect(expiresMs - clock.now()).toBe(ANON_DEFAULTS.ttlDays * 24 * 3600 * 1000);

    // C2: the stored blob holds the EXACT bytes uploaded, untouched.
    const blobPath = `frags/${body.id}.frag`;
    expect(storage.hasBlob(blobPath)).toBe(true);
    expect([...(storage.getBlobBytes(blobPath) ?? [])]).toEqual([...FRAG]);
    // S4: blob cache max-age = the TTL (7 days), NOT 1 year — so expired/deleted
    // content also falls out of the CDN cache and can't be fetched past its life.
    expect(storage.getBlobMaxAge(blobPath)).toBe(ANON_DEFAULTS.ttlDays * 24 * 3600);
  });

  it('never returns ownerId or the delete-token hash to the client', async () => {
    const storage = new InMemoryStorage();
    const res = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    const raw = await res.text();
    expect(raw).not.toContain('ownerId');
    expect(raw).not.toContain('deleteTokenHash');
  });
});

describe('api · upload validation (defensive)', () => {
  let storage: InMemoryStorage;
  beforeEach(() => {
    storage = new InMemoryStorage();
  });

  it('rejects non-POST', async () => {
    const req = new Request('https://x/api/uploads', { method: 'GET' });
    const res = await handleUpload(req, storage, { host: HOST });
    expect(res.status).toBe(405);
  });

  it('rejects a wrong content-type', async () => {
    const req = uploadRequest(FRAG, { 'content-type': 'application/json' });
    const res = await handleUpload(req, storage, { host: HOST });
    expect(res.status).toBe(415);
  });

  it('rejects an empty body', async () => {
    const res = await handleUpload(uploadRequest(new Uint8Array(0)), storage, { host: HOST });
    expect(res.status).toBe(400);
    expect(((await res.json()) as { error: string }).error).toBe('empty_body');
  });

  it('rejects a payload over the size cap (413)', async () => {
    const tooBig = new Uint8Array(ANON_DEFAULTS.maxUploadBytes + 1);
    const res = await handleUpload(uploadRequest(tooBig), storage, { host: HOST });
    expect(res.status).toBe(413);
    expect(((await res.json()) as { error: string }).error).toBe('payload_too_large');
  });

  it('rejects an oversize declared Content-Length before buffering the body (413)', async () => {
    const req = uploadRequest(FRAG, {
      'content-length': String(ANON_DEFAULTS.maxUploadBytes + 1),
    });
    const res = await handleUpload(req, storage, { host: HOST });
    expect(res.status).toBe(413);
  });
});

describe('api · per-owner quota (C4 seam)', () => {
  it(`allows exactly maxActiveUploads then returns 409`, async () => {
    const storage = new InMemoryStorage();
    // Same IP → same owner surrogate → same quota bucket.
    for (let i = 0; i < ANON_DEFAULTS.maxActiveUploads; i += 1) {
      const res = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
      expect(res.status).toBe(201);
    }
    const overflow = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    expect(overflow.status).toBe(409);
    expect(((await overflow.json()) as { error: string }).error).toBe('quota_exceeded');
  });

  it('frees a quota slot after the owner deletes an upload', async () => {
    const storage = new InMemoryStorage();
    const created: { id: string; deleteToken: string }[] = [];
    for (let i = 0; i < ANON_DEFAULTS.maxActiveUploads; i += 1) {
      const res = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
      created.push((await res.json()) as { id: string; deleteToken: string });
    }
    // Delete one → a slot frees → next upload succeeds.
    const del = await handleEmbedMeta(deleteRequest(created[0].id, created[0].deleteToken), created[0].id, storage);
    expect(del.status).toBe(200);
    const again = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    expect(again.status).toBe(201);
  });

  it('separates quota buckets by client IP', async () => {
    const storage = new InMemoryStorage();
    for (let i = 0; i < ANON_DEFAULTS.maxActiveUploads; i += 1) {
      await handleUpload(uploadRequest(FRAG, { 'x-forwarded-for': '198.51.100.1' }), storage, { host: HOST });
    }
    // A DIFFERENT IP has its own full quota.
    const other = await handleUpload(uploadRequest(FRAG, { 'x-forwarded-for': '198.51.100.2' }), storage, { host: HOST });
    expect(other.status).toBe(201);
  });

  it('S2: a spoofed left-most X-Forwarded-For does NOT mint a fresh quota bucket', async () => {
    const storage = new InMemoryStorage();
    // Attacker prepends a random client value each request, but the platform
    // appends the true connecting IP as the RIGHT-most hop (constant here).
    const spoofed = (n: number): Record<string, string> => ({
      'x-forwarded-for': `10.0.0.${n}, 203.0.113.9`, // left=spoof, right=real
    });
    for (let i = 0; i < ANON_DEFAULTS.maxActiveUploads; i += 1) {
      const res = await handleUpload(uploadRequest(FRAG, spoofed(i)), storage, { host: HOST });
      expect(res.status).toBe(201);
    }
    // The (maxActive+1)th, still spoofing a NEW left-most value, must still be
    // capped — the trusted right-most hop is unchanged, so it's the same owner.
    const overflow = await handleUpload(uploadRequest(FRAG, spoofed(999)), storage, { host: HOST });
    expect(overflow.status).toBe(409);
  });

  it('S2: x-real-ip (platform-trusted) wins over a client x-forwarded-for', async () => {
    const storage = new InMemoryStorage();
    // Fill the quota using x-real-ip as the trusted signal.
    for (let i = 0; i < ANON_DEFAULTS.maxActiveUploads; i += 1) {
      await handleUpload(
        uploadRequest(FRAG, { 'x-real-ip': '198.51.100.50', 'x-forwarded-for': `1.2.3.${i}` }),
        storage,
        { host: HOST },
      );
    }
    // A new spoofed XFF but the SAME x-real-ip → same bucket → rejected.
    const res = await handleUpload(
      uploadRequest(FRAG, { 'x-real-ip': '198.51.100.50', 'x-forwarded-for': '9.9.9.9' }),
      storage,
      { host: HOST },
    );
    expect(res.status).toBe(409);
  });

  it('S1: N CONCURRENT uploads store at most maxActiveUploads (no TOCTOU race)', async () => {
    const storage = new InMemoryStorage();
    const N = ANON_DEFAULTS.maxActiveUploads + 5; // fire well over the cap at once
    // All in flight before any awaits resolve — the old count-then-track would
    // let every one pass the check. reserveOwnerSlot must cap it atomically.
    const results = await Promise.all(
      Array.from({ length: N }, () => handleUpload(uploadRequest(FRAG), storage, { host: HOST })),
    );
    const created = results.filter((r) => r.status === 201).length;
    const rejected = results.filter((r) => r.status === 409).length;
    expect(created).toBe(ANON_DEFAULTS.maxActiveUploads);
    expect(rejected).toBe(N - ANON_DEFAULTS.maxActiveUploads);
    // And the store actually holds ≤ maxActive for the owner.
    const owner = await ownerIdFromRequest(uploadRequest(FRAG));
    expect(await storage.countActiveByOwner(owner)).toBeLessThanOrEqual(ANON_DEFAULTS.maxActiveUploads);
  });
});

describe('api · rate limiting', () => {
  it('returns 429 once the per-window upload limit is exceeded', async () => {
    const clock = makeClock();
    const storage = new InMemoryStorage(clock.now);
    // A generous quota keeps the quota check from masking the rate check.
    // Fire RATE_LIMIT_MAX uploads (each within quota by deleting immediately is
    // overkill — instead push the limit purely: rate limit is checked first).
    let last: Response | null = null;
    for (let i = 0; i < RATE_LIMIT_MAX + 1; i += 1) {
      last = await handleUpload(uploadRequest(FRAG), storage, { clock: clock.now, host: HOST });
      // Delete to keep quota free so we actually reach the rate ceiling.
      if (last.status === 201) {
        const body = (await last.clone().json()) as { id: string; deleteToken: string };
        await handleEmbedMeta(deleteRequest(body.id, body.deleteToken), body.id, storage);
      }
    }
    expect(last!.status).toBe(429);
    expect(((await last!.json()) as { error: string }).error).toBe('rate_limited');
  });
});

describe('api · GET /e/:id', () => {
  it('returns only public metadata for a live upload', async () => {
    const storage = new InMemoryStorage();
    const up = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    const { id, fragUrl } = (await up.json()) as { id: string; fragUrl: string };

    const res = await handleEmbedMeta(getRequest(id), id, storage);
    expect(res.status).toBe(200);
    const body = (await res.json()) as Record<string, unknown>;
    expect(body.id).toBe(id);
    expect(body.fragUrl).toBe(fragUrl);
    expect(body.expiresAt).toBeTruthy();
    expect(body).not.toHaveProperty('ownerId');
    expect(body).not.toHaveProperty('deleteTokenHash');
    expect(body).not.toHaveProperty('blobPath');
  });

  it('404s for an unknown id and 400s for a malformed id', async () => {
    const storage = new InMemoryStorage();
    expect((await handleEmbedMeta(getRequest('doesnotexist99'), 'doesnotexist99', storage)).status).toBe(404);
    expect((await handleEmbedMeta(getRequest('bad id!'), 'bad id!', storage)).status).toBe(400);
  });
});

describe('api · TTL expiry', () => {
  it('GET /e/:id 404s once the record TTL has elapsed', async () => {
    const clock = makeClock();
    const storage = new InMemoryStorage(clock.now);
    const up = await handleUpload(uploadRequest(FRAG), storage, { clock: clock.now, host: HOST });
    const { id } = (await up.json()) as { id: string };

    // Just before expiry: still live.
    clock.advance(ANON_DEFAULTS.ttlDays * 24 * 3600 * 1000 - 1000);
    expect((await handleEmbedMeta(getRequest(id), id, storage)).status).toBe(200);

    // Past expiry: gone.
    clock.advance(2000);
    expect((await handleEmbedMeta(getRequest(id), id, storage)).status).toBe(404);
  });
});

describe('api · DELETE /e/:id (delete-token auth)', () => {
  it('deletes blob + meta with a valid token', async () => {
    const storage = new InMemoryStorage();
    const up = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    const { id, deleteToken } = (await up.json()) as { id: string; deleteToken: string };
    const blobPath = `frags/${id}.frag`;
    expect(storage.hasBlob(blobPath)).toBe(true);

    const res = await handleEmbedMeta(deleteRequest(id, deleteToken), id, storage);
    expect(res.status).toBe(200);
    expect(((await res.json()) as { deleted: boolean }).deleted).toBe(true);
    expect(storage.hasBlob(blobPath)).toBe(false);
    expect((await handleEmbedMeta(getRequest(id), id, storage)).status).toBe(404);
  });

  it('rejects a missing token (401) and a wrong token (403), leaving the upload intact', async () => {
    const storage = new InMemoryStorage();
    const up = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    const { id } = (await up.json()) as { id: string };

    const noToken = new Request(`https://x/api/e/${id}`, { method: 'DELETE' });
    expect((await handleEmbedMeta(noToken, id, storage)).status).toBe(401);

    const wrong = await handleEmbedMeta(deleteRequest(id, 'not-the-real-token'), id, storage);
    expect(wrong.status).toBe(403);

    // Still there.
    expect(storage.hasBlob(`frags/${id}.frag`)).toBe(true);
    expect((await handleEmbedMeta(getRequest(id), id, storage)).status).toBe(200);
  });

  it('accepts the token via Authorization: Bearer as well as x-delete-token', async () => {
    const storage = new InMemoryStorage();
    const up = await handleUpload(uploadRequest(FRAG), storage, { host: HOST });
    const { id, deleteToken } = (await up.json()) as { id: string; deleteToken: string };
    const bearer = new Request(`https://x/api/e/${id}`, {
      method: 'DELETE',
      headers: { authorization: `Bearer ${deleteToken}` },
    });
    expect((await handleEmbedMeta(bearer, id, storage)).status).toBe(200);
  });
});

describe('api · cron cleanup', () => {
  it('sweeps expired uploads (blob + meta) and reports the count', async () => {
    const clock = makeClock();
    const storage = new InMemoryStorage(clock.now);
    const up = await handleUpload(uploadRequest(FRAG), storage, { clock: clock.now, host: HOST });
    const { id } = (await up.json()) as { id: string };
    const blobPath = `frags/${id}.frag`;

    // Advance past TTL, then run cleanup.
    clock.advance(ANON_DEFAULTS.ttlDays * 24 * 3600 * 1000 + 1000);
    const res = await handleCronCleanup(cronRequest(), storage, { clock: clock.now });
    expect(res.status).toBe(200);
    expect(((await res.json()) as { deleted: number }).deleted).toBe(1);
    expect(storage.hasBlob(blobPath)).toBe(false);
  });

  it('rejects an invalid cron secret when CRON_SECRET is set', async () => {
    const prev = process.env.CRON_SECRET;
    process.env.CRON_SECRET = 'top-secret';
    try {
      const storage = new InMemoryStorage();
      const bad = new Request('https://x/api/cron-cleanup', {
        method: 'GET',
        headers: { authorization: 'Bearer wrong' },
      });
      expect((await handleCronCleanup(bad, storage)).status).toBe(401);

      const good = new Request('https://x/api/cron-cleanup', {
        method: 'GET',
        headers: { authorization: 'Bearer top-secret' },
      });
      expect((await handleCronCleanup(good, storage)).status).toBe(200);
    } finally {
      if (prev === undefined) delete process.env.CRON_SECRET;
      else process.env.CRON_SECRET = prev;
    }
  });

  it('S5: FAILS CLOSED in production when CRON_SECRET is unset (500, no sweep)', async () => {
    const prevSecret = process.env.CRON_SECRET;
    const prevVercel = process.env.VERCEL;
    delete process.env.CRON_SECRET;
    process.env.VERCEL = '1'; // emulate a real Vercel deployment
    try {
      const storage = new InMemoryStorage();
      const req = new Request('https://x/api/cron-cleanup', { method: 'GET' });
      const res = await handleCronCleanup(req, storage);
      expect(res.status).toBe(500);
      expect(((await res.json()) as { error: string }).error).toBe('misconfigured');
    } finally {
      if (prevSecret === undefined) delete process.env.CRON_SECRET;
      else process.env.CRON_SECRET = prevSecret;
      if (prevVercel === undefined) delete process.env.VERCEL;
      else process.env.VERCEL = prevVercel;
    }
  });

  it('allows the sweep locally (no VERCEL, no secret) so tests can exercise it', async () => {
    const prevSecret = process.env.CRON_SECRET;
    const prevVercel = process.env.VERCEL;
    delete process.env.CRON_SECRET;
    delete process.env.VERCEL;
    try {
      const storage = new InMemoryStorage();
      const req = new Request('https://x/api/cron-cleanup', { method: 'GET' });
      expect((await handleCronCleanup(req, storage)).status).toBe(200);
    } finally {
      if (prevSecret !== undefined) process.env.CRON_SECRET = prevSecret;
      if (prevVercel !== undefined) process.env.VERCEL = prevVercel;
    }
  });
});

describe('api · sanitizeFileName', () => {
  it('strips path separators and control chars, caps length, keeps a sane default', () => {
    const NUL = String.fromCharCode(0);
    const UNIT_SEP = String.fromCharCode(0x1f);
    expect(sanitizeFileName('../../etc/passwd')).toBe('.._.._etc_passwd');
    expect(sanitizeFileName('a' + String.fromCharCode(92) + 'b' + String.fromCharCode(92) + 'c.frag')).toBe('a_b_c.frag');
    expect(sanitizeFileName('na' + NUL + UNIT_SEP + 'me.frag')).toBe('name.frag'); // controls stripped
    expect(sanitizeFileName('  Level 1.frag  ')).toBe('Level 1.frag'); // interior space kept; ends trimmed
    expect(sanitizeFileName('   ')).toBe('model.frag');
    expect(sanitizeFileName('x'.repeat(500)).length).toBe(120);
    expect(sanitizeFileName('Türme.frag')).toBe('Türme.frag'); // non-ASCII kept
  });
});

// ── request builders ──
function getRequest(id: string): Request {
  return new Request(`https://btc-ifc-viewer-2.vercel.app/api/e/${id}`, { method: 'GET' });
}
function deleteRequest(id: string, token: string): Request {
  return new Request(`https://btc-ifc-viewer-2.vercel.app/api/e/${id}`, {
    method: 'DELETE',
    headers: { 'x-delete-token': token },
  });
}
function cronRequest(): Request {
  return new Request('https://btc-ifc-viewer-2.vercel.app/api/cron-cleanup', { method: 'GET' });
}
