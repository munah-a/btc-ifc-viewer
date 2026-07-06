import { afterEach, describe, expect, it } from 'vitest';

import { clientIp, generateToken, ownerIdFromRequest, sha256Hex, verifyToken } from '../../api/_lib/http';

function req(headers: Record<string, string>): Request {
  return new Request('https://x/api/uploads', { method: 'POST', headers });
}

afterEach(() => {
  delete process.env.VERCEL;
  delete process.env.BTC_OWNER_SALT;
});

describe('http · clientIp (S2 anti-spoof)', () => {
  it('prefers the platform-trusted x-real-ip over any x-forwarded-for', () => {
    expect(clientIp(req({ 'x-real-ip': '203.0.113.1', 'x-forwarded-for': '10.0.0.1, 8.8.8.8' }))).toBe('203.0.113.1');
  });

  it('uses the RIGHT-most x-forwarded-for hop, never the client-supplied left', () => {
    expect(clientIp(req({ 'x-forwarded-for': '10.0.0.1, 203.0.113.9' }))).toBe('203.0.113.9');
    // A single value is both ends.
    expect(clientIp(req({ 'x-forwarded-for': '203.0.113.7' }))).toBe('203.0.113.7');
  });

  it('honours x-vercel-forwarded-for (right-most) when present', () => {
    expect(clientIp(req({ 'x-vercel-forwarded-for': '1.1.1.1, 203.0.113.50' }))).toBe('203.0.113.50');
  });

  it('returns null when no trusted header is present', () => {
    expect(clientIp(req({}))).toBeNull();
  });
});

describe('http · ownerIdFromRequest (S2/S7)', () => {
  it('a spoofed left-most XFF does NOT change the owner id (same trusted right hop)', async () => {
    const a = await ownerIdFromRequest(req({ 'x-forwarded-for': '10.0.0.1, 203.0.113.9' }));
    const b = await ownerIdFromRequest(req({ 'x-forwarded-for': '99.99.99.99, 203.0.113.9' }));
    expect(a).toBe(b); // same trusted IP → same bucket, spoof ignored
    const different = await ownerIdFromRequest(req({ 'x-forwarded-for': '203.0.113.10' }));
    expect(different).not.toBe(a);
  });

  it('S7: throws in production when BTC_OWNER_SALT is unset', async () => {
    process.env.VERCEL = '1';
    delete process.env.BTC_OWNER_SALT;
    await expect(ownerIdFromRequest(req({ 'x-real-ip': '1.2.3.4' }))).rejects.toThrow(/BTC_OWNER_SALT/);
  });

  it('S7: uses the configured salt in production (no throw)', async () => {
    process.env.VERCEL = '1';
    process.env.BTC_OWNER_SALT = 'a-real-salt';
    await expect(ownerIdFromRequest(req({ 'x-real-ip': '1.2.3.4' }))).resolves.toMatch(/^anon_/);
  });

  it('allows the dev fallback salt off-platform', async () => {
    delete process.env.VERCEL;
    delete process.env.BTC_OWNER_SALT;
    await expect(ownerIdFromRequest(req({ 'x-real-ip': '1.2.3.4' }))).resolves.toMatch(/^anon_/);
  });
});

describe('http · verifyToken (S6 constant-time)', () => {
  it('accepts the correct token and rejects a wrong one', async () => {
    const token = generateToken(32);
    const hash = await sha256Hex(token);
    expect(await verifyToken(token, hash)).toBe(true);
    expect(await verifyToken('wrong', hash)).toBe(false);
  });

  it('rejects empty/mismatched-length inputs without throwing', async () => {
    expect(await verifyToken('', 'abc')).toBe(false);
    expect(await verifyToken('abc', '')).toBe(false);
    // A stored hash of the wrong length must not throw timingSafeEqual.
    expect(await verifyToken('token', 'short')).toBe(false);
  });

  it('is deterministic across many correct/incorrect trials', async () => {
    const token = generateToken(32);
    const hash = await sha256Hex(token);
    for (let i = 0; i < 20; i += 1) {
      expect(await verifyToken(token, hash)).toBe(true);
      expect(await verifyToken(generateToken(32), hash)).toBe(false);
    }
  });
});
