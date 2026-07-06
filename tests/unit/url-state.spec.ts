import { describe, it, expect } from 'vitest';

import {
  buildShareUrl,
  decodeUrlState,
  decodeViewpoint,
  encodeUrlState,
  encodeViewpoint,
  HIDDEN_IDS_CAP,
  isAllowedModelUrl,
  URL_PARAM_MODEL,
  URL_PARAM_VIEWPOINT,
  VERCEL_BLOB_HOST_SUFFIX,
  type UrlViewpointState,
} from '../../src/core/url-state';

const fullViewpoint: UrlViewpointState = {
  camera: {
    position: { x: 12.3456789, y: -4.5, z: 88.1 },
    target: { x: 0, y: 1.25, z: 0 },
    projection: 'Orthographic',
    mode: 'Plan',
  },
  clippingPlanes: [
    { normal: { x: 0, y: 1, z: 0 }, origin: { x: 0, y: 3.5, z: 0 } },
    { normal: { x: 1, y: 0, z: 0 }, origin: { x: 2, y: 0, z: 0 } },
  ],
  visualStyle: 'pen',
  xray: true,
  edges: true,
  hidden: {
    'building.frag': { count: 3, ids: [10, 20, 30] },
  },
};

describe('url-state · viewpoint round-trip', () => {
  it('round-trips a full viewpoint (coords rounded to 4 decimals)', () => {
    const encoded = encodeViewpoint(fullViewpoint);
    expect(encoded).not.toBe('');
    const decoded = decodeViewpoint(encoded);
    expect(decoded).not.toBeNull();
    expect(decoded!.camera).toEqual({
      position: { x: 12.3457, y: -4.5, z: 88.1 }, // 12.3456789 → 12.3457
      target: { x: 0, y: 1.25, z: 0 },
      projection: 'Orthographic',
      mode: 'Plan',
    });
    expect(decoded!.clippingPlanes).toEqual(fullViewpoint.clippingPlanes);
    expect(decoded!.visualStyle).toBe('pen');
    expect(decoded!.xray).toBe(true);
    expect(decoded!.edges).toBe(true);
    expect(decoded!.hidden).toEqual({ 'building.frag': { count: 3, ids: [10, 20, 30] } });
  });

  it('round-trips a camera-only viewpoint (partial payloads survive)', () => {
    const vp: UrlViewpointState = {
      camera: {
        position: { x: 1, y: 2, z: 3 },
        target: { x: 0, y: 0, z: 0 },
        projection: 'Perspective',
        mode: 'Orbit',
      },
    };
    const decoded = decodeViewpoint(encodeViewpoint(vp));
    expect(decoded).toEqual(vp);
    expect(decoded!.clippingPlanes).toBeUndefined();
    expect(decoded!.xray).toBeUndefined();
  });

  it('omits falsy toggles so they do not appear on decode', () => {
    const vp: UrlViewpointState = {
      camera: { position: { x: 0, y: 0, z: 0 }, target: { x: 1, y: 1, z: 1 }, projection: 'Perspective', mode: 'Orbit' },
      xray: false,
      edges: false,
    };
    const decoded = decodeViewpoint(encodeViewpoint(vp))!;
    expect(decoded.xray).toBeUndefined();
    expect(decoded.edges).toBeUndefined();
  });

  it('returns empty string for an empty viewpoint', () => {
    expect(encodeViewpoint({})).toBe('');
  });

  it('caps the hidden-id sample at HIDDEN_IDS_CAP but keeps the true count', () => {
    const ids = Array.from({ length: HIDDEN_IDS_CAP + 50 }, (_, i) => i);
    const vp: UrlViewpointState = { hidden: { m: { count: 12345, ids } } };
    const decoded = decodeViewpoint(encodeViewpoint(vp))!;
    expect(decoded.hidden!.m.count).toBe(12345);
    expect(decoded.hidden!.m.ids).toHaveLength(HIDDEN_IDS_CAP);
    expect(decoded.hidden!.m.ids[0]).toBe(0);
    expect(decoded.hidden!.m.ids[HIDDEN_IDS_CAP - 1]).toBe(HIDDEN_IDS_CAP - 1);
  });

  it('drops hidden entries with a zero/negative count', () => {
    const vp: UrlViewpointState = { hidden: { m: { count: 0, ids: [1, 2] } }, xray: true };
    const decoded = decodeViewpoint(encodeViewpoint(vp))!;
    expect(decoded.hidden).toBeUndefined();
    expect(decoded.xray).toBe(true);
  });
});

describe('url-state · defensive decoding (untrusted input)', () => {
  it('returns null for empty / nullish hashes', () => {
    expect(decodeViewpoint('')).toBeNull();
    expect(decodeViewpoint(null)).toBeNull();
    expect(decodeViewpoint(undefined)).toBeNull();
  });

  it('returns null for non-base64 / non-JSON garbage', () => {
    expect(decodeViewpoint('!!!not-base64!!!')).toBeNull();
    expect(decodeViewpoint('bm90LWpzb24')).toBeNull(); // base64url of "not-json"
  });

  it('drops a malformed camera (bad coordinate array) but keeps valid fields', () => {
    // Hand-crafted wire object with a broken camera position and a good xray flag.
    const encoded = encodeViewpoint({ xray: true });
    const decoded = decodeViewpoint(encoded)!;
    expect(decoded.camera).toBeUndefined();
    expect(decoded.xray).toBe(true);
  });

  it('falls back to defaults for out-of-range projection / mode indices', () => {
    // Manually build a hash whose camera indices are out of range.
    const wire = { c: { p: [0, 0, 0], t: [1, 1, 1], o: 99, m: 42 } };
    const b64 = Buffer.from(JSON.stringify(wire), 'utf-8')
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');
    const decoded = decodeViewpoint(b64)!;
    expect(decoded.camera!.projection).toBe('Perspective');
    expect(decoded.camera!.mode).toBe('Orbit');
  });
});

describe('url-state · query string codec', () => {
  it('encodes and decodes ?m=&vp= with a Blob-CDN-style URL', () => {
    const modelUrl = 'https://cdn.example.com/blob/abc123.frag?token=xyz&v=1';
    const query = encodeUrlState({ modelUrl, viewpoint: fullViewpoint });
    expect(query.startsWith('?')).toBe(true);
    expect(query).toContain(`${URL_PARAM_MODEL}=`);
    expect(query).toContain(`${URL_PARAM_VIEWPOINT}=`);

    const decoded = decodeUrlState(query);
    expect(decoded.modelUrl).toBe(modelUrl); // query string in the URL survives
    expect(decoded.viewpoint?.camera?.projection).toBe('Orthographic');
  });

  it('handles model-only links (no viewpoint)', () => {
    const query = encodeUrlState({ modelUrl: 'https://cdn/x.frag' });
    expect(query).toBe(`?${URL_PARAM_MODEL}=${encodeURIComponent('https://cdn/x.frag')}`);
    const decoded = decodeUrlState(query);
    expect(decoded.modelUrl).toBe('https://cdn/x.frag');
    expect(decoded.viewpoint).toBeUndefined();
  });

  it('returns empty query for empty state', () => {
    expect(encodeUrlState({})).toBe('');
  });

  it('tolerates a search string without a leading ? and a bad vp', () => {
    const decoded = decodeUrlState('m=https%3A%2F%2Fcdn%2Fx.frag&vp=garbage!!');
    expect(decoded.modelUrl).toBe('https://cdn/x.frag');
    expect(decoded.viewpoint).toBeUndefined();
  });

  it('ignores a blank model param', () => {
    expect(decodeUrlState('?m=&vp=').modelUrl).toBeUndefined();
  });
});

describe('url-state · buildShareUrl', () => {
  it('appends the query to a clean base URL', () => {
    const url = buildShareUrl('https://btc-ifc-viewer-2.vercel.app/embed.html', {
      modelUrl: 'https://cdn/x.frag',
    });
    expect(url).toBe('https://btc-ifc-viewer-2.vercel.app/embed.html?m=' + encodeURIComponent('https://cdn/x.frag'));
  });

  it('strips an existing query/hash on the base before appending', () => {
    const url = buildShareUrl('https://app.example/?old=1#frag', { modelUrl: 'https://cdn/x.frag' });
    expect(url).toBe('https://app.example/?m=' + encodeURIComponent('https://cdn/x.frag'));
  });

  it('returns the bare base when there is nothing to encode', () => {
    expect(buildShareUrl('https://app.example/', {})).toBe('https://app.example/');
  });
});

describe('url-state · isAllowedModelUrl (S8 model-URL allowlist)', () => {
  const SELF = 'https://btc-ifc-viewer-2.vercel.app';

  it('allows the Vercel Blob CDN host', () => {
    expect(
      isAllowedModelUrl('https://abc123.public.blob.vercel-storage.com/frags/x.frag', { selfOrigin: SELF }),
    ).toBe(true);
    expect(VERCEL_BLOB_HOST_SUFFIX).toBe('.public.blob.vercel-storage.com');
  });

  it('allows same-origin model URLs', () => {
    expect(isAllowedModelUrl(`${SELF}/fixtures/x.frag`, { selfOrigin: SELF })).toBe(true);
  });

  it('allows configured extra origins (e.g. a custom Blob domain / BTC_EMBED_ORIGIN)', () => {
    expect(
      isAllowedModelUrl('https://cdn.mycorp.example/x.frag', {
        selfOrigin: SELF,
        allowedOrigins: ['https://cdn.mycorp.example'],
      }),
    ).toBe(true);
  });

  it('REJECTS an arbitrary foreign host (phishing / content-injection guard)', () => {
    expect(isAllowedModelUrl('https://evil.example.com/x.frag', { selfOrigin: SELF })).toBe(false);
  });

  it('REJECTS non-http(s) schemes (javascript:, data:)', () => {
    expect(isAllowedModelUrl('javascript:alert(1)', { selfOrigin: SELF })).toBe(false);
    expect(isAllowedModelUrl('data:text/html,<script>alert(1)</script>', { selfOrigin: SELF })).toBe(false);
  });

  it('REJECTS a malformed URL', () => {
    expect(isAllowedModelUrl('not a url', { selfOrigin: SELF })).toBe(false);
  });

  it('can disable the Vercel Blob default', () => {
    expect(
      isAllowedModelUrl('https://abc.public.blob.vercel-storage.com/x.frag', {
        selfOrigin: SELF,
        allowVercelBlob: false,
      }),
    ).toBe(false);
  });

  it('does not allow a look-alike host that merely contains the suffix as a substring', () => {
    // Suffix must be a real host suffix, not appear mid-host.
    expect(
      isAllowedModelUrl('https://public.blob.vercel-storage.com.evil.com/x.frag', { selfOrigin: SELF }),
    ).toBe(false);
  });
});
