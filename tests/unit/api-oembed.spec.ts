import { describe, expect, it } from 'vitest';

import { handleOEmbed, isOwnEmbedUrl, buildIframeHtml, type OEmbedResponse } from '../../api/_lib/oembed';
import type { HostConfig } from '../../api/_lib/hosting';

const HOST: HostConfig = { origin: 'https://btc-ifc-viewer-2.vercel.app' };
const EMBED_URL = 'https://btc-ifc-viewer-2.vercel.app/embed.html?m=https%3A%2F%2Fcdn%2Fx.frag&id=abc123';

function oembedRequest(params: Record<string, string>): Request {
  const url = new URL('https://btc-ifc-viewer-2.vercel.app/api/oembed');
  for (const [k, v] of Object.entries(params)) url.searchParams.set(k, v);
  return new Request(url.toString(), { method: 'GET' });
}

describe('api · oEmbed contract', () => {
  it('returns a valid rich oEmbed 1.0 response with an iframe for an own embed url', async () => {
    const res = handleOEmbed(oembedRequest({ url: EMBED_URL, format: 'json' }), { host: HOST });
    expect(res.status).toBe(200);
    expect(res.headers.get('content-type')).toContain('application/json');
    const body = (await res.json()) as OEmbedResponse;
    expect(body.version).toBe('1.0');
    expect(body.type).toBe('rich');
    expect(body.provider_name).toBe('BTC IFC Viewer');
    expect(body.provider_url).toBe(HOST.origin);
    expect(typeof body.width).toBe('number');
    expect(typeof body.height).toBe('number');
    expect(body.html).toContain('<iframe');
    expect(body.html).toContain('allowfullscreen');
    // The embed URL is reflected into the iframe src (ampersand escaped).
    expect(body.html).toContain('src="https://btc-ifc-viewer-2.vercel.app/embed.html?m=');
  });

  it('honours maxwidth / maxheight within bounds', async () => {
    const res = handleOEmbed(oembedRequest({ url: EMBED_URL, maxwidth: '1024', maxheight: '640' }), { host: HOST });
    const body = (await res.json()) as OEmbedResponse;
    expect(body.width).toBe(1024);
    expect(body.height).toBe(640);
  });

  it('clamps absurd dimensions', async () => {
    const res = handleOEmbed(oembedRequest({ url: EMBED_URL, maxwidth: '999999', maxheight: '1' }), { host: HOST });
    const body = (await res.json()) as OEmbedResponse;
    expect(body.width).toBe(4096); // MAX_DIMENSION
    expect(body.height).toBe(120); // MIN_DIMENSION
  });

  it('400s when url is missing', () => {
    const res = handleOEmbed(oembedRequest({ format: 'json' }), { host: HOST });
    expect(res.status).toBe(400);
  });

  it('501s for a non-json format', () => {
    const res = handleOEmbed(oembedRequest({ url: EMBED_URL, format: 'xml' }), { host: HOST });
    expect(res.status).toBe(501);
  });

  it('405s for a non-GET method', () => {
    const req = new Request('https://x/api/oembed?url=' + encodeURIComponent(EMBED_URL), { method: 'POST' });
    expect(handleOEmbed(req, { host: HOST }).status).toBe(405);
  });

  it('refuses to reflect a foreign / non-embed URL (SSRF/clickjacking hygiene)', () => {
    const foreign = handleOEmbed(oembedRequest({ url: 'https://evil.example.com/embed.html' }), { host: HOST });
    expect(foreign.status).toBe(404);

    const wrongPath = handleOEmbed(oembedRequest({ url: 'https://btc-ifc-viewer-2.vercel.app/index.html' }), { host: HOST });
    expect(wrongPath.status).toBe(404);
  });
});

describe('api · isOwnEmbedUrl', () => {
  it('accepts our embed.html and rejects everything else', () => {
    expect(isOwnEmbedUrl('https://btc-ifc-viewer-2.vercel.app/embed.html?m=x', HOST)).toBe(true);
    expect(isOwnEmbedUrl('https://btc-ifc-viewer-2.vercel.app/embed.html', HOST)).toBe(true);
    expect(isOwnEmbedUrl('https://evil.com/embed.html', HOST)).toBe(false);
    expect(isOwnEmbedUrl('https://btc-ifc-viewer-2.vercel.app/', HOST)).toBe(false);
    expect(isOwnEmbedUrl('javascript:alert(1)', HOST)).toBe(false);
    expect(isOwnEmbedUrl('not a url', HOST)).toBe(false);
  });
});

describe('api · buildIframeHtml escaping', () => {
  it('escapes attribute-breaking characters in the src URL', () => {
    const html = buildIframeHtml('https://x/embed.html?a="><script>alert(1)</script>', 800, 600);
    expect(html).not.toContain('"><script>');
    expect(html).toContain('&quot;');
    expect(html).toContain('&lt;script&gt;');
  });
});
