/**
 * oEmbed provider logic (W4.3) — lets paste-a-link boards (Miro, Notion,
 * Confluence, etc.) auto-expand a BTC embed URL into an inline iframe.
 *
 * Contract: https://oembed.com — we return a `rich` type with an `html` iframe
 * snippet. Pure `(Request) => Response` so it is unit-testable (the oEmbed
 * contract test in tests/unit/api-oembed.spec.ts) with no Vercel packages.
 *
 * We do NOT need storage here: the `url` param already points at a self-contained
 * embed page (which carries the model URL in its own query), so oEmbed just
 * reflects it into an iframe wrapper. No model bytes touch the function (C2).
 */
import { error, json, methodNotAllowed } from './http.js';
import { resolveHostConfig, type HostConfig } from './hosting.js';

const DEFAULT_WIDTH = 800;
const DEFAULT_HEIGHT = 600;
const MAX_DIMENSION = 4096;
const MIN_DIMENSION = 120;

export interface OEmbedResponse {
  version: '1.0';
  type: 'rich';
  provider_name: string;
  provider_url: string;
  title: string;
  width: number;
  height: number;
  html: string;
}

/** Builds the iframe HTML snippet for an embed URL (escaped for attribute context). */
export function buildIframeHtml(embedUrl: string, width: number, height: number): string {
  const safeUrl = escapeAttr(embedUrl);
  return (
    `<iframe src="${safeUrl}" width="${width}" height="${height}" ` +
    `frameborder="0" allow="fullscreen" allowfullscreen ` +
    `style="border:0;max-width:100%;" loading="lazy" ` +
    `title="BTC IFC Viewer embed"></iframe>`
  );
}

/**
 * Handles GET /api/oembed?url=<embed-url>&format=json&maxwidth=&maxheight=.
 * `host` is injectable for tests; otherwise resolved from env/request.
 */
export function handleOEmbed(request: Request, opts: { host?: HostConfig } = {}): Response {
  if (request.method !== 'GET') return methodNotAllowed(['GET']);

  const url = new URL(request.url);
  const target = url.searchParams.get('url');
  const format = url.searchParams.get('format') ?? 'json';

  if (format !== 'json') {
    // We only implement JSON (XML oEmbed is legacy and rarely required).
    return error(501, 'not_implemented', 'Only format=json is supported.');
  }
  if (!target) {
    return error(400, 'missing_url', 'The url parameter is required.');
  }

  const host = opts.host ?? resolveHostConfig(request);
  // Only expand URLs that belong to our own embed surface — never reflect an
  // arbitrary third-party URL into an iframe (SSRF/clickjacking hygiene).
  if (!isOwnEmbedUrl(target, host)) {
    return error(404, 'not_found', 'This URL is not a recognized BTC embed.');
  }

  const width = clampDimension(url.searchParams.get('maxwidth'), DEFAULT_WIDTH);
  const height = clampDimension(url.searchParams.get('maxheight'), DEFAULT_HEIGHT);

  const body: OEmbedResponse = {
    version: '1.0',
    type: 'rich',
    provider_name: 'BTC IFC Viewer',
    provider_url: host.origin,
    title: 'BTC IFC Viewer',
    width,
    height,
    html: buildIframeHtml(target, width, height),
  };
  // oEmbed responses may be cached briefly by consumers.
  return json(body, 200, { 'cache-control': 'public, max-age=3600' });
}

/** True when `candidate` is an embed page on our own origin. */
export function isOwnEmbedUrl(candidate: string, host: HostConfig): boolean {
  let parsed: URL;
  try {
    parsed = new URL(candidate);
  } catch {
    return false;
  }
  if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') return false;
  let originHost: string;
  try {
    originHost = new URL(host.origin).host;
  } catch {
    return false;
  }
  return parsed.host === originHost && parsed.pathname.replace(/\/+$/, '').endsWith('/embed.html');
}

function clampDimension(raw: string | null, fallback: number): number {
  if (!raw) return fallback;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(MAX_DIMENSION, Math.max(MIN_DIMENSION, n));
}

/** Escapes a string for use inside a double-quoted HTML attribute. */
function escapeAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
