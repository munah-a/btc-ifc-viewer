/**
 * URL-state codec (W4.2) — encodes a shareable view into `?m=<model-url>&vp=<hash>`
 * and decodes it back. Powers two things:
 *   1. the chromeless `/embed` entry (loads the model at `m` and applies `vp`), and
 *   2. "Copy link to view" deep links in the full app.
 *
 * `m` is a plain (encoded) URL string pointing at a fetchable `.frag` (or `.ifc`)
 * — typically a Blob-CDN URL from the hosting API (W4.3), but any CORS-enabled
 * URL works. `vp` is a compact, URL-safe base64 of a small JSON payload holding
 * the camera pose, projection, nav mode, section planes, the visual toggles and
 * a hidden-items summary (counts + a capped id list per model, so a deep link
 * stays short even for large hidden sets).
 *
 * DOM-free and engine-free on purpose so it is unit-testable in Node (round-trip
 * tests in tests/unit/url-state.spec.ts). Decoding is defensive: any malformed
 * component yields `null`/omitted fields rather than throwing, because the input
 * is untrusted (a hand-edited or truncated URL).
 */

import type { CameraProjection, NavigationMode, Vector3Record, VisualStyle } from './persistence';

/** The camera pose carried in a viewpoint hash. */
export interface UrlCameraState {
  position: Vector3Record;
  target: Vector3Record;
  projection: CameraProjection;
  mode: NavigationMode;
}

/** One section/clipping plane (world-space normal + a point on the plane). */
export interface UrlClippingPlane {
  normal: Vector3Record;
  origin: Vector3Record;
}

/**
 * The decoded viewpoint state. All fields optional so a partial hash (e.g. just
 * a camera, no sections) round-trips cleanly and callers apply what is present.
 */
export interface UrlViewpointState {
  camera?: UrlCameraState;
  clippingPlanes?: UrlClippingPlane[];
  visualStyle?: VisualStyle;
  xray?: boolean;
  edges?: boolean;
  /**
   * Hidden-items summary per model id. `count` is the true number hidden;
   * `ids` is a capped sample (see HIDDEN_IDS_CAP) so the URL stays short — the
   * embed hides the listed ids and can show "+N more hidden" from `count`.
   */
  hidden?: Record<string, { count: number; ids: number[] }>;
}

/** The fully-decoded URL state: a model URL and/or a viewpoint. */
export interface UrlState {
  modelUrl?: string;
  viewpoint?: UrlViewpointState;
}

/** Query-param names (kept short for compact links). */
export const URL_PARAM_MODEL = 'm';
export const URL_PARAM_VIEWPOINT = 'vp';

/** Cap on hidden-id samples per model in a shared link (keeps URLs bounded). */
export const HIDDEN_IDS_CAP = 200;

/** Coordinates are rounded to this many decimals before encoding (compactness). */
const COORD_PRECISION = 4;

const VISUAL_STYLES: readonly VisualStyle[] = [
  'basic',
  'pen',
  'color-pen',
  'color-shadows',
  'color-pen-shadows',
];
const PROJECTIONS: readonly CameraProjection[] = ['Perspective', 'Orthographic'];
const NAV_MODES: readonly NavigationMode[] = ['Orbit', 'Plan', 'FirstPerson'];

// ─────────────────────────────────────────────────────────────────────────────
// base64url (URL-safe, no padding). Uses only Web-standard btoa/atob +
// TextEncoder/TextDecoder, which are globally available in browsers AND in the
// Node runtime the API/tests use — so this module has no Node `Buffer` dep and
// type-checks under the browser (src) tsconfig. The TextEncoder step makes the
// round-trip safe for non-ASCII characters (btoa alone is Latin-1 only).
// ─────────────────────────────────────────────────────────────────────────────

function utf8ToBase64Url(input: string): string {
  const bytes = new TextEncoder().encode(input);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  const base64 = btoa(binary);
  return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function base64UrlToUtf8(input: string): string | null {
  try {
    const normalized = input.replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
    const binary = atob(padded);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Compact wire format. Short keys keep the encoded hash small; the shape is an
// internal detail (only this module reads/writes it), so terseness is fine.
//   c: camera { p:[x,y,z], t:[x,y,z], o: projection index, m: nav-mode index }
//   s: sections [ { n:[x,y,z], o:[x,y,z] }, … ]
//   v: visual-style index
//   x: xray (0|1)   e: edges (0|1)
//   h: hidden { modelId: [count, id, id, …(capped)] }
// ─────────────────────────────────────────────────────────────────────────────

interface WireCamera {
  p: number[];
  t: number[];
  o: number;
  m: number;
}
interface WirePlane {
  n: number[];
  o: number[];
}
interface WireViewpoint {
  c?: WireCamera;
  s?: WirePlane[];
  v?: number;
  x?: 0 | 1;
  e?: 0 | 1;
  h?: Record<string, number[]>;
}

const round = (n: number): number => {
  const factor = 10 ** COORD_PRECISION;
  return Math.round(n * factor) / factor;
};

const vecToArr = (v: Vector3Record): number[] => [round(v.x), round(v.y), round(v.z)];

const isFiniteNum = (n: unknown): n is number => typeof n === 'number' && Number.isFinite(n);

const arrToVec = (a: unknown): Vector3Record | null => {
  if (!Array.isArray(a) || a.length !== 3) return null;
  const [x, y, z] = a as unknown[];
  if (!isFiniteNum(x) || !isFiniteNum(y) || !isFiniteNum(z)) return null;
  return { x, y, z };
};

/**
 * Encodes a viewpoint state into a compact URL-safe hash. Returns '' for an
 * empty viewpoint (caller can then omit the `vp` param entirely).
 */
export function encodeViewpoint(state: UrlViewpointState): string {
  const wire: WireViewpoint = {};

  if (state.camera) {
    wire.c = {
      p: vecToArr(state.camera.position),
      t: vecToArr(state.camera.target),
      o: Math.max(0, PROJECTIONS.indexOf(state.camera.projection)),
      m: Math.max(0, NAV_MODES.indexOf(state.camera.mode)),
    };
  }
  if (state.clippingPlanes && state.clippingPlanes.length > 0) {
    wire.s = state.clippingPlanes.map((plane) => ({
      n: vecToArr(plane.normal),
      o: vecToArr(plane.origin),
    }));
  }
  if (state.visualStyle) {
    const idx = VISUAL_STYLES.indexOf(state.visualStyle);
    if (idx >= 0) wire.v = idx;
  }
  if (state.xray) wire.x = 1;
  if (state.edges) wire.e = 1;
  if (state.hidden) {
    const h: Record<string, number[]> = {};
    for (const [modelId, entry] of Object.entries(state.hidden)) {
      if (!entry || !isFiniteNum(entry.count) || entry.count <= 0) continue;
      const ids = Array.isArray(entry.ids)
        ? entry.ids.filter(isFiniteNum).slice(0, HIDDEN_IDS_CAP)
        : [];
      // Leading element is the true count; the rest is the capped id sample.
      h[modelId] = [entry.count, ...ids];
    }
    if (Object.keys(h).length > 0) wire.h = h;
  }

  if (Object.keys(wire).length === 0) return '';
  return utf8ToBase64Url(JSON.stringify(wire));
}

const isRecord = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && !Array.isArray(v);

/**
 * Decodes a viewpoint hash. Returns null when the hash is empty/undecodable;
 * otherwise returns whatever valid fields were present (malformed fields are
 * dropped, never thrown — the input is untrusted).
 */
export function decodeViewpoint(hash: string | null | undefined): UrlViewpointState | null {
  if (!hash) return null;
  const json = base64UrlToUtf8(hash);
  if (json === null) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    return null;
  }
  if (!isRecord(parsed)) return null;

  const state: UrlViewpointState = {};

  const c = parsed.c;
  if (isRecord(c)) {
    const position = arrToVec(c.p);
    const target = arrToVec(c.t);
    if (position && target) {
      const projIdx = isFiniteNum(c.o) ? c.o : 0;
      const modeIdx = isFiniteNum(c.m) ? c.m : 0;
      state.camera = {
        position,
        target,
        projection: PROJECTIONS[projIdx] ?? 'Perspective',
        mode: NAV_MODES[modeIdx] ?? 'Orbit',
      };
    }
  }

  if (Array.isArray(parsed.s)) {
    const planes: UrlClippingPlane[] = [];
    for (const raw of parsed.s) {
      if (!isRecord(raw)) continue;
      const normal = arrToVec(raw.n);
      const origin = arrToVec(raw.o);
      if (normal && origin) planes.push({ normal, origin });
    }
    if (planes.length > 0) state.clippingPlanes = planes;
  }

  if (isFiniteNum(parsed.v) && VISUAL_STYLES[parsed.v]) {
    state.visualStyle = VISUAL_STYLES[parsed.v];
  }
  if (parsed.x === 1) state.xray = true;
  if (parsed.e === 1) state.edges = true;

  if (isRecord(parsed.h)) {
    const hidden: Record<string, { count: number; ids: number[] }> = {};
    for (const [modelId, value] of Object.entries(parsed.h)) {
      if (!Array.isArray(value) || value.length === 0) continue;
      const [count, ...ids] = value as unknown[];
      if (!isFiniteNum(count) || count <= 0) continue;
      hidden[modelId] = {
        count,
        ids: ids.filter(isFiniteNum).slice(0, HIDDEN_IDS_CAP),
      };
    }
    if (Object.keys(hidden).length > 0) state.hidden = hidden;
  }

  return Object.keys(state).length > 0 ? state : null;
}

/**
 * Reads `?m=&vp=` from a search string (e.g. `location.search`) into a UrlState.
 * A missing/blank `m` yields no `modelUrl`; a bad `vp` yields no `viewpoint`.
 */
export function decodeUrlState(search: string): UrlState {
  const params = new URLSearchParams(search.startsWith('?') ? search.slice(1) : search);
  const state: UrlState = {};

  const model = params.get(URL_PARAM_MODEL);
  if (model && model.trim()) state.modelUrl = model.trim();

  const viewpoint = decodeViewpoint(params.get(URL_PARAM_VIEWPOINT));
  if (viewpoint) state.viewpoint = viewpoint;

  return state;
}

/**
 * Builds a `?m=&vp=` query string (leading '?', or '' when nothing to encode)
 * from a model URL and an optional viewpoint. The model URL is percent-encoded
 * so a Blob-CDN URL with its own query string survives intact.
 */
export function encodeUrlState(state: UrlState): string {
  const params = new URLSearchParams();
  if (state.modelUrl && state.modelUrl.trim()) {
    params.set(URL_PARAM_MODEL, state.modelUrl.trim());
  }
  if (state.viewpoint) {
    const hash = encodeViewpoint(state.viewpoint);
    if (hash) params.set(URL_PARAM_VIEWPOINT, hash);
  }
  const query = params.toString();
  return query ? `?${query}` : '';
}

/**
 * Builds an absolute deep link to the full app or the embed. `baseUrl` is the
 * page origin+path (e.g. `https://btc-ifc-viewer-2.vercel.app/` or
 * `.../embed.html`). Used by "Copy link to view" (W4.2) and the share dialog.
 */
export function buildShareUrl(baseUrl: string, state: UrlState): string {
  const query = encodeUrlState(state);
  // Strip any existing query/hash on the base so we own the search string.
  const base = baseUrl.split('#')[0].split('?')[0];
  return `${base}${query}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// S8: model-URL allowlist. A BTC-branded embed that renders ANY attacker-supplied
// `?m=` URL is a phishing / content-injection vector. `isAllowedModelUrl` gates
// which hosts the embed will fetch a model from: the page's own origin, the
// configured hosting/embed origins, and the Vercel Blob CDN. Pure + unit-tested.
// ─────────────────────────────────────────────────────────────────────────────

/** The Vercel Blob CDN host suffix (all hosted `.frag` URLs live under this). */
export const VERCEL_BLOB_HOST_SUFFIX = '.public.blob.vercel-storage.com';

export interface ModelUrlPolicy {
  /** The embed page's own origin (same-origin models always allowed). */
  selfOrigin?: string;
  /** Extra allowed origins (e.g. BTC_EMBED_ORIGIN / a custom Blob domain). */
  allowedOrigins?: string[];
  /** Allow any `*.public.blob.vercel-storage.com` host (default true). */
  allowVercelBlob?: boolean;
}

/**
 * Returns true iff `modelUrl` is a fetchable model source we trust to render in
 * the embed. Only http(s) URLs on an allowed host pass; everything else
 * (javascript:, data:, foreign hosts) is rejected.
 */
export function isAllowedModelUrl(modelUrl: string, policy: ModelUrlPolicy = {}): boolean {
  let parsed: URL;
  try {
    parsed = new URL(modelUrl);
  } catch {
    return false;
  }
  if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') return false;

  const host = parsed.host.toLowerCase();
  const allowVercelBlob = policy.allowVercelBlob ?? true;
  if (allowVercelBlob && host.endsWith(VERCEL_BLOB_HOST_SUFFIX)) return true;

  const allowedHosts = new Set<string>();
  const addHost = (origin: string | undefined): void => {
    if (!origin) return;
    try {
      allowedHosts.add(new URL(origin).host.toLowerCase());
    } catch {
      // ignore malformed configured origin
    }
  };
  addHost(policy.selfOrigin);
  for (const origin of policy.allowedOrigins ?? []) addHost(origin);

  return allowedHosts.has(host);
}
