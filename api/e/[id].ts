/**
 * GET/DELETE /api/e/:id — Vercel dynamic route wrapper (W4.3).
 *
 * GET returns metadata only — crucially the DIRECT Blob-CDN `fragUrl`, so the
 * embed fetches model bytes straight from the CDN with ZERO function egress
 * (C2/cost envelope). DELETE removes the blob + metadata when a valid delete
 * token is presented.
 *
 * The `[id]` path param is read from the request URL (Vercel maps the file name
 * to the dynamic segment). We parse it here rather than relying on a
 * framework-specific context object, so the logic stays framework-agnostic and
 * unit-testable.
 */
import { handleEmbedMeta } from '../_lib/hosting.js';
import { error } from '../_lib/http.js';
import { createRealStorage } from '../_lib/storage.js';

export const config = { maxDuration: 10 };

/** Extracts the `:id` segment from `/api/e/<id>`. */
function idFromUrl(request: Request): string {
  const { pathname } = new URL(request.url);
  const parts = pathname.split('/').filter(Boolean); // ['api','e','<id>']
  return decodeURIComponent(parts[parts.length - 1] ?? '');
}

export async function GET(request: Request): Promise<Response> {
  try {
    const storage = await createRealStorage();
    return await handleEmbedMeta(request, idFromUrl(request), storage);
  } catch (err) {
    console.error('[api/e/:id GET] unexpected error', err);
    return error(500, 'internal_error');
  }
}

export async function DELETE(request: Request): Promise<Response> {
  try {
    const storage = await createRealStorage();
    return await handleEmbedMeta(request, idFromUrl(request), storage);
  } catch (err) {
    console.error('[api/e/:id DELETE] unexpected error', err);
    return error(500, 'internal_error');
  }
}
