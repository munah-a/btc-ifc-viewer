/**
 * POST /api/uploads — Vercel route wrapper (W4.3).
 *
 * Thin: instantiate the real (Blob + Redis) storage adapter and delegate to the
 * unit-tested `handleUpload` logic. The browser sends the already-converted
 * `.frag` bytes (application/octet-stream) with `x-file-name`; the server
 * validates size/rate/quota and stores bytes only (C2 — no model processing).
 *
 * maxDuration is bounded low: this is a bytes-in, metadata-out endpoint, not a
 * compute endpoint.
 */
import { handleUpload } from './_lib/hosting';
import { error } from './_lib/http';
import { createRealStorage } from './_lib/storage';

export const config = { maxDuration: 30 };

export async function POST(request: Request): Promise<Response> {
  try {
    const storage = await createRealStorage();
    return await handleUpload(request, storage);
  } catch (err) {
    console.error('[api/uploads] unexpected error', err);
    return error(500, 'internal_error', 'Upload failed. Please try again.');
  }
}
