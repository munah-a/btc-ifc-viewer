/**
 * GET /api/cron-cleanup — daily TTL sweep (W4.3, C3).
 *
 * Invoked by the Vercel Cron entry in vercel.json. Deletes expired uploads
 * (blob + metadata). Guarded by CRON_SECRET (Vercel sends it as a Bearer token);
 * see handleCronCleanup. The metadata store (Redis) also auto-expires records,
 * so this primarily reaps blob objects (the blob store has no TTL of its own).
 */
import { handleCronCleanup } from './_lib/hosting.js';
import { error } from './_lib/http.js';
import { createRealStorage } from './_lib/storage.js';

export const config = { maxDuration: 60 };

export async function GET(request: Request): Promise<Response> {
  try {
    const storage = await createRealStorage();
    return await handleCronCleanup(request, storage);
  } catch (err) {
    console.error('[api/cron-cleanup] unexpected error', err);
    return error(500, 'internal_error');
  }
}
