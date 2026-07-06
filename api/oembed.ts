/**
 * GET /api/oembed — oEmbed provider route wrapper (W4.3).
 *
 * Pure reflection of an own-origin embed URL into an iframe snippet; no storage,
 * no model bytes (C2). Register this endpoint in embed.html via a
 * <link rel="alternate" type="application/json+oembed"> so consumers discover it.
 */
import { handleOEmbed } from './_lib/oembed';
import { error } from './_lib/http';

export const config = { maxDuration: 10 };

export function GET(request: Request): Response {
  try {
    return handleOEmbed(request);
  } catch (err) {
    console.error('[api/oembed] unexpected error', err);
    return error(500, 'internal_error');
  }
}
