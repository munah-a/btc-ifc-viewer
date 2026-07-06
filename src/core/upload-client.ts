/**
 * Client for the hosting API (W4.4). Thin fetch wrappers the share dialog uses
 * to publish the browser-converted `.frag` and to delete an upload with its
 * token. Kept DOM-free so it is unit-testable with a mocked `fetch`.
 *
 * C2: the CLIENT converts IFC→fragments and uploads the `.frag` bytes here; the
 * server only stores/serves them. This module sends opaque bytes.
 */

/** The successful publish response from POST /api/uploads. */
export interface PublishResult {
  id: string;
  embedUrl: string;
  viewerUrl: string;
  fragUrl: string;
  deleteToken: string;
  expiresAt: string;
}

export interface UploadClientOptions {
  /** Base path for the API (default '/api'). */
  apiBase?: string;
  /** Injected fetch (tests). Defaults to global fetch. */
  fetchImpl?: typeof fetch;
}

export class UploadClient {
  private readonly apiBase: string;
  private readonly fetchImpl: typeof fetch;

  constructor(options: UploadClientOptions = {}) {
    this.apiBase = options.apiBase ?? '/api';
    this.fetchImpl = options.fetchImpl ?? globalThis.fetch.bind(globalThis);
  }

  /**
   * Uploads `.frag` bytes and returns the share result. Throws an Error with the
   * server's message on non-2xx (size cap, quota, rate limit).
   */
  async publish(fragBytes: Uint8Array, fileName: string): Promise<PublishResult> {
    const response = await this.fetchImpl(`${this.apiBase}/uploads`, {
      method: 'POST',
      headers: {
        'content-type': 'application/octet-stream',
        'x-file-name': fileName,
      },
      body: fragBytes as BodyInit,
    });
    if (!response.ok) {
      throw new Error(await extractError(response));
    }
    return (await response.json()) as PublishResult;
  }

  /** Deletes an upload with its delete token. Throws on failure. */
  async remove(id: string, deleteToken: string): Promise<void> {
    const response = await this.fetchImpl(`${this.apiBase}/e/${encodeURIComponent(id)}`, {
      method: 'DELETE',
      headers: { 'x-delete-token': deleteToken },
    });
    if (!response.ok) {
      throw new Error(await extractError(response));
    }
  }
}

/** Pulls a human message out of a JSON `{error,message}` body, else the status. */
async function extractError(response: Response): Promise<string> {
  try {
    const body = (await response.json()) as { message?: string; error?: string };
    return body.message ?? body.error ?? `HTTP ${response.status}`;
  } catch {
    return `HTTP ${response.status}`;
  }
}
