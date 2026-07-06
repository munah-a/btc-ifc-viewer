import { describe, expect, it, vi } from 'vitest';

import { UploadClient } from '../../src/core/upload-client';
import { buildIframeSnippet } from '../../src/ui/share-dialog';

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

/** A fetch mock that resolves to `response` (no async keyword — avoids require-await). */
function mockFetch(response: Response) {
  return vi.fn(() => Promise.resolve(response));
}

const FRAG = new Uint8Array([1, 2, 3]);

const PUBLISH_OK = {
  id: 'abc',
  embedUrl: 'https://host/embed.html?m=x',
  viewerUrl: 'https://host/?m=x',
  fragUrl: 'https://cdn/x.frag',
  deleteToken: 'tok',
  expiresAt: '2026-07-13T00:00:00.000Z',
};

describe('UploadClient · publish', () => {
  it('POSTs octet-stream with the file name header and returns the result', async () => {
    const fetchMock = mockFetch(jsonResponse(PUBLISH_OK, 201));
    const client = new UploadClient({ fetchImpl: fetchMock });
    const result = await client.publish(FRAG, 'tower.frag');

    expect(result.id).toBe('abc');
    expect(result.deleteToken).toBe('tok');
    const mock = vi.mocked(fetchMock);
    expect(mock).toHaveBeenCalledTimes(1);
    const [url, init] = mock.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('/api/uploads');
    expect(init.method).toBe('POST');
    expect((init.headers as Record<string, string>)['content-type']).toBe('application/octet-stream');
    expect((init.headers as Record<string, string>)['x-file-name']).toBe('tower.frag');
  });

  it('throws the server message on a non-2xx response', async () => {
    const client = new UploadClient({ fetchImpl: mockFetch(jsonResponse({ error: 'payload_too_large', message: 'Too big' }, 413)) });
    await expect(client.publish(FRAG, 'x.frag')).rejects.toThrow('Too big');
  });

  it('falls back to the status when the error body is not JSON', async () => {
    const client = new UploadClient({ fetchImpl: mockFetch(new Response('nope', { status: 500 })) });
    await expect(client.publish(FRAG, 'x.frag')).rejects.toThrow('HTTP 500');
  });
});

describe('UploadClient · remove', () => {
  it('DELETEs with the delete-token header', async () => {
    const fetchMock = mockFetch(jsonResponse({ id: 'abc', deleted: true }));
    const client = new UploadClient({ fetchImpl: fetchMock });
    await client.remove('abc', 'tok');
    const [url, init] = vi.mocked(fetchMock).mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('/api/e/abc');
    expect(init.method).toBe('DELETE');
    expect((init.headers as Record<string, string>)['x-delete-token']).toBe('tok');
  });

  it('throws the server message on a failed delete', async () => {
    const client = new UploadClient({ fetchImpl: mockFetch(jsonResponse({ error: 'forbidden', message: 'Invalid delete token.' }, 403)) });
    await expect(client.remove('abc', 'wrong')).rejects.toThrow('Invalid delete token.');
  });

  it('uses a custom apiBase when provided', async () => {
    const fetchMock = mockFetch(jsonResponse({ deleted: true }));
    const client = new UploadClient({ apiBase: '/custom', fetchImpl: fetchMock });
    await client.remove('id1', 'tok');
    expect((vi.mocked(fetchMock).mock.calls[0] as unknown as [string])[0]).toBe('/custom/e/id1');
  });
});

describe('share-dialog · buildIframeSnippet', () => {
  it('builds an iframe with the embed URL and escapes quotes', () => {
    const html = buildIframeSnippet('https://host/embed.html?m=a"b');
    expect(html).toContain('<iframe');
    expect(html).toContain('allowfullscreen');
    expect(html).toContain('src="https://host/embed.html?m=a&quot;b"');
  });
});
