import { describe, expect, it } from 'vitest';

// The plugin is a plain build-time .mjs (no types, C1 dep-free). Import it via a
// typed alias so the strict no-unsafe-* lint rules don't trip on `any`.
import * as pwaPlugin from '../../scripts/pwa-plugin.mjs';

type HtmlEntriesFromInput = (
  input: string | string[] | Record<string, string> | undefined,
) => string[];
const htmlEntriesFromInput = (pwaPlugin as unknown as { htmlEntriesFromInput: HtmlEntriesFromInput })
  .htmlEntriesFromInput;

/**
 * P1 (W5-fixups): the precache manifest must enumerate EVERY configured HTML
 * entry (index.html AND embed.html), not hard-code index.html — otherwise an
 * offline navigation to /embed serves the full-app shell. htmlEntriesFromInput
 * derives the entry basenames from the resolved rollupOptions.input, which is
 * how the plugin knows about embed.html at generateBundle time (Vite emits the
 * HTML assets later, so they aren't in the bundle yet).
 */
describe('pwaPlugin htmlEntriesFromInput (P1 — W5-fixups)', () => {
  it('extracts both entries from the MPA object input (basenames)', () => {
    const entries = htmlEntriesFromInput({
      main: '/abs/src/index.html',
      embed: '/abs/src/embed.html',
    });
    expect(entries.sort()).toEqual(['embed.html', 'index.html']);
  });

  it('handles an array input', () => {
    expect(htmlEntriesFromInput(['/x/index.html', '/x/embed.html']).sort()).toEqual([
      'embed.html',
      'index.html',
    ]);
  });

  it('handles a single string input', () => {
    expect(htmlEntriesFromInput('/x/index.html')).toEqual(['index.html']);
  });

  it('ignores non-HTML inputs', () => {
    const entries = htmlEntriesFromInput({ a: '/x/index.html', b: '/x/main.ts' });
    expect(entries).toEqual(['index.html']);
  });

  it('falls back to index.html when no HTML entry is configured', () => {
    expect(htmlEntriesFromInput(undefined)).toEqual(['index.html']);
    expect(htmlEntriesFromInput({})).toEqual(['index.html']);
    expect(htmlEntriesFromInput(['/x/main.ts'])).toEqual(['index.html']);
  });

  it('dedupes repeated entries', () => {
    expect(htmlEntriesFromInput(['/a/index.html', '/b/index.html'])).toEqual(['index.html']);
  });
});
