import { describe, expect, it } from 'vitest';

import { escapeHtml, filterChipMarkup } from '../../src/core/markup';

// AUDIT A1: IFC-derived strings (storey/class names, model ids) reach
// innerHTML — a hostile name must never survive as live markup.
const HOSTILE_NAME = '<img src=x onerror=window.__xss=1>';
const HOSTILE_ATTR = '" onmouseover="window.__xss=1';

describe('escapeHtml', () => {
  it('escapes every HTML-significant character', () => {
    expect(escapeHtml('&<>"\'')).toBe('&amp;&lt;&gt;&quot;&#39;');
  });

  it('neutralizes a hostile storey name', () => {
    const escaped = escapeHtml(HOSTILE_NAME);
    expect(escaped).not.toContain('<img');
    expect(escaped).toContain('&lt;img');
  });
});

describe('filterChipMarkup (A1)', () => {
  it('escapes hostile class/level names in value attributes and labels', () => {
    for (const type of ['class', 'level'] as const) {
      const markup = filterChipMarkup([HOSTILE_NAME, HOSTILE_ATTR], type, 'empty');
      expect(markup).not.toContain('<img');
      expect(markup).toContain('&lt;img');
      expect(markup).not.toContain('" onmouseover="');
      expect(markup).toContain(`data-filter-type="${type}"`);
      // Chips are real buttons with aria-pressed (U6).
      expect(markup).toContain('<button');
      expect(markup).toContain('aria-pressed="false"');
    }
  });

  it('renders the empty message when no names exist', () => {
    expect(filterChipMarkup([], 'class', 'No classes detected')).toContain('No classes detected');
  });
});
