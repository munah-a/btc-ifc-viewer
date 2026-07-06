import { describe, expect, it } from 'vitest';

import { escapeHtml, filterListMarkup, spatialTreeMarkup } from '../../src/core/markup';

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

describe('spatialTreeMarkup (A1)', () => {
  it('escapes hostile model ids and storey names in text and attributes', () => {
    const markup = spatialTreeMarkup([
      { modelId: HOSTILE_NAME, levels: [{ name: HOSTILE_NAME, count: 3 }, { name: HOSTILE_ATTR, count: 1 }] },
    ]);
    expect(markup).not.toContain('<img');
    expect(markup).toContain('&lt;img');
    // Attribute injection stays inside the quoted attribute.
    expect(markup).not.toContain('" onmouseover="');
    expect(markup).toContain('&quot; onmouseover=&quot;');
  });

  it('renders counts and empty state', () => {
    const markup = spatialTreeMarkup([{ modelId: 'm.ifc', levels: [] }]);
    expect(markup).toContain('m.ifc');
    expect(markup).toContain('No storeys detected');
  });
});

describe('filterListMarkup (A1)', () => {
  it('escapes hostile class/level names in value attributes and labels', () => {
    for (const type of ['class', 'level'] as const) {
      const markup = filterListMarkup([HOSTILE_NAME, HOSTILE_ATTR], type, 'empty');
      expect(markup).not.toContain('<img');
      expect(markup).toContain('&lt;img');
      expect(markup).not.toContain('" onmouseover="');
      expect(markup).toContain(`data-filter-type="${type}"`);
    }
  });

  it('renders the empty message when no names exist', () => {
    expect(filterListMarkup([], 'class', 'No classes detected')).toContain('No classes detected');
  });
});
