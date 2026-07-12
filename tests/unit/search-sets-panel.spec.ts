import { describe, expect, it } from 'vitest';
import {
  buildSearchSetListMarkup,
  SEARCH_SET_PALETTE,
  searchSetColor,
  type SearchSetLabels,
  type SearchSetView,
} from '../../src/ui/search-sets-panel';

const labels: SearchSetLabels = {
  countText: (count) => `${count} elements`,
  colorTitle: (active) => (active ? 'Remove color override' : 'Apply color override'),
  visibilityTitle: (visible) => (visible ? 'Hide set elements' : 'Show set elements'),
  selectTitle: 'Select set elements',
  deleteTitle: 'Delete set',
};

const set = (overrides: Partial<SearchSetView> = {}): SearchSetView => ({
  id: 'set-1',
  name: 'walls',
  color: '#e4572e',
  colorActive: true,
  visible: true,
  count: 42,
  ...overrides,
});

describe('buildSearchSetListMarkup', () => {
  it('returns null for an empty list (caller hides the group)', () => {
    expect(buildSearchSetListMarkup([], labels)).toBeNull();
  });

  it('renders a row with swatch, name, count, eye and delete actions', () => {
    const markup = buildSearchSetListMarkup([set()], labels)!;
    expect(markup).toContain('data-set-id="set-1"');
    expect(markup).toContain('data-set-action="color"');
    expect(markup).toContain('data-set-action="select"');
    expect(markup).toContain('data-set-action="visibility"');
    expect(markup).toContain('data-set-action="delete"');
    expect(markup).toContain('--set-color:#e4572e');
    expect(markup).toContain('42 elements');
    expect(markup).toContain('walls');
  });

  it('reflects override + visibility state in classes and aria', () => {
    const active = buildSearchSetListMarkup([set()], labels)!;
    expect(active).toContain('set-swatch is-active');
    expect(active).not.toContain('set-row is-hidden');
    expect(active).toContain('aria-pressed="true"');

    const hidden = buildSearchSetListMarkup([set({ colorActive: false, visible: false })], labels)!;
    expect(hidden).not.toContain('set-swatch is-active');
    expect(hidden).toContain('set-row is-hidden');
    expect(hidden).toContain('Show set elements');
  });

  it('escapes hostile names (A1)', () => {
    const markup = buildSearchSetListMarkup([set({ name: '<img src=x onerror=alert(1)>' })], labels)!;
    expect(markup).not.toContain('<img src=x');
    expect(markup).toContain('&lt;img');
  });
});

describe('searchSetColor', () => {
  it('cycles the palette and never goes out of bounds', () => {
    expect(searchSetColor(0)).toBe(SEARCH_SET_PALETTE[0]);
    expect(searchSetColor(SEARCH_SET_PALETTE.length)).toBe(SEARCH_SET_PALETTE[0]);
    expect(searchSetColor(SEARCH_SET_PALETTE.length + 3)).toBe(SEARCH_SET_PALETTE[3]);
    expect(searchSetColor(-1)).toBe(SEARCH_SET_PALETTE[SEARCH_SET_PALETTE.length - 1]);
  });

  it('assigns distinct colors to consecutive sets', () => {
    const first = new Set([0, 1, 2, 3, 4, 5, 6, 7].map(searchSetColor));
    expect(first.size).toBe(SEARCH_SET_PALETTE.length);
  });
});
