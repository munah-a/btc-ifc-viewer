/**
 * HTML escaping + pure markup builders for the render paths that interpolate
 * IFC-derived strings (AUDIT A1: the class/level filters and tree renderers are
 * the innerHTML sinks that receive IFC-derived class/level names).
 *
 * DOM-free so the escaping contract is unit-testable (tests/unit/markup.spec.ts
 * feeds a hostile storey name through every builder).
 */

export const escapeHtml = (value: string): string => value
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

/**
 * Filter chips (W3 design: pill buttons, not checkboxes). `data-filter-value`
 * carries the (escaped) name; `aria-pressed` reflects state. U6: real
 * `<button>` rows, keyboard-reachable. Every interpolation is escaped (A1).
 */
export const filterChipMarkup = (
  names: string[],
  filterType: 'class' | 'level',
  emptyMessage: string,
): string => {
  if (names.length === 0) return `<div class="list-empty">${escapeHtml(emptyMessage)}</div>`;
  return names
    .map((name) => {
      const escaped = escapeHtml(name);
      return `<button type="button" class="filter-chip" data-filter-type="${filterType}" data-filter-value="${escaped}" aria-pressed="false">${escaped}</button>`;
    })
    .join('');
};
