/**
 * HTML escaping + pure markup builders for the render paths that interpolate
 * IFC-derived strings (AUDIT A1: renderSpatialTree / renderClassFilters /
 * renderLevelFilters were the only unescaped innerHTML sinks).
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

export interface SpatialTreeLevelEntry {
  name: string;
  count: number;
}

export interface SpatialTreeModelEntry {
  modelId: string;
  levels: SpatialTreeLevelEntry[];
}

/** Rows for the Explorer spatial tree. Every interpolation is escaped. */
export const spatialTreeMarkup = (models: SpatialTreeModelEntry[]): string => {
  const rows: string[] = [];
  for (const model of models) {
    const escapedModelId = escapeHtml(model.modelId);
    rows.push(`<div class="tree-item"><strong>${escapedModelId}</strong></div>`);
    for (const level of model.levels) {
      const escapedLevel = escapeHtml(level.name);
      rows.push(
        `<div class="tree-item" data-model-id="${escapedModelId}" data-level="${escapedLevel}">Level: ${escapedLevel} (${level.count})</div>`,
      );
    }
    if (model.levels.length === 0) rows.push('<div class="tree-item">No storeys detected</div>');
  }
  return rows.join('');
};

/** Checkbox list for the class/level filters. Every interpolation is escaped. */
export const filterListMarkup = (
  names: string[],
  filterType: 'class' | 'level',
  emptyMessage: string,
): string => {
  if (names.length === 0) return `<div class="filter-item">${escapeHtml(emptyMessage)}</div>`;
  return names
    .map((name) => {
      const escaped = escapeHtml(name);
      return `
        <label class="filter-item">
          <input type="checkbox" data-filter-type="${filterType}" value="${escaped}" />
          <span>${escaped}</span>
        </label>
      `;
    })
    .join('');
};
