/**
 * Properties panel (Properties tab) — pure markup builder for the property
 * accordion. The section data is produced by core/property-engine
 * (buildPropertySections); this only renders it. DOM-free: the orchestrator
 * writes the returned markup into `#propSections` and applies the text filter
 * (applyPropertiesFilter) separately. Every value is escaped (A1).
 */
import { escapeHtml } from '../core/markup';
import type { PropertySectionData } from '../core/property-engine';
import { treeIco } from './model-browser';

export function buildPropertySectionsMarkup(sections: PropertySectionData[]): string {
  return sections.map((section) => `
    <details class="prop-section" data-prop-section data-search="${escapeHtml(section.title.toLowerCase())}" ${section.defaultOpen ? 'open' : ''}>
      <summary class="prop-section-summary">
        ${treeIco('chevron_right', 'browser-twist')}
        <span class="prop-section-name">${escapeHtml(section.title)}</span>
        <span class="prop-section-count" data-prop-count>${section.rows.length}</span>
      </summary>
      <div class="prop-section-body">
          ${section.rows.map((row) => `
            <div class="prop-row" data-prop-row data-search="${escapeHtml(row.searchText)}">
              <span class="prop-key">${escapeHtml(row.key)}</span>
              <span class="prop-val">${escapeHtml(row.value)}</span>
            </div>
          `).join('')}
      </div>
    </details>
  `).join('');
}
