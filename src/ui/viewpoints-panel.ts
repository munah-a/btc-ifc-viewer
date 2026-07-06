/**
 * Viewpoints panel (Viewpoints tab) — pure markup builder for the saved
 * viewpoint list (thumbnail + name + timestamp + apply/delete row actions).
 * DOM-free: the orchestrator writes the returned markup into `#viewpointList`
 * and handles clicks via one delegated listener (data-viewpoint-id /
 * data-viewpoint-action). Every value is escaped (A1); the snapshot data URI is
 * escaped for the src attribute.
 */
import { escapeHtml } from '../core/markup';
import type { SavedViewpoint } from '../core/persistence';
import { icon } from './icons';

export interface ViewpointListLabels {
  /** "Apply" (row action button). */
  apply: string;
  /** "Delete viewpoint" (button title/aria). */
  deleteTitle: string;
  /** Localized created-at timestamp (brand DD.MM.YYYY, HH:MM). */
  formatDate(iso: string): string;
}

/**
 * Builds the viewpoint list markup. Returns `null` when there are no
 * viewpoints so the caller can render its own empty state. i18n (C7): the
 * caller passes already-translated labels + a date formatter so this stays
 * DOM-free and language-agnostic.
 */
export function buildViewpointListMarkup(
  viewpoints: SavedViewpoint[],
  selectedViewpointId: string | null,
  labels: ViewpointListLabels,
): string | null {
  if (viewpoints.length === 0) return null;

  return viewpoints
    .map((entry) => {
      const active = entry.id === selectedViewpointId ? ' is-active' : '';
      const escapedId = escapeHtml(entry.id);
      const escapedName = escapeHtml(entry.name);
      const thumbnail = entry.snapshot
        ? `<img class="vp-thumb" src="${escapeHtml(entry.snapshot)}" alt="" loading="lazy" />`
        : `<span class="vp-icon">${icon('photo_camera')}</span>`;
      return `
        <div class="vp-row${active}" data-viewpoint-id="${escapedId}">
          <div class="vp-row-main">
            ${thumbnail}
            <div class="vp-text">
              <div class="vp-name">${escapedName}</div>
              <div class="vp-meta">${escapeHtml(labels.formatDate(entry.createdAt))}</div>
            </div>
            <button type="button" class="row-btn" data-viewpoint-action="apply">${escapeHtml(labels.apply)}</button>
            <button type="button" class="icon-btn-danger" data-viewpoint-action="delete" title="${escapeHtml(labels.deleteTitle)}" aria-label="${escapeHtml(labels.deleteTitle)}">${icon('delete', 16)}</button>
          </div>
        </div>
      `;
    })
    .join('');
}
