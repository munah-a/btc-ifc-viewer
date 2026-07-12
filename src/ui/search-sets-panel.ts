/**
 * Search sets (Explorer tab) — pure markup builder for the saved-search list.
 *
 * A search set is a Navisworks-style saved search: the matched elements are
 * captured as a named set with a color override and a visibility eye, so field
 * users can paint and hide/unhide whole systems in one click. DOM-free: the
 * orchestrator writes the markup into `#searchSetList` and handles clicks via
 * one delegated listener (data-set-id / data-set-action). Every value is
 * escaped (A1).
 *
 * i18n (C7): language-agnostic — the orchestrator passes an already-translated
 * labels bundle so this module never imports the catalog.
 */
import { escapeHtml } from '../core/markup';
import { icon } from './icons';

export interface SearchSetView {
  id: string;
  name: string;
  /** Override color (hex string, e.g. '#e4572e'). */
  color: string;
  /** Whether the color override is currently painted on the elements. */
  colorActive: boolean;
  /** Whether the set's elements are currently shown. */
  visible: boolean;
  /** Total captured element count across models. */
  count: number;
}

export interface SearchSetLabels {
  /** `{count} elements` — pre-interpolated by the caller. */
  countText(count: number): string;
  /** Toggle-override button title (state-aware). */
  colorTitle(active: boolean): string;
  /** Eye button title (state-aware). */
  visibilityTitle(visible: boolean): string;
  /** Row button title ("Select set elements"). */
  selectTitle: string;
  /** Delete button title. */
  deleteTitle: string;
}

/**
 * Builds the search-set list markup. Returns `null` when there are no sets so
 * the caller can hide the panel group entirely.
 */
export function buildSearchSetListMarkup(sets: SearchSetView[], labels: SearchSetLabels): string | null {
  if (sets.length === 0) return null;

  return sets
    .map((set) => {
      const escapedId = escapeHtml(set.id);
      const escapedName = escapeHtml(set.name);
      const escapedColor = escapeHtml(set.color);
      const swatchState = set.colorActive ? ' is-active' : '';
      const rowState = set.visible ? '' : ' is-hidden';
      return `
        <div class="set-row${rowState}" data-set-id="${escapedId}">
          <button type="button" class="set-swatch${swatchState}" data-set-action="color"
            style="--set-color:${escapedColor};" title="${escapeHtml(labels.colorTitle(set.colorActive))}"
            aria-label="${escapeHtml(labels.colorTitle(set.colorActive))}" aria-pressed="${set.colorActive}"></button>
          <button type="button" class="set-name" data-set-action="select" title="${escapeHtml(labels.selectTitle)}">
            <span class="set-title">${escapedName}</span>
            <span class="set-count">${escapeHtml(labels.countText(set.count))}</span>
          </button>
          <button type="button" class="set-eye" data-set-action="visibility"
            title="${escapeHtml(labels.visibilityTitle(set.visible))}"
            aria-label="${escapeHtml(labels.visibilityTitle(set.visible))}" aria-pressed="${set.visible}">
            ${icon(set.visible ? 'visibility' : 'visibility_off', 16)}
          </button>
          <button type="button" class="set-delete" data-set-action="delete"
            title="${escapeHtml(labels.deleteTitle)}" aria-label="${escapeHtml(labels.deleteTitle)}">
            ${icon('close', 16)}
          </button>
        </div>`;
    })
    .join('');
}

/**
 * The auto-assigned override palette — distinct, readable on both themes.
 * `searchSetColor(i)` cycles it so consecutive sets get different colors.
 */
export const SEARCH_SET_PALETTE = [
  '#e4572e',
  '#17bebb',
  '#ffc914',
  '#76b041',
  '#9b5de5',
  '#00b4d8',
  '#f15bb5',
  '#8ac926',
] as const;

export function searchSetColor(index: number): string {
  return SEARCH_SET_PALETTE[((index % SEARCH_SET_PALETTE.length) + SEARCH_SET_PALETTE.length) % SEARCH_SET_PALETTE.length];
}
