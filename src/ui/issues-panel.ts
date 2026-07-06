/**
 * Issues panel (Issues tab) — pure markup builders for the issue list and the
 * active issue's comment thread. DOM-free: the orchestrator writes the returned
 * markup into `#issuesList` / `#issueComments` and handles clicks via one
 * delegated listener (data-issue-id / data-issue-action). Every value is
 * escaped (A1).
 *
 * i18n (C7): these builders are language-agnostic. The orchestrator passes an
 * `IssueListLabels` bundle (already-translated strings + a status/priority
 * localizer + a date formatter) so the module never imports the catalog and
 * stays unit-testable in any language.
 */
import { escapeHtml } from '../core/markup';
import { icon } from './icons';
import type { IssueRecord } from '../core/viewer-types';

const PRIORITY_COLOR: Record<string, string> = {
  Critical: 'var(--error)',
  High: 'var(--warning)',
  Medium: 'var(--primary)',
  Low: 'var(--outline)',
};

export interface IssueListLabels {
  /** `{count} linked · {models} model(s)` — pre-interpolated by the caller. */
  linked(count: number, models: number): string;
  /** "No element link". */
  noLink: string;
  /** English enum → localized display (value stays English in storage). */
  localizePriority(priority: string): string;
  localizeStatus(status: string): string;
  /** "Delete issue" (button title/aria). */
  deleteTitle: string;
}

/**
 * Builds the issue list markup. Returns `null` when there are no issues so the
 * caller can render its own empty state.
 */
export function buildIssueListMarkup(
  issues: IssueRecord[],
  activeIssueId: string | null,
  labels: IssueListLabels,
): string | null {
  if (issues.length === 0) return null;

  return issues
    .map((issue) => {
      const active = issue.id === activeIssueId ? ' is-active' : '';
      const modelCount = issue.elementsByModel ? Object.keys(issue.elementsByModel).length : (issue.localIds.length > 0 ? 1 : 0);
      const linkedText = issue.localIds.length > 0 ? labels.linked(issue.localIds.length, modelCount) : labels.noLink;
      const escapedId = escapeHtml(issue.id);
      const escapedTitle = escapeHtml(issue.title);
      const escapedMeta = escapeHtml(`${labels.localizePriority(issue.priority)} · ${linkedText}`);
      const dot = PRIORITY_COLOR[issue.priority] ?? 'var(--outline)';
      return `
        <div class="issue-row${active}" data-issue-id="${escapedId}">
          <div class="issue-head">
            <span class="issue-prio-dot" style="background:${dot};"></span>
            <span class="issue-title">${escapedTitle}</span>
            <button type="button" class="icon-btn-danger" data-issue-action="delete" title="${escapeHtml(labels.deleteTitle)}" aria-label="${escapeHtml(labels.deleteTitle)}">${icon('delete', 15)}</button>
          </div>
          <div class="issue-meta">${escapedMeta}</div>
          <div class="issue-status-row">
            <span class="status-badge" style="background:var(--surface-container-high);color:var(--on-surface-variant);">${escapeHtml(labels.localizeStatus(issue.status))}</span>
          </div>
        </div>
      `;
    })
    .join('');
}

/**
 * Builds the comment-thread markup for the active issue. Returns `null` when
 * there is no active issue (caller shows the "select an issue" prompt) and the
 * `noComments` string is returned when the issue simply has no comments yet.
 * `formatCommentDate` localizes each comment's timestamp (brand DD.MM.YYYY).
 */
export function buildIssueCommentsMarkup(
  issue: IssueRecord | undefined,
  noComments: string,
  formatCommentDate: (iso: string) => string,
): string | null {
  if (!issue) return null;
  if (issue.comments.length === 0) return `<div class="comment-item">${escapeHtml(noComments)}</div>`;

  return issue.comments
    .map((comment) => `
      <div class="comment-item">
        <div><strong>${escapeHtml(comment.author)}</strong> - ${escapeHtml(formatCommentDate(comment.createdAt))}</div>
        <div>${escapeHtml(comment.text)}</div>
      </div>
    `)
    .join('');
}
