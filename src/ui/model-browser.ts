/**
 * Model Browser panel (Explorer tab) — pure markup builders.
 *
 * These build the nested `<details>` tree HTML for the loaded models from the
 * in-memory `ModelIndex`/`FederatedModelRecord` data. They are DOM-free string
 * builders: the orchestrator (viewer.ts) writes the returned markup into
 * `#modelBrowserTree` via `renderPreservingDetails` (A11 open-state preserve)
 * and handles clicks via one delegated listener keyed on `data-browser-action`.
 *
 * Every IFC-derived string is escaped (A1). `data-node-key` is preserved on all
 * `<details>` so the orchestrator's re-render keeps expanded state.
 */
import { escapeHtml } from '../core/markup';
import { icon, type IconName } from './icons';
import type { BrowserTreeNode, FederatedModelRecord, ModelIndex } from '../core/viewer-types';

export const MAX_BROWSER_LEVELS = 120;
export const MAX_BROWSER_CLASSES_PER_LEVEL = 28;
export const MAX_BROWSER_ELEMENTS_PER_CLASS = 26;
export const MAX_BROWSER_SPATIAL_DEPTH = 7;
export const MAX_BROWSER_SPATIAL_CHILDREN = 80;

// Maps the tree's decorative glyph names to the inline SVG icon set (U5/A2:
// no Material Symbols font). Unmapped names fall back to a neutral node icon.
const TREE_ICON: Record<string, IconName> = {
  chevron_right: 'chevron_right',
  view_in_ar: 'view_in_ar',
  subdirectory_arrow_right: 'chevron_right',
  more_horiz: 'more_horiz',
  hourglass_top: 'more_horiz',
  hide_source: 'visibility_off',
  category: 'deployed_code',
  folder_open: 'account_tree',
  account_tree: 'account_tree',
  my_location: 'center_focus_strong',
  deployed_code: 'deployed_code',
};

export const treeIco = (name: string, cls = 'browser-ico'): string =>
  `<span class="${cls}">${icon(TREE_ICON[name] ?? 'deployed_code', 16)}</span>`;

/** Intersection of a class set and a level set within one model index. */
export function getClassIdsForModelLevel(
  modelIndices: Map<string, ModelIndex>,
  modelId: string,
  level: string,
  className: string,
): Set<number> {
  const index = modelIndices.get(modelId);
  const levelIds = index?.levels.get(level);
  const classIds = index?.classes.get(className);
  if (!levelIds || !classIds) return new Set<number>();

  const result = new Set<number>();
  const source = classIds.size <= levelIds.size ? classIds : levelIds;
  const target = source === classIds ? levelIds : classIds;
  for (const localId of source) {
    if (target.has(localId)) result.add(localId);
  }
  return result;
}

/** The non-empty class buckets for a level, sorted by name. */
export function getLevelClassEntries(
  modelIndices: Map<string, ModelIndex>,
  modelId: string,
  level: string,
): Array<{ className: string; count: number }> {
  const index = modelIndices.get(modelId);
  const levelIds = index?.levels.get(level);
  if (!index || !levelIds || levelIds.size === 0) return [];

  const entries: Array<{ className: string; count: number }> = [];
  for (const className of index.classes.keys()) {
    const ids = getClassIdsForModelLevel(modelIndices, modelId, level, className);
    if (ids.size === 0) continue;
    entries.push({ className, count: ids.size });
  }
  entries.sort((a, b) => a.className.localeCompare(b.className));
  return entries;
}

function renderSpatialBrowserNode(
  modelId: string,
  index: ModelIndex,
  node: BrowserTreeNode,
  depth: number,
): string {
  const hasChildren = node.children.length > 0;
  const escapedModelId = escapeHtml(modelId);
  const categoryUpper = node.category.toUpperCase();
  const isStoreyNode = categoryUpper.includes('IFCBUILDINGSTOREY') && index.levels.has(node.label);
  const isElement = node.localId !== null && index.allIds.has(node.localId);
  const countText = node.geometryCount > 0 ? String(node.geometryCount) : (node.localId ?? '-').toString();

  if (!hasChildren || depth >= MAX_BROWSER_SPATIAL_DEPTH) {
    const leafContent = isElement && node.localId !== null
      ? `
        <span
          class="browser-action"
          data-browser-action="select-item"
          data-model-id="${escapedModelId}"
          data-local-id="${node.localId}"
          title="Select ${escapeHtml(node.label)}"
        >
          ${escapeHtml(node.label)}
        </span>
      `
      : `<span>${escapeHtml(node.label)}</span>`;
    const leafIcon = isElement ? 'view_in_ar' : 'subdirectory_arrow_right';
    return `
      <div class="browser-leaf">
        ${treeIco(leafIcon)}
        ${leafContent}
        <span class="browser-count">${countText}</span>
      </div>
    `;
  }

  const visibleChildren = node.children.slice(0, MAX_BROWSER_SPATIAL_CHILDREN);
  const childrenMarkup = visibleChildren
    .map((child) => renderSpatialBrowserNode(modelId, index, child, depth + 1))
    .join('');
  const moreMarkup = node.children.length > MAX_BROWSER_SPATIAL_CHILDREN
    ? `<div class="browser-leaf">${treeIco('more_horiz')}<span>${node.children.length - MAX_BROWSER_SPATIAL_CHILDREN} more nodes</span><span class="browser-count">+</span></div>`
    : '';

  const labelMarkup = isStoreyNode
    ? `
      <span
        class="browser-action"
        data-browser-action="isolate-level"
        data-model-id="${escapedModelId}"
        data-level="${escapeHtml(node.label)}"
        title="Isolate ${escapeHtml(node.label)}"
      >
        ${escapeHtml(node.label)}
      </span>
    `
    : `<span>${escapeHtml(node.label)}</span>`;

  const spatialKey = node.localId !== null
    ? `spatial:${escapedModelId}:id:${node.localId}`
    : `spatial:${escapedModelId}:${depth}:${escapeHtml(node.category)}:${escapeHtml(node.label)}`;
  return `
    <details class="browser-node" data-node-key="${spatialKey}">
      <summary class="browser-summary">
        ${treeIco('chevron_right', 'browser-twist')}
        ${labelMarkup}
        <span class="browser-count">${countText}</span>
      </summary>
      <div class="browser-children">
        ${childrenMarkup}
        ${moreMarkup}
      </div>
    </details>
  `;
}

/**
 * Builds the full model-browser tree markup for every loaded model. Returns
 * `null` when there are no models so the caller can render its own empty state.
 */
export function buildModelBrowserMarkup(
  federatedModels: Map<string, FederatedModelRecord>,
  modelIndices: Map<string, ModelIndex>,
): string | null {
  if (federatedModels.size === 0) return null;

  return [...federatedModels.values()]
    .map((record) => {
      const modelId = String(record.modelId);
      const escapedModelId = escapeHtml(modelId);
      const index = modelIndices.get(modelId);
      const visibilitySuffix = record.visible ? '' : ' (Hidden)';

      if (!index) {
        return `
          <details class="browser-node is-model" data-node-key="model:${escapedModelId}" open>
            <summary class="browser-summary">
              ${treeIco('chevron_right', 'browser-twist')}
              <span
                class="browser-action"
                data-browser-action="select-model"
                data-model-id="${escapedModelId}"
                title="Select full model"
              >
                ${escapeHtml(record.fileName)}${visibilitySuffix}
              </span>
              <span class="browser-count">${record.elementCount}</span>
            </summary>
            <div class="browser-children">
              <div class="browser-leaf">
                ${treeIco('hourglass_top')}
                <span>Building model tree...</span>
                <span class="browser-count">-</span>
              </div>
            </div>
          </details>
        `;
      }

      const levelEntries = [...index.levels.entries()].sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true }));
      const levelMarkup = levelEntries.slice(0, MAX_BROWSER_LEVELS).map(([levelName, ids]) => {
        const classEntries = getLevelClassEntries(modelIndices, modelId, levelName).slice(0, MAX_BROWSER_CLASSES_PER_LEVEL);
        const classMarkup = classEntries.map(({ className }) => {
          const classIds = [...getClassIdsForModelLevel(modelIndices, modelId, levelName, className)];
          classIds.sort((a, b) => {
            const aName = index.itemNames.get(a) ?? `Element ${a}`;
            const bName = index.itemNames.get(b) ?? `Element ${b}`;
            return aName.localeCompare(bName, undefined, { numeric: true });
          });
          const visibleIds = classIds.slice(0, MAX_BROWSER_ELEMENTS_PER_CLASS);
          const elementsMarkup = visibleIds.map((localId) => {
            const label = index.itemNames.get(localId) ?? `Element ${localId}`;
            return `
              <div class="browser-leaf">
                ${treeIco('view_in_ar')}
                <span
                  class="browser-action"
                  data-browser-action="select-item"
                  data-model-id="${escapedModelId}"
                  data-local-id="${localId}"
                  title="Select ${escapeHtml(label)}"
                >
                  ${escapeHtml(label)}
                </span>
                <span class="browser-count">${localId}</span>
              </div>
            `;
          }).join('');
          const hiddenCount = classIds.length - visibleIds.length;
          const moreElementsMarkup = hiddenCount > 0
            ? `<div class="browser-leaf">${treeIco('more_horiz')}<span>${hiddenCount} more elements</span><span class="browser-count">+</span></div>`
            : '';

          return `
            <details class="browser-node" data-node-key="class:${escapedModelId}:${escapeHtml(levelName)}:${escapeHtml(className)}">
              <summary class="browser-summary">
                ${treeIco('chevron_right', 'browser-twist')}
                <span
                  class="browser-action"
                  data-browser-action="isolate-class-level"
                  data-model-id="${escapedModelId}"
                  data-level="${escapeHtml(levelName)}"
                  data-class="${escapeHtml(className)}"
                  title="Isolate ${escapeHtml(className)} in ${escapeHtml(levelName)}"
                >
                  ${escapeHtml(className)}
                </span>
                <span class="browser-count">${classIds.length}</span>
              </summary>
              <div class="browser-children">
                ${elementsMarkup || `<div class="browser-leaf">${treeIco('hide_source')}<span>No elements</span><span class="browser-count">0</span></div>`}
                ${moreElementsMarkup}
              </div>
            </details>
          `;
        }).join('');

        return `
          <details class="browser-node" data-node-key="level:${escapedModelId}:${escapeHtml(levelName)}">
            <summary class="browser-summary">
              ${treeIco('chevron_right', 'browser-twist')}
              <span
                class="browser-action"
                data-browser-action="isolate-level"
                data-model-id="${escapedModelId}"
                data-level="${escapeHtml(levelName)}"
                title="Isolate level ${escapeHtml(levelName)}"
              >
                ${escapeHtml(levelName)}
              </span>
              <span class="browser-count">${ids.size}</span>
            </summary>
            <div class="browser-children">
              ${classMarkup || `<div class="browser-leaf">${treeIco('category')}<span>No classes</span><span class="browser-count">0</span></div>`}
            </div>
          </details>
        `;
      }).join('');

      const levelMoreMarkup = levelEntries.length > MAX_BROWSER_LEVELS
        ? `<div class="browser-leaf">${treeIco('more_horiz')}<span>${levelEntries.length - MAX_BROWSER_LEVELS} more levels</span><span class="browser-count">+</span></div>`
        : '';

      const spatialRootNodes = index.spatialRoot?.children?.length ? index.spatialRoot.children : (index.spatialRoot ? [index.spatialRoot] : []);
      const spatialVisible = spatialRootNodes.slice(0, MAX_BROWSER_SPATIAL_CHILDREN);
      const spatialMarkup = spatialVisible
        .map((node) => renderSpatialBrowserNode(modelId, index, node, 0))
        .join('');
      const spatialMoreMarkup = spatialRootNodes.length > MAX_BROWSER_SPATIAL_CHILDREN
        ? `<div class="browser-leaf">${treeIco('more_horiz')}<span>${spatialRootNodes.length - MAX_BROWSER_SPATIAL_CHILDREN} more nodes</span><span class="browser-count">+</span></div>`
        : '';

      return `
        <details class="browser-node is-model" data-node-key="model:${escapedModelId}" open>
          <summary class="browser-summary">
            ${treeIco('chevron_right', 'browser-twist')}
            <span
              class="browser-action"
              data-browser-action="select-model"
              data-model-id="${escapedModelId}"
              title="Select full model"
            >
              ${escapeHtml(record.fileName)}${visibilitySuffix}
            </span>
            <span class="browser-count">${record.elementCount}</span>
          </summary>
          <div class="browser-children">
            <div class="browser-leaf">
              ${treeIco('my_location')}
              <span
                class="browser-action"
                data-browser-action="fit-model"
                data-model-id="${escapedModelId}"
                title="Fit camera to model"
              >
                Default
              </span>
              <span class="browser-count">${levelEntries.length} lvls</span>
            </div>

            <details class="browser-node" data-node-key="group:${escapedModelId}:levels" ${levelEntries.length > 0 ? 'open' : ''}>
              <summary class="browser-summary">
                ${treeIco('chevron_right', 'browser-twist')}
                <span>Levels</span>
                <span class="browser-count">${levelEntries.length}</span>
              </summary>
              <div class="browser-children">
                ${levelMarkup || `<div class="browser-leaf">${treeIco('folder_open')}<span>No levels detected</span><span class="browser-count">-</span></div>`}
                ${levelMoreMarkup}
              </div>
            </details>

            <details class="browser-node" data-node-key="group:${escapedModelId}:spatial">
              <summary class="browser-summary">
                ${treeIco('chevron_right', 'browser-twist')}
                <span>Spatial Structure</span>
                <span class="browser-count">${spatialRootNodes.length}</span>
              </summary>
              <div class="browser-children">
                ${spatialMarkup || `<div class="browser-leaf">${treeIco('account_tree')}<span>No spatial tree data</span><span class="browser-count">-</span></div>`}
                ${spatialMoreMarkup}
              </div>
            </details>
          </div>
        </details>
      `;
    })
    .join('');
}
