/**
 * Federation panel (Models tab) — pure markup builder.
 *
 * Builds the per-model cards (name, visibility, opacity slider, offset/rotation
 * transform inputs, action buttons, level list) from the federation records.
 * DOM-free: the orchestrator writes the returned markup into `#federationTree`
 * and handles interaction via delegated listeners keyed on `data-model-action`,
 * `data-model-opacity`, `data-transform`, `data-level` (bound once).
 *
 * Every IFC-derived / model-id string is escaped (A1).
 */
import { escapeHtml } from '../core/markup';
import type { FederatedModelRecord, ModelIndex } from '../core/viewer-types';

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const formatModelSize = (sizeBytes: number): string => {
  if (!sizeBytes || sizeBytes <= 0) return '-';
  return `${(sizeBytes / 1024 / 1024).toFixed(1)} MB`;
};

/**
 * Already-translated federation-panel labels (C7). Passed in by the caller so
 * this pure builder never imports the i18n catalog.
 */
export interface FederationLabels {
  show: string;
  hide: string;
  noStoreys: string;
  opacity: string;
  offsetXyz: string;
  rotationXyz: string;
  select: string;
  gizmo: string;
  fit: string;
  reset: string;
  unload: string;
  selectFullModel: string;
  fitCamera: string;
  unloadTitle: string;
  /** `{count} elements` for the meta line — pre-interpolated. */
  elements(count: number): string;
  /** `Isolate level {level}` — pre-interpolated. */
  isolateLevel(level: string): string;
  /** `Levels ({count})` — pre-interpolated. */
  levelsCount(count: number): string;
}

/**
 * Builds the federation cards markup for every loaded model. Returns `null`
 * when there are no models so the caller can render its own empty state.
 */
export function buildFederationTreeMarkup(
  federatedModels: Map<string, FederatedModelRecord>,
  modelIndices: Map<string, ModelIndex>,
  activeGizmoModelId: string | null,
  labels: FederationLabels,
): string | null {
  if (federatedModels.size === 0) return null;

  return [...federatedModels.values()]
    .map((record) => {
      const modelId = String(record.modelId);
      const escapedModelId = escapeHtml(modelId);
      const opacityPct = Math.round(clamp(record.opacity, 0, 1) * 100);
      const visibilityLabel = escapeHtml(record.visible ? labels.hide : labels.show);
      const visibilityStateClass = record.visible ? '' : 'is-off';
      const gizmoStateClass = activeGizmoModelId === modelId ? 'is-active' : '';
      const levels = modelIndices.get(record.modelId)?.levels;
      const levelEntries = levels ? [...levels.entries()] : [];
      const levelMarkup = levelEntries.length === 0
        ? `<div class="federated-level-empty">${escapeHtml(labels.noStoreys)}</div>`
        : levelEntries
          .slice(0, 80)
          .map(([levelName, ids]) => `
            <button
              class="federated-level-btn"
              type="button"
              data-model-id="${escapedModelId}"
              data-level="${escapeHtml(levelName)}"
              title="${escapeHtml(labels.isolateLevel(levelName))}"
            >
              ${escapeHtml(levelName)} (${ids.size})
            </button>
          `)
          .join('');

      return `
        <div class="federated-model">
          <div class="federated-model-header">
            <div class="federated-header-row">
              <button
                class="federated-model-name-btn"
                type="button"
                data-model-id="${escapedModelId}"
                data-model-action="select-model"
                title="${escapeHtml(labels.selectFullModel)}"
              >
                ${escapeHtml(record.fileName)}
              </button>
              <button
                class="federated-model-btn federated-visibility-btn ${visibilityStateClass}"
                type="button"
                data-model-id="${escapedModelId}"
                data-model-action="toggle-visibility"
              >
                ${visibilityLabel}
              </button>
            </div>
            <div class="federated-model-meta">
              ${escapedModelId} | ${escapeHtml(labels.elements(record.elementCount))} | ${formatModelSize(record.sizeBytes)}
            </div>
          </div>

          <div class="federated-opacity">
            <div class="federated-opacity-head">
              <span>${escapeHtml(labels.opacity)}</span>
              <span data-opacity-value>${opacityPct}%</span>
            </div>
            <input
              type="range"
              min="0"
              max="100"
              step="1"
              value="${opacityPct}"
              data-model-id="${escapedModelId}"
              data-model-opacity="1"
            />
          </div>

          <div class="federated-transform-title">${escapeHtml(labels.offsetXyz)}</div>
          <div class="federated-transform-grid">
            <label>X<input type="number" step="0.1" value="${record.offsetPosition.x.toFixed(2)}" data-model-id="${escapedModelId}" data-transform="px" /></label>
            <label>Y<input type="number" step="0.1" value="${record.offsetPosition.y.toFixed(2)}" data-model-id="${escapedModelId}" data-transform="py" /></label>
            <label>Z<input type="number" step="0.1" value="${record.offsetPosition.z.toFixed(2)}" data-model-id="${escapedModelId}" data-transform="pz" /></label>
          </div>

          <div class="federated-transform-title">${escapeHtml(labels.rotationXyz)}</div>
          <div class="federated-transform-grid">
            <label>Rx<input type="number" step="1" value="${record.offsetRotation.x.toFixed(1)}" data-model-id="${escapedModelId}" data-transform="rx" /></label>
            <label>Ry<input type="number" step="1" value="${record.offsetRotation.y.toFixed(1)}" data-model-id="${escapedModelId}" data-transform="ry" /></label>
            <label>Rz<input type="number" step="1" value="${record.offsetRotation.z.toFixed(1)}" data-model-id="${escapedModelId}" data-transform="rz" /></label>
          </div>

          <div class="federated-model-actions">
            <button class="federated-model-btn" type="button" data-model-id="${escapedModelId}" data-model-action="select-model">${escapeHtml(labels.select)}</button>
            <button class="federated-model-btn ${gizmoStateClass}" type="button" data-model-id="${escapedModelId}" data-model-action="toggle-gizmo">${escapeHtml(labels.gizmo)}</button>
            <button class="federated-model-btn" type="button" data-model-id="${escapedModelId}" data-model-action="fit">${escapeHtml(labels.fit)}</button>
            <button class="federated-model-btn" type="button" data-model-id="${escapedModelId}" data-model-action="reset">${escapeHtml(labels.reset)}</button>
            <button class="federated-model-btn federated-unload-btn" type="button" data-model-id="${escapedModelId}" data-model-action="unload" title="${escapeHtml(labels.unloadTitle)}">${escapeHtml(labels.unload)}</button>
          </div>

          <div class="federated-levels">
            <details>
              <summary>${escapeHtml(labels.levelsCount(levelEntries.length))}</summary>
              <div class="federated-level-list">${levelMarkup}</div>
            </details>
          </div>
        </div>
      `;
    })
    .join('');
}
