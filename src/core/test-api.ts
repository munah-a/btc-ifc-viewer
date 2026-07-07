/**
 * Explicit, frozen contract for the e2e test hook (AUDIT T6, W2.5).
 *
 * Playwright previously reached into `window.__viewer` (the whole ViewerApp
 * instance) and its private fields — every rename silently broke tests and the
 * surface was undocumented. `window.__viewerTestApi` is a small, frozen object
 * with a stated contract: only these members are supported by e2e. It is
 * exposed ONLY in builds made with the `VITE_E2E` define (vite.e2e.config.ts);
 * plain dev/prod builds ship without it.
 *
 * This module is type-only — the concrete object is built in viewer.ts where
 * the engine/world/state live.
 */

export interface TestVec3 {
  x: number;
  y: number;
  z: number;
}

export interface TestCameraState {
  position: TestVec3;
  target: TestVec3;
}

export interface TestItemRef {
  modelId: string;
  localId: number;
}

/** The full, supported e2e surface. Frozen at build time. */
export interface ViewerTestApi {
  /** Contract version — bump on any breaking change to this shape. */
  readonly version: 1;

  // ---- state reads ----
  /** Number of federated (loaded) models. */
  modelCount(): number;
  /** Number of indexed models (index built). Equals modelCount once ready. */
  indexedModelCount(): number;
  /** First item whose indexed name contains `keyword` (case-insensitive). */
  findItemByName(keyword: string): TestItemRef | null;
  /**
   * A usable selection context from the first model: its id, a search term
   * (first named item, else first class name) and a first selectable localId.
   * Null if no model is indexed.
   */
  firstModelContext(): { modelId: string; searchTerm: string; firstItemId: number } | null;
  /** The id of the first loaded model, or null. */
  firstModelId(): string | null;
  /** All loaded model ids, in load order (C8 multi-model assertions). */
  allModelIds(): string[];
  /**
   * A class name + level name from the first model whose element sets do NOT
   * overlap (for the F11 disjoint-filter test). Null if none exists.
   */
  findDisjointClassLevel(): { className: string; levelName: string } | null;
  /** Total selected element count across all models. */
  selectionCount(): number;
  /** Whether X-ray / edges tools are on. */
  isXrayEnabled(): boolean;
  isEdgesEnabled(): boolean;
  /** The model id currently showing a transform gizmo, or null. */
  activeGizmoModelId(): string | null;
  /** Count of saved viewpoints / issues. */
  viewpointCount(): number;
  issueCount(): number;
  /** The first viewpoint's snapshot data URI (F2 thumbnail test), or null. */
  firstViewpointSnapshot(): string | null;
  /** The first issue's linked-element count (for federation assertions). */
  firstIssueLinkedCount(): number;
  /** Number of models referenced by the first issue's elementsByModel map (F9). */
  firstIssueModelCount(): number;
  /** Whether the first issue kept a legacy string modelId (BCF back-compat). */
  firstIssueHasLegacyModelId(): boolean;
  /**
   * Engine model bookkeeping after loads/unloads (F6): fragment list size,
   * federated-record count, index count and scene-object count.
   */
  engineModelState(): {
    fragmentsCount: number;
    federatedCount: number;
    indexCount: number;
    objectCount: number;
  };
  /** Current camera position + target in world space. */
  cameraState(): TestCameraState;
  /**
   * The world-space direction the camera should look FROM for a view-cube local
   * vector, using the current anchor model's basis (mirrors the app's cube
   * navigation math). Null if no model is loaded. Used by cube-nav assertions.
   */
  anchorDirectionForCube(localVector: readonly [number, number, number]): TestVec3 | null;

  // ---- actions ----
  /** Select a single element (optionally zoom to it). Resolves when applied. */
  selectItem(modelId: string, localId: number, zoom?: boolean): Promise<void>;
  /**
   * Select the first indexed element of EVERY loaded model (F9 multi-model
   * selection). Resolves after selection visuals settle.
   */
  selectFirstItemPerModel(): Promise<void>;
  /** Set the visual style (bypassing status/persist side effects for tests). */
  setVisualStyle(style: string): Promise<void>;
  /**
   * Run the canvas pick+select path at the given client coordinates (T6: real
   * interaction coverage — drives the same pickAndSelect() a pointer click
   * does). Resolves after the selection settles. Returns the item hit, or null
   * (a miss in single-select mode clears the current selection, exactly like a
   * real empty-space click). NOTE: a positive raycast HIT depends on live GPU
   * render state and is not reliable under headless software WebGL — tests
   * assert the clear-on-empty-click behaviour, not a specific hit.
   */
  clickCanvasAt(clientX: number, clientY: number): Promise<TestItemRef | null>;
  /**
   * W4.5: exports the current visibility state to a binary GLB and returns
   * `{ byteLength, valid }` (valid = well-formed GLB header) WITHOUT triggering a
   * file download — so the e2e can assert a non-empty, valid .glb is produced.
   */
  exportGlbBytes(): Promise<{ byteLength: number; valid: boolean }>;

  // ---- C8 full-session persistence (W5.2) ----
  /**
   * The per-model modifications currently persisted for the given model id
   * (transform offsets, opacity, visibility), read straight from the live
   * federation record — so the e2e can assert modifications survived a reload.
   * Null when no such model is loaded.
   */
  modelModifications(modelId: string): {
    offsetPosition: TestVec3;
    offsetRotation: TestVec3;
    opacity: number;
    visible: boolean;
    fragKey: string | null;
  } | null;
  /** How many models the saved session in localStorage would restore. */
  persistedModelCount(): number;
  /** Applies the given per-model opacity (0–1) via the same path as the slider. */
  setModelOpacity(modelId: string, opacity: number): void;
  /** Applies a translation offset (metres) to a model via the transform path. */
  setModelOffset(modelId: string, x: number, y: number, z: number): void;
  /** Persists the current session immediately (the explicit Save affordance). */
  saveSession(): void;
  /**
   * W5-fixups: monotonic count of requestRender() calls. In MANUAL-mode
   * on-demand rendering a visual mutation must re-arm a frame; this lets an e2e
   * assert render parity (e.g. a measurement/section tool action bumped it).
   */
  renderRequestCount(): number;
}

declare global {
  interface Window {
    __viewerTestApi?: ViewerTestApi;
  }
}
