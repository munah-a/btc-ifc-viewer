import * as THREE from 'three';
import type * as OBC from '@thatopen/components';
import type * as OBCF from '@thatopen/components-front';
import type { TransformControls } from 'three/examples/jsm/controls/TransformControls.js';

import { isModelNotFoundError, serializeError } from './core/errors';
import { isProbablyIfc } from './core/ifc-format';
import { escapeHtml, filterChipMarkup } from './core/markup';
import { buildModelIndex } from './core/model-index';
import { ModelRegistry } from './core/model-registry';
import {
  buildPersistedState,
  normalizePersistedState,
  type NavigationMode,
  type PersistedCamera,
  type PersistedModelRecord,
  type PersistedViewerState,
  type SavedViewpoint,
  type SelectionMode,
  type Vector3Record,
  type VisualStyle,
} from './core/persistence';
import {
  hashFragBytes,
  IndexedDbFragCache,
  type FragCacheAdapter,
} from './core/frag-cache';
import { IfcConversionClient } from './core/ifc-conversion-client';
import { registerServiceWorker } from './core/pwa';
import { DEFAULT_MODEL_UNITS, resolveModelUnits, type ModelUnits } from './core/units';
import {
  clearMap,
  cloneMap,
  countMapItems,
  intersectMaps,
  isMapEmpty,
  toSetMap,
} from './core/model-id-map';
import {
  buildPropertySections,
  readPrimitiveValue,
  toPropertyString,
  type GeometryProbe,
  type PropertySectionData,
} from './core/property-engine';
import { getClipperPlaneGizmoHelper, type FragmentsModelLike } from './core/fragments-model';
import {
  t,
  setLanguage,
  getLanguage,
  initLanguage,
  hydrateI18n,
  onLanguageChange,
  formatDateTime,
  type Language,
} from './core/i18n';
import { ShaderWarningFilter } from './core/engine-lite';
// The engine module (core/viewer-core) statically pulls three + @thatopen +
// web-ifc (~1MB gzip). It is dynamically imported in initEngine (W5.1 / P1) so
// the initial app shell excludes it — type-only imports here are erased.
import type { BootstrapEngineOptions, EngineHandles } from './core/viewer-core';
type ViewerCoreModule = typeof import('./core/viewer-core');
import { exportModelsToGlb, hasActiveClipping, isValidGlb } from './core/glb-export';
import { buildShareUrl, decodeUrlState, isAllowedModelUrl, type UrlViewpointState } from './core/url-state';
import { UploadClient } from './core/upload-client';
import { ShareDialogController } from './ui/share-dialog';
import type { TestItemRef, ViewerTestApi } from './core/test-api';
import { getViewCubeAxes, getViewCubeNavigationDistance, resolveViewCubeCameraUp } from './core/view-cube';
import { buildEdgeOverlays, EdgeGeometryCache } from './tools/edges';
import { routeKeyboardEvent, type KeyboardActions } from './input/keyboard-router';
import { sectionBoxPlanes, sectionPlanePoint } from './tools/section';
import { computeXrayOpacity } from './tools/xray';
import type {
  FederatedModelRecord,
  IssueRecord,
  MeasureMode,
  ModelIndex,
  SearchResult,
  TransformVector3,
} from './core/viewer-types';
import { createDomCache, type ViewerDom } from './ui/dom-cache';
import { buildMobileSheet as buildMobileSheetView } from './ui/mobile-sheet';
import { hydrateIcons, setIcon, type IconName } from './ui/icons';
import { buildFederationTreeMarkup } from './ui/federation-panel';
import { buildIssueCommentsMarkup, buildIssueListMarkup } from './ui/issues-panel';
import {
  buildModelBrowserMarkup,
  getClassIdsForModelLevel,
  type BrowserLabels,
} from './ui/model-browser';
import { buildPropertySectionsMarkup } from './ui/properties-panel';
import { buildViewpointListMarkup } from './ui/viewpoints-panel';

// Shared runtime types (ModelIndex, BrowserTreeNode, FederatedModelRecord,
// IssueRecord, SearchResult, MeasureMode) live in core/viewer-types.ts so the
// extracted ui/* panel controllers can import them without pulling in the engine.

const STORAGE_KEY = 'bim_for_field_viewer_state_v1';
const DEFAULT_BACKGROUND_COLOR = '#0b1220';
const DEFAULT_LIGHT_BACKGROUND_COLOR = '#c6d5e8';
// F2: viewpoint thumbnails are downscaled JPEGs; anything bigger than this
// (i.e. a full-resolution capture) never enters the localStorage payload.
const VIEWPOINT_THUMBNAIL_MAX_DIM = 320;
const VIEWPOINT_THUMBNAIL_JPEG_QUALITY = 0.72;
const MAX_PERSISTED_SNAPSHOT_CHARS = 150_000;
// Model-browser tree caps (MAX_BROWSER_*) live in ui/model-browser.ts.
// The interactive 3D view-cube widget was replaced (W3, design) by the glass
// view-control buttons (fit / orbit-home / front / top). The camera-direction
// math is retained: preset views + the anchor basis power the buttons and the
// `anchorDirectionForCube` test hook.
const VIEW_CUBE_HOME_VECTOR: readonly [number, number, number] = [1, 1, 1];
const FRONT_VIEW_VECTOR: readonly [number, number, number] = [0, 0, 1];
const TOP_VIEW_VECTOR: readonly [number, number, number] = [0, 1, 0];

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const uniqueId = (): string => {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
};

const downloadBlob = (name: string, blob: Blob): void => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
};

const debounce = <T extends (...args: any[]) => void>(fn: T, ms: number): T => {
  let timer: ReturnType<typeof setTimeout>;
  return ((...args: any[]) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  }) as unknown as T;
};


class ViewerApp {
  private readonly abortController = new AbortController();
  private fpsMonitor: { stop: () => void } | null = null;

  private get signal(): AbortSignal {
    return this.abortController.signal;
  }

  private readonly dom: ViewerDom = createDomCache();

  private components!: OBC.Components;
  private world!: OBC.SimpleWorld<OBC.SimpleScene, OBC.OrthoPerspectiveCamera, OBCF.PostproductionRenderer>;
  // NOTE: the IFC parse no longer runs through OBC.IfcLoader in the full app —
  // P4 (W5.4) moved conversion to ifcConversionClient (a dedicated worker). The
  // engine still bootstraps IfcLoader (shared with the embed's ?m=.ifc path).
  private fragments!: OBC.FragmentsManager;
  private clipper!: OBC.Clipper;
  private hider!: OBC.Hider;
  private raycaster!: ReturnType<OBC.Raycasters['get']>;
  private lengthMeasurement!: OBCF.LengthMeasurement;
  private areaMeasurement!: OBCF.AreaMeasurement;
  private markerManager!: OBCF.Marker;
  private transformControls: TransformControls | null = null;
  private transformControlsHelper: THREE.Object3D | null = null;

  private readonly selectedItems: OBC.ModelIdMap = {};
  private selectionMode: SelectionMode = 'single';
  private measureMode: MeasureMode = 'none';
  private navigationMode: NavigationMode = 'Orbit';
  private xrayEnabled = false;
  private edgesEnabled = false;
  private visualStyle: VisualStyle = 'color-pen-shadows';
  private issuePinMode = false;
  // F4: grid defaults off — the single source of truth (the old 500 ms
  // checkbox-reset hack in index.html raced init and clobbered persisted
  // preferences and the status bar).
  private gridVisible = false;
  private themeMode: 'dark' | 'light' = 'dark';
  private backgroundColor = DEFAULT_BACKGROUND_COLOR;
  // F8: per-theme background memory — the theme toggle swaps stored values
  // instead of force-overwriting a custom background.
  private backgroundByTheme: Record<'dark' | 'light', string> = {
    dark: DEFAULT_BACKGROUND_COLOR,
    light: DEFAULT_LIGHT_BACKGROUND_COLOR,
  };
  private gridHelper: THREE.Object3D | null = null;
  private readonly appliedModelOpacity = new Map<string, number>();
  // C8 (W5.2): per-model hidden element ids, refreshed from the hider whenever
  // visibility changes, so the synchronous session serializer can persist them.
  private readonly hiddenIdsByModel = new Map<string, number[]>();

  private edgeOverlays: THREE.LineSegments[] = [];
  private readonly edgeMaterial = new THREE.LineBasicMaterial({ color: 0xc8145c, transparent: true, opacity: 0.65 });
  // P3 (W5.3): cache EdgesGeometry per source geometry so applyEdges doesn't
  // rebuild O(triangles) edges on every opacity/gizmo tick — reposition matrices.
  private readonly edgeGeometryCache = new EdgeGeometryCache();
  private modelObjects: THREE.Object3D[] = [];
  private readonly federatedModels = new Map<string, FederatedModelRecord>();
  private modelIndices = new Map<string, ModelIndex>();
  // F5: per-model display units resolved from the model's unit entities;
  // the properties panel uses the selected element's model units.
  private readonly modelUnits = new Map<string, ModelUnits>();
  private activePropertyUnits: ModelUnits = DEFAULT_MODEL_UNITS;
  // Load-lifecycle bookkeeping (A6/A10): metadata keyed by the model id passed
  // to ifcLoader.load; stale ids track timed-out loads for late disposal.
  private readonly modelRegistry = new ModelRegistry();
  // One registration promise per model id — dedupes the onModelLoaded event
  // and the awaited load path without polling.
  private readonly modelRegistrations = new Map<string, Promise<void>>();
  // W4.4: share/host — the hosting API client, the dialog controller, and the
  // hosted `.frag` URL per model id (set after a publish; enables deep links).
  private readonly uploadClient = new UploadClient();
  private shareController: ShareDialogController | null = null;
  private readonly hostedModelUrls = new Map<string, string>();
  // C8 (W5.2): IndexedDB cache of converted `.frag` bytes for full-session
  // restore. Injectable so unit tests can swap in an in-memory adapter.
  private fragCache: FragCacheAdapter = new IndexedDbFragCache();
  // P4 (W5.4): IFC→fragments conversion runs in a dedicated worker so the UI
  // stays responsive on big models (the @thatopen fragments worker only streams;
  // it does NOT parse IFC).
  private readonly ifcConversionClient = new IfcConversionClient();
  // Guards persistLocalState during a restore so re-applying state doesn't churn
  // localStorage, and suppresses the auto-fit/status spam while models stream in.
  private restoringSession = false;
  // C8: fragKeys for models loaded FROM the cache during restore — onModelAdded
  // reuses these instead of re-hashing/re-caching bytes it just read from IDB.
  private readonly restoreFragKeys = new Map<string, string>();
  private lastHitPoint: THREE.Vector3 | null = null;
  private pendingIssuePoint: THREE.Vector3 | null = null;
  // U4: failed loads are kept so the overlay's Retry can replay them.
  private lastFailedLoadFiles: File[] = [];
  private viewpoints: SavedViewpoint[] = [];
  private selectedViewpointId: string | null = null;
  private issues: IssueRecord[] = [];
  private activeIssueId: string | null = null;
  private lastPointerDown = { x: 0, y: 0 };
  private pointerDragged = false;
  private isModelLoading = false;
  private loadRequestId = 0;
  private suppressAutoFit = false;
  private activeGizmoModelId: string | null = null;
  private gizmoDragging = false;
  private propertyFilterText = '';
  private activeView: 'orbit' | 'front' | 'top' = 'orbit';
  private activeTab = 'explorer';
  // A5: scoped console.warn filter (install/uninstall paired; restored in
  // destroy() rather than leaving console.warn permanently monkey-patched).
  private readonly shaderWarningFilter = new ShaderWarningFilter();
  // W5.1: the dynamically-imported engine module (three/@thatopen/web-ifc live
  // here). Loaded once in initEngine so its enum/helper values are reusable.
  private engineModule!: ViewerCoreModule;
  // P6 (W5.3): on-demand rendering. `requestRender` re-arms the MANUAL-mode
  // PostproductionRenderer on any visual change; `onDemand.stop` detaches on
  // destroy. requestRender is a no-op until initEngine wires it.
  private requestRender: () => void = () => {};
  private onDemandRender: { requestRender: () => void; stop: () => void } | null = null;
  private engineHandles: EngineHandles | null = null;
  // R1 (W5-fixups): while a section-plane gizmo is dragged the camera controls
  // are disabled, so the camera 'update' render pump does NOT fire and the
  // re-clipped view freezes. A per-frame rAF pump (armed on the clipper's
  // onBeforeDrag, cancelled on onAfterDrag) calls requestRender every frame so
  // the clipped geometry repaints live during the drag. 0 = not running.
  private sectionDragRaf = 0;
  // W5-fixups: monotonically-incrementing count of requestRender() calls, so an
  // e2e can assert MANUAL-mode render parity (a visual mutation re-armed a
  // frame). Only observable via the VITE_E2E test API.
  private renderRequestCount = 0;

  constructor() {
    // C7: resolve the persisted language (default EN) and localize the static
    // shell before any panels render, so first paint is already in-language.
    initLanguage();
    hydrateIcons(document);
    hydrateI18n(document);
    this.syncLanguageUi(getLanguage());
    this.patchEventListenersWithAbort();
    this.bindUiEvents();
    // Re-render JS-built panels + counters whenever the language changes
    // (hydrateI18n handles the static [data-i18n] shell; this covers the
    // dynamically-generated markup that hydrateI18n cannot reach).
    onLanguageChange((language) => this.onLanguageChanged(language));
  }

  /** Reflects the active language on the toggle control (shows EN / DE). */
  private syncLanguageUi(language: Language): void {
    this.dom.langCode.textContent = language.toUpperCase();
  }

  /**
   * Called after setLanguage(): the static shell was already re-hydrated by the
   * i18n module; here we re-render everything built in JS (panels, counters,
   * status label, view/mode labels, tab title) so nothing stays stale, and we
   * persist the choice into the viewer state (C8 restore).
   */
  private onLanguageChanged(language: Language): void {
    this.syncLanguageUi(language);
    this.renderModelBrowser();
    this.renderFederatedTree();
    this.updateCounters();
    this.updateTopbarModel();
    this.renderClassFilters();
    this.renderLevelFilters();
    this.updateViewpointList();
    this.updateIssuesList();
    this.updateIssueComments();
    void this.updatePropertiesPanel();
    this.updateActiveTabTitle();
    this.applyNavigationMode(this.navigationMode);
    this.setActiveView(this.activeView);
    this.syncMeasureHint();
    this.syncMobileSheet();
    // Refresh the current status line into the new language where we can.
    this.setStatus(t('status.languageChanged'));
    this.persistLocalState();
  }

  /** Toggles EN⇄DE (the top-bar language control). */
  private toggleLanguage(): void {
    setLanguage(getLanguage() === 'en' ? 'de' : 'en');
  }

  /**
   * Patches addEventListener on all cached DOM elements so that every listener
   * is automatically tied to this.abortController.signal.
   * Calling destroy() aborts the controller and removes all listeners at once.
   */
  private patchEventListenersWithAbort(): void {
    const signal = this.signal;
    const patchTarget = (el: EventTarget): void => {
      const original = el.addEventListener.bind(el);
      (el as any).addEventListener = (
        type: string,
        listener: EventListenerOrEventListenerObject | null,
        options?: boolean | AddEventListenerOptions,
      ) => {
        if (typeof options === 'boolean') {
          original(type, listener, { capture: options, signal });
        } else {
          original(type, listener, { ...options, signal });
        }
      };
    };

    for (const value of Object.values(this.dom)) {
      if (value instanceof EventTarget) {
        patchTarget(value);
      } else if (Array.isArray(value)) {
        value.forEach((el) => { if (el instanceof EventTarget) patchTarget(el); });
      }
    }
  }

  /**
   * Tears down the viewer: aborts all event listeners, cancels animation frames,
   * disposes THREE.js resources, and clears data structures.
   */
  destroy(): void {
    // 1. Abort all event listeners at once
    this.abortController.abort();

    // 2. Cancel animation frames + detach the on-demand render listeners (P6)
    this.fpsMonitor?.stop();
    this.fpsMonitor = null;
    this.onDemandRender?.stop();
    this.onDemandRender = null;
    // R1 (W5-fixups): cancel any in-flight section-drag render pump.
    this.stopSectionDragPump();
    // P4 (W5.4): free the IFC-conversion worker.
    this.ifcConversionClient.terminate();

    // 3. Dispose THREE.js resources
    this.edgeMaterial.dispose();
    // P3: overlay geometries are owned by the cache — remove overlays but let
    // the cache free the shared EdgesGeometry.
    this.edgeOverlays = [];
    this.edgeGeometryCache.dispose();

    // Dispose transform controls
    if (this.transformControls) {
      this.transformControls.dispose();
      this.transformControls = null;
    }

    // 4. Clear data structures
    this.federatedModels.clear();
    this.modelIndices.clear();
    this.modelUnits.clear();
    this.modelRegistry.clear();
    this.modelRegistrations.clear();
    this.appliedModelOpacity.clear();
    clearMap(this.selectedItems);
    this.viewpoints = [];
    this.issues = [];
    this.modelObjects = [];

    // 5. Dispose components
    if (this.components) {
      this.components.dispose();
    }

    // 6. Restore the patched console.warn (A5).
    this.shaderWarningFilter.uninstall();
  }

  async init(): Promise<void> {
    try {
      this.setStatus(t('status.initializing'));
      this.shaderWarningFilter.install();
      await this.initEngine();
      this.initShareDialog();
      await this.restoreLocalState();
      this.syncVisualSettingsUi();
      this.applySelectionMode(this.selectionMode);
      this.applyNavigationMode(this.navigationMode);
      this.renderModelBrowser();
      this.renderFederatedTree();
      this.updateCounters();
      this.updateIssuesList();
      this.updateIssueComments();
      this.updateViewpointList();
      this.setStatus(t('status.ready'));
      // P6 (W5.3): NOW switch to on-demand rendering. Postproduction has been
      // configured (setVisualStyle, above) and the AUTO loop that started in
      // components.init() has rendered a few frames, so the composer/basePass is
      // built — safe to go MANUAL. Wait two rAFs to be sure a frame landed.
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      );
      this.enableOnDemandRendering();
      this.startFpsMonitor();
      // W5.5 (C6): register the PWA service worker so the shell + self-hosted
      // wasm/worker/fonts precache and the app works offline.
      registerServiceWorker();
      // W4.2: if the app was opened with ?m=<url>, load that hosted model and
      // apply ?vp=. Errors are surfaced via the normal load-error path.
      this.fireAndForget(this.loadFromUrlParams(), 'Load model from URL');
    } catch (error) {
      this.showToast(t('status.initFailed', { error: serializeError(error) }), 'error', 8000);
      this.setStatus(t('status.initFailed', { error: serializeError(error) }));
      console.error(error);
    }
  }

  private bindUiEvents(): void {
    this.dom.btnUpload.addEventListener('click', () => this.dom.fileInput.click());
    this.dom.btnUploadEmpty.addEventListener('click', () => this.dom.fileInput.click());
    this.dom.fileInput.addEventListener('change', (event) => {
      const files = Array.from((event.target as HTMLInputElement).files ?? []);
      if (files.length === 0) return;
      this.fireAndForget(this.loadIfcFiles(files), 'Load IFC files');
      this.dom.fileInput.value = '';
    });

    this.dom.btnRetryLoad.addEventListener('click', () => {
      const files = this.lastFailedLoadFiles;
      this.lastFailedLoadFiles = [];
      this.hideLoadError();
      if (files.length > 0) this.fireAndForget(this.loadIfcFiles(files), 'Load IFC files');
    });
    this.dom.btnDismissLoadError.addEventListener('click', () => {
      this.lastFailedLoadFiles = [];
      this.hideLoadError();
    });

    this.dom.btnShare.addEventListener('click', () => this.shareController?.open());
    this.dom.btnExportGlb.addEventListener('click', () => this.fireAndForget(this.exportGlb(), 'Export GLB'));
    this.dom.btnExportScreenshot.addEventListener('click', () => this.exportScreenshot());
    this.dom.btnExportState.addEventListener('click', () => this.exportViewerState());
    this.dom.btnImportState.addEventListener('click', () => this.dom.importStateInput.click());
    // C8 (W5.2): explicit save/restore-session controls (the session also
    // auto-saves on every change and auto-restores on boot).
    this.dom.btnSaveSession.addEventListener('click', () => this.saveSession());
    this.dom.btnRestoreSession.addEventListener('click', () => this.fireAndForget(this.restoreSavedSession(), 'Restore session'));
    this.dom.importStateInput.addEventListener('change', (event) => {
      const file = (event.target as HTMLInputElement).files?.[0];
      if (!file) return;
      this.fireAndForget(this.importViewerState(file), 'Import viewer state');
      this.dom.importStateInput.value = '';
    });

    this.dom.propFilterInput.addEventListener('input', debounce(() => {
      this.propertyFilterText = this.dom.propFilterInput.value.trim().toLowerCase();
      this.applyPropertiesFilter();
    }, 300));
    this.bindTabs();
    this.bindResponsiveShell();
    this.bindModelBrowserEvents();
    this.bindFederationTreeEvents();
    this.bindViewpointListEvents();
    this.bindIssueListEvents();

    // View controls (top-right glass) + nav-mode pill
    this.dom.btnModeOrbit.addEventListener('click', () => this.applyNavigationMode('Orbit'));
    this.dom.btnModePlan.addEventListener('click', () => this.applyNavigationMode('Plan'));
    this.dom.btnModeFirstPerson.addEventListener('click', () => this.applyNavigationMode('FirstPerson'));
    this.dom.btnFitAll.addEventListener('click', () => this.fitToModel());
    this.dom.btnFront.addEventListener('click', () => this.setFrontView());
    this.dom.btnTop.addEventListener('click', () => this.setTopView());
    this.dom.cubeHome.addEventListener('click', () => this.setHomeView());

    // Tool rail — selection
    this.dom.btnSelectSingle.addEventListener('click', () => this.applySelectionMode('single'));
    this.dom.btnSelectMulti.addEventListener('click', () => this.applySelectionMode('multi'));
    this.dom.btnIsolate.addEventListener('click', () => this.isolateSelection());
    this.dom.btnHide.addEventListener('click', () => this.hideSelection());
    this.dom.btnResetVisibility.addEventListener('click', () => this.resetVisibility());
    this.dom.btnPropsIsolate.addEventListener('click', () => this.isolateSelection());
    this.dom.btnPropsHide.addEventListener('click', () => this.hideSelection());

    // Tool rail — sections (A12: one path per axis)
    this.dom.btnSectionX.addEventListener('click', () => this.toggleSectionPlane(this.dom.btnSectionX, new THREE.Vector3(-1, 0, 0)));
    this.dom.btnSectionY.addEventListener('click', () => this.toggleSectionPlane(this.dom.btnSectionY, new THREE.Vector3(0, 0, -1)));
    this.dom.btnSectionZ.addEventListener('click', () => this.toggleSectionPlane(this.dom.btnSectionZ, new THREE.Vector3(0, -1, 0)));
    this.dom.btnSectionBox.addEventListener('click', () => {
      if (this.dom.btnSectionBox.classList.contains('is-active')) {
        this.clearSections();
      } else {
        this.createSectionBox();
        this.dom.btnSectionBox.classList.add('is-active');
        this.dom.btnSectionBox.setAttribute('aria-pressed', 'true');
      }
    });
    this.dom.btnClearSections.addEventListener('click', () => {
      this.clearSections();
      this.setStatus(t('status.sectionsCleared'));
    });
    this.dom.sectionPos.addEventListener('input', () => this.onSectionSliderInput());
    this.dom.btnClearSectionSlider.addEventListener('click', () => {
      this.clearSections();
      this.setStatus(t('status.sectionsCleared'));
    });

    // Tool rail — measure
    this.dom.btnMeasureLength.addEventListener('click', () => this.setMeasureMode(this.measureMode === 'length' ? 'none' : 'length'));
    this.dom.btnMeasureArea.addEventListener('click', () => this.setMeasureMode(this.measureMode === 'area' ? 'none' : 'area'));
    this.dom.btnClearMeasurements.addEventListener('click', () => {
      this.clearMeasurements();
      this.setStatus(t('status.measurementsCleared'));
    });
    this.dom.btnCancelMeasure.addEventListener('click', () => this.setMeasureMode('none'));

    // Tool rail — visual toggles
    this.dom.btnTransparency.addEventListener('click', () => this.toggleXray());
    this.dom.btnWireframe.addEventListener('click', () => this.toggleEdges());
    this.dom.btnToggleGrid.addEventListener('click', () => {
      this.setGridVisible(!this.gridVisible, true);
      this.persistLocalState();
    });
    this.dom.btnIssuePinMode.addEventListener('click', () => this.toggleIssuePinMode());

    // Selection chip
    this.dom.btnClearSelection.addEventListener('click', () => {
      this.fireAndForget(this.clearSelection(), 'Clear selection');
    });

    // Theme toggle (F8: per-theme background memory)
    this.dom.btnThemeToggle.addEventListener('click', () => this.toggleTheme());
    this.dom.btnLangToggle.addEventListener('click', () => this.toggleLanguage());

    // Search — live on input (design has no explicit Go button)
    const runSearch = debounce(() => {
      const term = this.dom.searchInput.value.trim();
      this.dom.btnClearSearch.hidden = term.length === 0;
      this.fireAndForget(this.searchElements(term), 'Search');
    }, 300);
    this.dom.searchInput.addEventListener('input', runSearch);
    this.dom.btnClearSearch.addEventListener('click', () => {
      this.dom.searchInput.value = '';
      this.dom.btnClearSearch.hidden = true;
      this.dom.elementResults.replaceChildren();
      this.dom.searchResultsGroup.hidden = true;
      this.setStatus(t('status.searchCleared'));
    });

    // Filters — chip toggles apply immediately
    this.dom.classFilterList.addEventListener('click', (event) => this.onFilterChipClick(event));
    this.dom.levelFilterList.addEventListener('click', (event) => this.onFilterChipClick(event));

    // Viewpoints / issues
    this.dom.btnSaveViewpoint.addEventListener('click', () => {
      this.fireAndForget(this.saveViewpoint(), 'Save viewpoint');
    });
    this.dom.btnCreateIssue.addEventListener('click', () => this.createIssueFromCurrentContext());
    this.dom.btnDeleteIssue.addEventListener('click', () => {
      this.fireAndForget(this.deleteSelectedIssue(), 'Delete issue');
    });
    this.dom.btnAddIssueComment.addEventListener('click', () => this.addCommentToActiveIssue());

    // Window + viewport pointer
    window.addEventListener('resize', () => {
      if (this.world?.renderer) this.world.renderer.resize();
    });
    window.addEventListener('keydown', (event) => this.onKeyDown(event));

    this.dom.viewerContainer.addEventListener('pointerdown', (event) => {
      this.lastPointerDown = { x: event.clientX, y: event.clientY };
      this.pointerDragged = false;
    });
    this.dom.viewerContainer.addEventListener('pointermove', (event) => {
      const dx = event.clientX - this.lastPointerDown.x;
      const dy = event.clientY - this.lastPointerDown.y;
      if (Math.hypot(dx, dy) > 5) this.pointerDragged = true;
    });
    this.dom.viewerContainer.addEventListener('click', (event) => {
      this.fireAndForget(this.onViewerClick(event), 'Selection');
    });
    this.dom.viewerContainer.addEventListener('dragover', (event) => {
      event.preventDefault();
      this.dom.viewerContainer.classList.add('is-dragover');
    });
    this.dom.viewerContainer.addEventListener('dragleave', () => {
      this.dom.viewerContainer.classList.remove('is-dragover');
    });
    this.dom.viewerContainer.addEventListener('drop', (event) => {
      event.preventDefault();
      this.dom.viewerContainer.classList.remove('is-dragover');
      const files = Array.from(event.dataTransfer?.files ?? []);
      if (files.length === 0) return;
      const ifcFiles = files.filter((file) => file.name.toLowerCase().endsWith('.ifc'));
      if (ifcFiles.length === 0) {
        this.showToast(t('toast.onlyIfc'), 'warning');
        this.setStatus(t('toast.onlyIfc'));
        return;
      }
      this.fireAndForget(this.loadIfcFiles(ifcFiles), 'Load IFC files');
    });
  }

  /** Shared selection-visibility actions (rail + Properties buttons). */
  private isolateSelection(): void {
    const selectionMap = this.getValidModelIdMap(this.selectedItems);
    if (isMapEmpty(selectionMap)) {
      this.setStatus(t('status.noSelectionToIsolate'));
      return;
    }
    this.fireAndForget((async () => {
      await this.hider.isolate(cloneMap(selectionMap));
      await this.updateVisibilityCount();
      this.persistLocalState();
      this.setStatus(t('status.selectionIsolated'));
    })(), 'Isolate selection');
  }

  private hideSelection(): void {
    const selectionMap = this.getValidModelIdMap(this.selectedItems);
    if (isMapEmpty(selectionMap)) {
      this.setStatus(t('status.noSelectionToHide'));
      return;
    }
    this.fireAndForget((async () => {
      await this.hider.set(false, cloneMap(selectionMap));
      await this.updateVisibilityCount();
      this.persistLocalState();
      this.setStatus(t('status.selectionHidden'));
    })(), 'Hide selection');
  }

  private resetVisibility(): void {
    this.fireAndForget((async () => {
      await this.hider.set(true);
      this.clearFilterChecks();
      await this.updateVisibilityCount();
      this.persistLocalState();
      this.setStatus(t('status.visibilityReset'));
    })(), 'Reset visibility');
  }

  private toggleXray(): void {
    this.xrayEnabled = !this.xrayEnabled;
    this.setRailPressed(this.dom.btnTransparency, this.xrayEnabled);
    this.applyXRay();
    this.fireAndForget(this.updateFragments(true), 'Toggle x-ray');
    this.persistLocalState();
    this.setStatus(this.xrayEnabled ? t('status.xrayEnabled') : t('status.xrayDisabled'));
    this.syncMobileSheet();
  }

  private toggleEdges(): void {
    this.edgesEnabled = !this.edgesEnabled;
    this.setRailPressed(this.dom.btnWireframe, this.edgesEnabled);
    this.applyEdges();
    this.persistLocalState();
    this.setStatus(this.edgesEnabled ? t('status.edgesEnabled') : t('status.edgesDisabled'));
    this.syncMobileSheet();
  }

  private toggleIssuePinMode(): void {
    this.issuePinMode = !this.issuePinMode;
    this.setRailPressed(this.dom.btnIssuePinMode, this.issuePinMode);
    this.dom.viewerHint.hidden = !this.issuePinMode;
    this.setStatus(this.issuePinMode ? t('status.issuePinEnabled') : t('status.issuePinDisabled'));
  }

  private toggleTheme(): void {
    this.themeMode = this.themeMode === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', this.themeMode);
    setIcon(this.dom.btnThemeToggle, this.themeMode === 'dark' ? 'light_mode' : 'dark_mode');
    // F8: restore this theme's remembered background.
    this.setBackgroundColor(this.backgroundByTheme[this.themeMode], false);
    this.persistLocalState();
    this.syncMobileSheet();
  }

  /** Marks a rail/toggle button active and keeps aria-pressed in sync. */
  private setRailPressed(button: HTMLButtonElement, active: boolean): void {
    button.classList.toggle('is-active', active);
    if (button.hasAttribute('aria-pressed')) button.setAttribute('aria-pressed', String(active));
  }

  // P3 (W5.3): the section slider fires `input` per pixel of drag; recreating
  // the clip plane + fragments update each tick is expensive. Update the label
  // immediately (responsive), debounce the heavy re-section.
  private readonly debouncedSetSectionPosition = debounce((pct: number) => this.setSectionPosition(pct), 40);
  // P3: debounced per-model opacity apply for the slider drag (label stays live).
  private readonly debouncedApplyModelOpacity = debounce(
    (modelId: string, opacity: number) => this.applyModelOpacity(modelId, opacity),
    40,
  );

  private onSectionSliderInput(): void {
    const pct = Number(this.dom.sectionPos.value);
    this.dom.sectionPosLabel.textContent = t('label.percent', { value: pct });
    this.debouncedSetSectionPosition(pct);
  }

  private onFilterChipClick(event: Event): void {
    const chip = (event.target as HTMLElement).closest<HTMLButtonElement>('.filter-chip[data-filter-value]');
    if (!chip) return;
    const active = chip.classList.toggle('is-active');
    chip.setAttribute('aria-pressed', String(active));
    this.fireAndForget(this.applyFilters(), 'Apply filters');
  }

  private bindModelBrowserEvents(): void {
    this.dom.modelBrowserTree.addEventListener('click', (event) => {
      const target = event.target as HTMLElement;
      const actionButton = target.closest<HTMLButtonElement>('[data-browser-action]');
      if (!actionButton) return;
      event.preventDefault();
      event.stopPropagation();

      const modelId = actionButton.dataset.modelId;
      const action = actionButton.dataset.browserAction;
      if (!modelId || !action) return;

      if (action === 'select-model') {
        this.fireAndForget(this.selectWholeModel(modelId), 'Select model');
        return;
      }

      if (action === 'fit-model') {
        this.fitToModelById(modelId);
        return;
      }

      if (action === 'isolate-level') {
        const level = actionButton.dataset.level;
        if (!level) return;
        this.fireAndForget(this.isolateLevelForModel(modelId, level), 'Isolate level');
        return;
      }

      if (action === 'isolate-class-level') {
        const level = actionButton.dataset.level;
        const className = actionButton.dataset.class;
        if (!level || !className) return;
        this.fireAndForget(this.isolateClassForModelLevel(modelId, level, className), 'Isolate class');
        return;
      }

      if (action === 'select-item') {
        const localId = Number(actionButton.dataset.localId);
        if (!Number.isFinite(localId)) return;
        this.fireAndForget(this.selectSingleItem(modelId, localId, true), 'Select model tree item');
      }
    });
  }

  /**
   * U7: real tab pattern — role=tab/tabpanel, aria-selected, roving tabindex,
   * Left/Right/Home/End arrow keys. Opening a tab also opens the panel drawer
   * on small screens (U1).
   */
  private bindTabs(): void {
    const buttons = this.dom.tabStripButtons;
    buttons.forEach((button, i) => {
      button.addEventListener('click', () => {
        this.activateTab(button.dataset.tab ?? 'explorer');
        if (this.isSmallScreen()) this.openPanel();
      });
      button.addEventListener('keydown', (event) => {
        let next = -1;
        if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (i + 1) % buttons.length;
        else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (i - 1 + buttons.length) % buttons.length;
        else if (event.key === 'Home') next = 0;
        else if (event.key === 'End') next = buttons.length - 1;
        if (next < 0) return;
        event.preventDefault();
        const target = buttons[next];
        this.activateTab(target.dataset.tab ?? 'explorer');
        target.focus();
      });
    });
  }

  /** U1/U2/U9: panel toggle, drawer + scrim, mobile bottom nav + More sheet, splitter. */
  private bindResponsiveShell(): void {
    this.dom.btnPanelToggle.addEventListener('click', () => this.togglePanel());

    this.dom.scrim.addEventListener('click', () => {
      this.closePanel();
      this.closeSheet();
    });

    // Mobile bottom nav (U1)
    this.dom.mobileNavButtons.forEach((button) => {
      button.addEventListener('click', () => this.onMobileNav(button.dataset.mobileNav ?? ''));
    });
    this.dom.mobileFab.addEventListener('click', () => this.fitToModel());
    this.dom.btnCloseSheet.addEventListener('click', () => this.closeSheet());

    this.bindSplitter();
  }

  /** U9: pointer-event splitter with touch + keyboard (arrows resize). */
  private bindSplitter(): void {
    const splitter = this.dom.panelSplitter;
    const min = 280;
    const max = 400;
    let startX = 0;
    let startW = 0;
    const currentW = (): number =>
      parseInt(getComputedStyle(document.documentElement).getPropertyValue('--panel-w'), 10) || 320;
    const setW = (w: number): void => {
      const width = Math.min(max, Math.max(min, w));
      document.documentElement.style.setProperty('--panel-w', `${width}px`);
      splitter.setAttribute('aria-valuenow', String(width));
    };
    splitter.addEventListener('pointerdown', (event) => {
      splitter.setPointerCapture(event.pointerId);
      splitter.classList.add('dragging');
      startX = event.clientX;
      startW = currentW();
      event.preventDefault();
    });
    splitter.addEventListener('pointermove', (event) => {
      if (!splitter.hasPointerCapture(event.pointerId)) return;
      setW(startW + (startX - event.clientX));
    });
    const end = (event: PointerEvent): void => {
      if (splitter.hasPointerCapture(event.pointerId)) splitter.releasePointerCapture(event.pointerId);
      splitter.classList.remove('dragging');
      this.persistLocalState();
    };
    splitter.addEventListener('pointerup', end);
    splitter.addEventListener('pointercancel', end);
    splitter.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowLeft') setW(currentW() + 16);
      else if (event.key === 'ArrowRight') setW(currentW() - 16);
      else return;
      event.preventDefault();
      this.persistLocalState();
    });
  }

  private isSmallScreen(): boolean {
    return window.matchMedia('(max-width: 1023px)').matches;
  }

  private openPanel(): void {
    this.dom.root.classList.remove('panel-collapsed');
    this.dom.root.classList.add('panel-open');
    this.closeSheet();
    this.updatePanelToggleUi();
  }

  private closePanel(): void {
    this.dom.root.classList.remove('panel-open');
    this.updateScrim();
    this.updatePanelToggleUi();
  }

  private togglePanel(): void {
    if (this.isSmallScreen()) {
      if (this.dom.root.classList.contains('panel-open')) this.closePanel();
      else this.openPanel();
    } else {
      this.dom.root.classList.toggle('panel-collapsed');
      this.updatePanelToggleUi();
      const renderer = this.world?.renderer;
      if (renderer) requestAnimationFrame(() => renderer.resize());
    }
    this.updateScrim();
  }

  private updatePanelToggleUi(): void {
    const collapsed = this.dom.root.classList.contains('panel-collapsed');
    const openMobile = this.dom.root.classList.contains('panel-open');
    const shown = this.isSmallScreen() ? openMobile : !collapsed;
    setIcon(this.dom.btnPanelToggle, shown ? 'right_panel_close' : 'right_panel_open');
    this.dom.btnPanelToggle.setAttribute('aria-expanded', String(shown));
  }

  private updateScrim(): void {
    const open = this.dom.root.classList.contains('panel-open') || this.dom.root.classList.contains('sheet-open');
    this.dom.scrim.hidden = !open;
  }

  private onMobileNav(nav: string): void {
    switch (nav) {
      case 'tree':
        this.activateTab('explorer');
        this.openPanel();
        break;
      case 'orbit':
        this.applyNavigationMode('Orbit');
        this.setHomeView();
        this.closePanel();
        this.closeSheet();
        break;
      case 'section':
        if (this.dom.btnSectionZ.classList.contains('is-active')) this.clearSections();
        else this.toggleSectionPlane(this.dom.btnSectionZ, new THREE.Vector3(0, -1, 0));
        this.closePanel();
        this.closeSheet();
        break;
      case 'measure':
        this.setMeasureMode(this.measureMode === 'length' ? 'none' : 'length');
        this.closePanel();
        this.closeSheet();
        break;
      case 'more':
        this.openSheet();
        break;
    }
    this.updateMobileNavActive(nav);
  }

  private updateMobileNavActive(active: string): void {
    this.dom.mobileNavButtons.forEach((button) => {
      button.classList.toggle('is-active', button.dataset.mobileNav === active);
    });
  }

  /** U2: view settings (theme/style/toggles/background) reachable on phones. */
  private openSheet(): void {
    this.buildMobileSheet();
    this.dom.root.classList.add('sheet-open');
    this.closePanel();
    this.updateScrim();
  }

  private closeSheet(): void {
    this.dom.root.classList.remove('sheet-open');
    this.updateScrim();
  }

  private buildMobileSheet(): void {
    buildMobileSheetView(
      this.dom.sheetBody,
      {
        xrayEnabled: this.xrayEnabled,
        edgesEnabled: this.edgesEnabled,
        gridVisible: this.gridVisible,
        themeMode: this.themeMode,
        visualStyle: this.visualStyle,
      },
      {
        toggleXray: () => this.toggleXray(),
        toggleEdges: () => this.toggleEdges(),
        toggleGrid: () => { this.setGridVisible(!this.gridVisible, true); this.persistLocalState(); this.syncMobileSheet(); },
        toggleTheme: () => this.toggleTheme(),
        setVisualStyle: (value: string) =>
          this.fireAndForget(this.setVisualStyle(this.parseVisualStyle(value), true, true, true), 'Set visual style'),
      },
      {
        xray: t('mobileSheet.xray'),
        edges: t('mobileSheet.edges'),
        grid: t('mobileSheet.grid'),
        lightTheme: t('mobileSheet.lightTheme'),
        style: t('mobileSheet.style'),
        styleOption: (value) => this.getVisualStyleLabel(this.parseVisualStyle(value)),
      },
    );
  }

  /** Keeps the More sheet's toggle states current when opened. */
  private syncMobileSheet(): void {
    if (this.dom.root.classList.contains('sheet-open')) this.buildMobileSheet();
  }

  private bindFederationTreeEvents(): void {
    this.dom.federationTree.addEventListener('click', (event) => {
      const target = event.target as HTMLElement;
      const actionButton = target.closest<HTMLButtonElement>('[data-model-action]');
      if (actionButton) {
        const modelId = actionButton.dataset.modelId;
        const action = actionButton.dataset.modelAction;
        if (!modelId || !action) return;
        if (action === 'fit') {
          this.fitToModelById(modelId);
          return;
        }
        if (action === 'reset') {
          this.resetModelOffsets(modelId);
          return;
        }
        if (action === 'select-model') {
          this.fireAndForget(this.selectWholeModel(modelId), 'Select model');
          return;
        }
        if (action === 'toggle-visibility') {
          this.toggleModelVisibility(modelId);
          return;
        }
        if (action === 'toggle-gizmo') {
          this.toggleModelGizmo(modelId);
          return;
        }
        if (action === 'unload') {
          this.fireAndForget((async () => {
            const record = this.federatedModels.get(modelId);
            if (!record) return;
            const confirmed = await this.confirm(
              t('confirm.unloadModel', { name: record.fileName }),
              t('confirm.unload'),
              t('confirm.cancel'),
            );
            if (!confirmed) return;
            await this.unloadModel(modelId);
          })(), 'Unload model');
          return;
        }
      }

      const levelButton = target.closest<HTMLButtonElement>('[data-model-id][data-level]');
      if (levelButton) {
        const modelId = levelButton.dataset.modelId;
        const level = levelButton.dataset.level;
        if (!modelId || !level) return;
        this.fireAndForget(this.isolateLevelForModel(modelId, level), 'Isolate level');
      }
    });

    this.dom.federationTree.addEventListener('input', (event) => {
      const target = event.target as HTMLElement;
      const opacityInput = target.closest<HTMLInputElement>('input[data-model-id][data-model-opacity]');
      if (!opacityInput) return;
      const modelId = opacityInput.dataset.modelId;
      if (!modelId) return;
      const opacity = Number(opacityInput.value) / 100;
      // P3: the label updates instantly for feedback; the heavy opacity apply
      // (fragments update + persist) is debounced so a drag doesn't thrash it.
      const card = opacityInput.closest<HTMLElement>('.federated-opacity');
      const valueLabel = card?.querySelector<HTMLElement>('[data-opacity-value]');
      if (valueLabel) valueLabel.textContent = t('label.percent', { value: Math.round(this.clamp(opacity, 0, 1) * 100) });
      this.debouncedApplyModelOpacity(modelId, opacity);
    });

    this.dom.federationTree.addEventListener('change', (event) => {
      const target = event.target as HTMLElement;
      const opacityInput = target.closest<HTMLInputElement>('input[data-model-id][data-model-opacity]');
      if (opacityInput) {
        const modelId = opacityInput.dataset.modelId;
        if (!modelId) return;
        const opacity = Number(opacityInput.value) / 100;
        this.applyModelOpacity(modelId, opacity);
        this.renderFederatedTree();
        return;
      }
      const input = target.closest<HTMLInputElement>('input[data-model-id][data-transform]');
      if (!input) return;
      this.applyTransformInput(input);
    });
  }

  /** A11/U6: delegated viewpoint-row actions (bound once; survives re-renders). */
  private bindViewpointListEvents(): void {
    const rowId = (event: Event): string | null => {
      const row = (event.target as HTMLElement).closest<HTMLElement>('[data-viewpoint-id]');
      return row?.dataset.viewpointId ?? null;
    };
    this.dom.viewpointList.addEventListener('click', (event) => {
      const id = rowId(event);
      if (id === null) return;
      this.selectedViewpointId = id;
      const action = (event.target as HTMLElement).closest<HTMLElement>('[data-viewpoint-action]')?.dataset.viewpointAction;
      if (action === 'apply') {
        this.fireAndForget(this.applySelectedViewpoint(), 'Apply viewpoint');
      } else if (action === 'delete') {
        this.fireAndForget(this.deleteSelectedViewpoint(), 'Delete viewpoint');
      } else {
        this.updateViewpointList();
      }
    });
    this.dom.viewpointList.addEventListener('dblclick', (event) => {
      const id = rowId(event);
      if (id === null) return;
      this.selectedViewpointId = id;
      this.fireAndForget(this.applySelectedViewpoint(), 'Apply viewpoint');
    });
  }

  /** A11/U6: delegated issue-row selection + delete (bound once). */
  private bindIssueListEvents(): void {
    const rowId = (event: Event): string | null => {
      const row = (event.target as HTMLElement).closest<HTMLElement>('[data-issue-id]');
      return row?.dataset.issueId ?? null;
    };
    this.dom.issuesList.addEventListener('click', (event) => {
      const id = rowId(event);
      if (id === null) return;
      const action = (event.target as HTMLElement).closest<HTMLElement>('[data-issue-action]')?.dataset.issueAction;
      if (action === 'delete') {
        this.activeIssueId = id;
        this.fireAndForget(this.deleteSelectedIssue(), 'Delete issue');
        return;
      }
      this.selectIssue(id, false);
    });
    this.dom.issuesList.addEventListener('dblclick', (event) => {
      const id = rowId(event);
      if (id !== null) this.selectIssue(id, true);
    });
  }

  private async initEngine(): Promise<void> {
    // Engine construction lives in core/viewer-core.ts (shared with the /embed
    // entry). W5.1 (P1): the module is DYNAMICALLY imported here so three +
    // @thatopen + web-ifc (~1MB gzip) form async chunks the initial shell does
    // not download. The `this`-coupled callbacks (model registration,
    // camera->fragments update, gizmo->panel re-render) are wired here against
    // the returned handles so behaviour matches the pre-split inline setup.
    this.engineModule = await import('./core/viewer-core');
    const engineOptions: BootstrapEngineOptions = {
      container: this.dom.viewerContainer,
      backgroundColor: this.backgroundColor,
      gridVisible: this.gridVisible,
      getBackgroundColor: () => this.backgroundColor,
      getPostproductionRenderer: () => this.getPostproductionRenderer(),
    };
    const engine: EngineHandles = await this.engineModule.bootstrapEngine(engineOptions);
    this.engineHandles = engine;
    this.components = engine.components;
    this.world = engine.world;
    this.fragments = engine.fragments;
    this.clipper = engine.clipper;
    this.hider = engine.hider;
    this.raycaster = engine.raycaster;
    this.lengthMeasurement = engine.lengthMeasurement;
    this.areaMeasurement = engine.areaMeasurement;
    this.markerManager = engine.markerManager;
    this.transformControls = engine.transformControls;
    this.transformControlsHelper = engine.transformControlsHelper;
    this.gridHelper = engine.gridHelper;

    // Expose test/debug handles only in builds made with the explicit VITE_E2E
    // define (vite.e2e.config.ts) so e2e can exercise the real production
    // artifact (AUDIT T4). Plain dev/prod builds ship without these hooks.
    // T6 (W2.5): expose ONLY the frozen, explicit test contract — not the whole
    // instance or its private fields — and only in VITE_E2E builds.
    if (import.meta.env.VITE_E2E === 'true') {
      window.__viewerTestApi = this.buildTestApi();
    }

    // P6 (W5.3): on-demand rendering is enabled LATER (end of init) so the
    // engine renders a few AUTO frames first — the PostproductionRenderer lazily
    // builds its composer/basePass on those frames, and its `enabled` setter
    // reads the throwing `basePass` getter, so configuring postproduction before
    // the first render would throw ("Base pass not initialized"). Until then
    // requestRender() stays a no-op (AUTO renders every frame anyway).
    this.world.camera.controls.addEventListener('update', () => {
      this.fireAndForget(this.fragments.core.update(), 'Camera update');
      this.requestRender();
    });

    this.fragments.core.onModelLoaded.add((model) => {
      const modelId = String(model?.modelId ?? '');
      if (!modelId) return;
      // Streamed LOD/mesh updates for this model → paint the new geometry.
      model.onViewUpdated?.add?.(() => this.requestRender());
      this.fireAndForget(this.registerModel(modelId, model), 'Register model');
    });

    this.transformControls.addEventListener('dragging-changed', (event) => {
      const dragging = Boolean((event as { value?: unknown }).value);
      this.gizmoDragging = dragging;
      this.world.camera.controls.enabled = !dragging;
    });
    this.transformControls.addEventListener('objectChange', () => {
      if (!this.activeGizmoModelId) return;
      const model = this.federatedModels.get(this.activeGizmoModelId);
      if (!model) return;
      this.updateModelOffsetsFromObject(model);
      if (this.edgesEnabled) this.applyEdges();
      this.fireAndForget(this.updateFragments(true), 'Gizmo update');
    });
    this.transformControls.addEventListener('mouseUp', () => {
      if (!this.activeGizmoModelId) return;
      const model = this.federatedModels.get(this.activeGizmoModelId);
      if (!model) return;
      this.updateModelOffsetsFromObject(model);
      this.renderModelBrowser();
      this.renderFederatedTree();
      this.persistLocalState();
      this.setStatus(t('status.gizmoUpdated', { name: model.fileName }));
    });

    // R1 (W5-fixups): dragging a section-plane arrow disables the camera controls
    // (no 'update' pump), so re-clipped geometry would freeze mid-drag. Run a
    // per-frame rAF render pump for the duration of the drag. onBeforeDrag /
    // onAfterDrag are the component-level hooks (they wrap each plane's
    // onDraggingStarted/onDraggingEnded in this pinned @thatopen version).
    this.clipper.onBeforeDrag.add(() => this.startSectionDragPump());
    this.clipper.onAfterDrag.add(() => this.stopSectionDragPump());

    // R2 (W5-fixups): the live measurement rubber-band mutates on pointer move
    // between clicks, but MANUAL mode only repaints on requestRender. Subscribe
    // the measurements' pointer events so the in-progress preview repaints.
    for (const measurement of [this.lengthMeasurement, this.areaMeasurement]) {
      measurement.onPointerMove.add(() => this.requestRender());
      measurement.onPointerStop.add(() => this.requestRender());
    }

    await this.updateVisibilityCount();
  }

  /**
   * R1 (W5-fixups): starts a bounded per-frame render pump for a section-plane
   * gizmo drag (camera controls are disabled during the drag, so the camera
   * pump doesn't fire). Cancelled by stopSectionDragPump on drag end / destroy.
   */
  private startSectionDragPump(): void {
    if (this.sectionDragRaf) return;
    const pump = (): void => {
      this.requestRender();
      this.sectionDragRaf = requestAnimationFrame(pump);
    };
    this.sectionDragRaf = requestAnimationFrame(pump);
  }

  private stopSectionDragPump(): void {
    if (!this.sectionDragRaf) return;
    cancelAnimationFrame(this.sectionDragRaf);
    this.sectionDragRaf = 0;
    // Paint one final frame at the drag's resting position.
    this.requestRender();
  }

  /**
   * Single registration path for loaded models (A6/A10): dedupes the
   * onModelLoaded event and the awaited load path via one promise per model
   * id, and diverts late arrivals of timed-out loads to disposal.
   */
  private registerModel(modelId: string, model: FragmentsModelLike): Promise<void> {
    if (this.modelRegistry.isStale(modelId)) {
      return this.disposeStaleModel(modelId);
    }
    let registration = this.modelRegistrations.get(modelId);
    if (!registration) {
      registration = this.onModelAdded(modelId, model).catch((error: unknown) => {
        // Allow a retry after a failed registration.
        this.modelRegistrations.delete(modelId);
        throw error;
      });
      this.modelRegistrations.set(modelId, registration);
    }
    return registration;
  }

  /** Disposes a model whose load timed out but completed later (A10). */
  private async disposeStaleModel(modelId: string): Promise<void> {
    if (!this.modelRegistry.consumeStale(modelId)) return;
    if (this.federatedModels.has(modelId)) {
      // The registration raced ahead of the timeout — run the full unload.
      await this.unloadModel(modelId);
      return;
    }
    this.modelRegistrations.delete(modelId);
    if (this.fragments?.list?.has(modelId)) {
      await this.fragments.core.disposeModel(modelId);
    }
    console.debug(`Disposed late-arriving model after load timeout: ${modelId}`);
  }

  /**
   * Per-model unload (F6): frees the engine-side fragments (worker memory,
   * meshes, materials) and clears every piece of viewer state tied to the id.
   */
  private async unloadModel(modelId: string): Promise<void> {
    const record = this.federatedModels.get(modelId);
    if (!record) return;

    if (this.activeGizmoModelId === modelId) this.detachModelGizmo();
    this.federatedModels.delete(modelId);
    this.modelIndices.delete(modelId);
    this.modelUnits.delete(modelId);
    this.modelRegistrations.delete(modelId);
    this.appliedModelOpacity.delete(modelId);
    delete this.selectedItems[modelId];
    this.modelObjects = this.modelObjects.filter((object) => object !== record.object);

    if (this.fragments?.list?.has(modelId)) {
      await this.fragments.core.disposeModel(modelId);
    }

    // Pins referencing the unloaded model are hidden, not deleted (F9).
    this.refreshIssueMarkers();
    this.applyXRay();
    this.applyEdges();
    await this.refreshSelectionVisuals();
    this.renderModelBrowser();
    this.renderFederatedTree();
    this.renderClassFilters();
    this.renderLevelFilters();
    this.updateElementCounter();
    await this.updateVisibilityCount();
    await this.updateFragments(true);
    this.dom.emptyState.hidden = this.federatedModels.size > 0;
    this.setStatus(t('status.modelUnloaded', { name: record.fileName }));
  }

  private async onModelAdded(modelId: string, model: FragmentsModelLike): Promise<void> {
    if (!modelId || this.federatedModels.has(modelId)) return;

    model.useCamera(this.world.camera.three);
    if (typeof model?.graphicsQuality === 'number') model.graphicsQuality = 1;
    this.world.scene.three.add(model.object);

    const modelObject = model.object;
    if (!this.modelObjects.includes(modelObject)) this.modelObjects.push(modelObject);

    const ids = await model.getItemsIdsWithGeometry();
    this.dom.emptyState.hidden = true;

    const meta = this.modelRegistry.completeLoad(modelId);
    const fileName = meta?.fileName || modelId;
    const modelRecord: FederatedModelRecord = {
      modelId,
      fileName,
      sizeBytes: meta?.sizeBytes ?? 0,
      elementCount: ids.length,
      visible: true,
      opacity: 1,
      object: modelObject,
      basePosition: {
        x: modelObject.position.x,
        y: modelObject.position.y,
        z: modelObject.position.z,
      },
      baseRotation: {
        x: modelObject.rotation.x,
        y: modelObject.rotation.y,
        z: modelObject.rotation.z,
      },
      offsetPosition: { x: 0, y: 0, z: 0 },
      offsetRotation: { x: 0, y: 0, z: 0 },
    };
    this.federatedModels.set(modelId, modelRecord);
    this.updateElementCounter();
    this.renderModelBrowser();
    this.renderFederatedTree();

    // C8 (W5.2): cache the converted `.frag` bytes in IndexedDB so the model can
    // be restored next session without re-conversion. During a restore the bytes
    // came FROM the cache, so reuse the known key and skip the re-write. Run it
    // in the background — it must never delay model registration / the load flow.
    this.fireAndForget(this.cacheFragForModel(modelRecord, model), 'Cache fragments');

    await this.indexModel(modelId, model);
    // F5: resolve the model's display units from its own unit entities;
    // failures fall back to the metric defaults.
    try {
      this.modelUnits.set(modelId, resolveModelUnits(await this.fetchUnitRows(model)));
    } catch (error) {
      console.debug(`Unit resolution failed for ${modelId}; using defaults`, error);
      this.modelUnits.set(modelId, DEFAULT_MODEL_UNITS);
    }
    this.renderModelBrowser();
    this.renderFederatedTree();

    await this.setVisualStyle(this.visualStyle, false, false);

    if (!this.suppressAutoFit) this.fitToModel();
    await this.updateVisibilityCount();

    this.renderClassFilters();
    this.renderLevelFilters();

    this.refreshIssueMarkers();
    // C8 (W5.2): persist the session now that the model is registered with its
    // fragKey — so a plain load (no further edits) is restorable next session /
    // offline. Skipped during a restore (restoringSession guards persist).
    this.persistLocalState();
    this.setStatus(t('status.modelLoaded', { name: fileName }));
  }

  private async indexModel(modelId: string, model: FragmentsModelLike): Promise<void> {
    this.modelIndices.set(modelId, await buildModelIndex(modelId, model));
  }

  /**
   * C8 (W5.2): caches the model's converted `.frag` bytes in IndexedDB and stamps
   * the record's fragKey so the session serializer can reference it. During a
   * restore the key is already known (bytes came from the cache) — reuse it and
   * skip the write. Failures (private mode / quota / no IDB) degrade gracefully:
   * the model still works this session, it just won't be restorable.
   */
  private async cacheFragForModel(record: FederatedModelRecord, model: FragmentsModelLike): Promise<void> {
    const restoreKey = this.restoreFragKeys.get(record.modelId);
    if (restoreKey) {
      record.fragKey = restoreKey;
      this.restoreFragKeys.delete(record.modelId);
      return;
    }
    try {
      const buffer = await model.getBuffer(false);
      const bytes = new Uint8Array(buffer);
      const key = hashFragBytes(bytes);
      record.fragKey = key;
      await this.fragCache.put({
        key,
        fileName: record.fileName,
        sizeBytes: bytes.byteLength,
        storedAt: Date.now(),
        bytes,
      });
    } catch (error) {
      console.debug(`Frag cache write failed for ${record.modelId}; model not restorable`, error);
    }
  }

  private renderClassFilters(): void {
    const classes = new Set<string>();
    for (const index of this.modelIndices.values()) {
      for (const className of index.classes.keys()) classes.add(className);
    }
    const sorted = [...classes].sort((a, b) => a.localeCompare(b));
    // A1: escaped by the pure markup builder.
    this.dom.classFilterList.innerHTML = filterChipMarkup(sorted, 'class', t('empty.noClasses'));
  }

  private renderLevelFilters(): void {
    const levels = new Set<string>();
    for (const index of this.modelIndices.values()) {
      for (const levelName of index.levels.keys()) levels.add(levelName);
    }
    const sorted = [...levels].sort((a, b) => a.localeCompare(b));
    // A1: escaped by the pure markup builder.
    this.dom.levelFilterList.innerHTML = filterChipMarkup(sorted, 'level', t('empty.noLevels'));
  }

  private clearFilterChecks(): void {
    for (const list of [this.dom.classFilterList, this.dom.levelFilterList]) {
      list.querySelectorAll<HTMLButtonElement>('.filter-chip.is-active').forEach((chip) => {
        chip.classList.remove('is-active');
        chip.setAttribute('aria-pressed', 'false');
      });
    }
  }

  private updateElementCounter(): void {
    let total = 0;
    for (const model of this.federatedModels.values()) total += model.elementCount;
    this.dom.elementCount.textContent = t('label.elements', { count: total });
    this.updateTopbarModel();
  }

  /** Top-bar model name + per-model tag chips (design: "Name  [ARC] [STR]"). */
  private updateTopbarModel(): void {
    const models = [...this.federatedModels.values()];
    this.dom.topbarModel.replaceChildren();
    this.dom.navPill.hidden = models.length === 0;
    this.dom.emptyState.hidden = models.length > 0;
    if (models.length === 0) {
      const empty = document.createElement('span');
      empty.className = 'topbar-model-empty';
      empty.textContent = t('empty.noModelLoaded');
      this.dom.topbarModel.append(empty);
      return;
    }
    const name = document.createElement('span');
    name.className = 'topbar-model-name';
    name.textContent = models.length === 1 ? models[0].fileName : t('label.models', { count: models.length });
    this.dom.topbarModel.append(name);
    for (const model of models.slice(0, 4)) {
      const chip = document.createElement('span');
      chip.className = 'model-chip';
      // Short tag: uppercased base name without extension, capped.
      const base = model.fileName.replace(/\.ifc$/i, '');
      chip.textContent = base.length > 6 ? base.slice(0, 6).toUpperCase() : base.toUpperCase();
      chip.title = model.fileName;
      this.dom.topbarModel.append(chip);
    }
  }

  /** Builds the already-translated label bundle for the model-browser tree. */
  private browserLabels(): BrowserLabels {
    return {
      hidden: t('tree.hidden'),
      building: t('tree.building'),
      noElements: t('tree.noElements'),
      noClasses: t('tree.noClasses'),
      default: t('tree.default'),
      levels: t('tree.levels'),
      spatialStructure: t('tree.spatialStructure'),
      noLevelsDetected: t('tree.noLevelsDetected'),
      noSpatialData: t('tree.noSpatialData'),
      select: t('tree.select'),
      isolate: t('tree.isolate'),
      isolateLevel: t('tree.isolateLevel'),
      fitCamera: t('fed.fitCamera'),
      selectFullModel: t('fed.selectFullModel'),
      elementFallback: (id) => t('label.elementFallback', { id }),
      moreNodes: (count) => t('tree.moreNodes', { count }),
      moreElements: (count) => t('tree.moreElements', { count }),
      moreLevels: (count) => t('tree.moreLevels', { count }),
      levelsShort: (count) => t('tree.levelsShort', { count }),
    };
  }

  private renderModelBrowser(): void {
    const markup = buildModelBrowserMarkup(this.federatedModels, this.modelIndices, this.browserLabels());
    if (markup === null) {
      this.dom.modelBrowserTree.innerHTML = `<div class="tree-item">${escapeHtml(t('empty.noModelsYet'))}</div>`;
      return;
    }
    this.renderPreservingDetails(this.dom.modelBrowserTree, markup);
  }

  /**
   * A11: replaces a container's innerHTML while preserving the open/closed
   * state of any `<details data-node-key>` the user toggled. Without this a
   * full re-render (e.g. after a model loads) would collapse every expanded
   * tree node back to its markup default. Nodes present after the render but
   * absent from the snapshot keep their markup default (fresh nodes).
   */
  private renderPreservingDetails(container: HTMLElement, markup: string): void {
    const openState = new Map<string, boolean>();
    container.querySelectorAll<HTMLDetailsElement>('details[data-node-key]').forEach((node) => {
      const key = node.dataset.nodeKey;
      if (key) openState.set(key, node.open);
    });

    container.innerHTML = markup;

    container.querySelectorAll<HTMLDetailsElement>('details[data-node-key]').forEach((node) => {
      const key = node.dataset.nodeKey;
      if (key && openState.has(key)) node.open = openState.get(key)!;
    });
  }

  private renderFederatedTree(): void {
    const markup = buildFederationTreeMarkup(this.federatedModels, this.modelIndices, this.activeGizmoModelId, {
      show: t('fed.show'),
      hide: t('fed.hide'),
      noStoreys: t('fed.noStoreys'),
      opacity: t('fed.opacity'),
      offsetXyz: t('fed.offsetXyz'),
      rotationXyz: t('fed.rotationXyz'),
      select: t('fed.select'),
      gizmo: t('fed.gizmo'),
      fit: t('fed.fit'),
      reset: t('fed.reset'),
      unload: t('fed.unload'),
      selectFullModel: t('fed.selectFullModel'),
      fitCamera: t('fed.fitCamera'),
      unloadTitle: t('fed.unloadTitle'),
      elements: (count) => t('label.elements', { count }),
      isolateLevel: (level) => t('fed.isolateLevel', { level }),
      levelsCount: (count) => t('fed.levelsCount', { count }),
    });
    if (markup === null) {
      this.dom.federationTree.innerHTML = `<div class="tree-item">${escapeHtml(t('empty.noModelsYet'))}</div>`;
      return;
    }
    this.dom.federationTree.innerHTML = markup;
  }

  private normalizeHexColor(value: string | null | undefined, fallback = DEFAULT_BACKGROUND_COLOR): string {
    if (!value) return fallback;
    const normalized = value.trim().toLowerCase();
    if (/^#[0-9a-f]{6}$/.test(normalized)) return normalized;
    if (/^#[0-9a-f]{3}$/.test(normalized)) {
      return `#${normalized[1]}${normalized[1]}${normalized[2]}${normalized[2]}${normalized[3]}${normalized[3]}`;
    }
    return fallback;
  }

  private setGridVisible(visible: boolean, updateStatus: boolean): void {
    this.gridVisible = visible;
    if (this.gridHelper) this.gridHelper.visible = visible;
    this.setRailPressed(this.dom.btnToggleGrid, visible);
    this.requestRender();
    if (updateStatus) this.setStatus(visible ? t('status.gridEnabled') : t('status.gridHidden'));
  }

  private setBackgroundColor(color: string, updateStatus: boolean): void {
    const normalized = this.normalizeHexColor(color);
    this.backgroundColor = normalized;
    this.backgroundByTheme[this.themeMode] = normalized;
    const threeColor = new THREE.Color(normalized);
    if (this.world?.scene?.three) this.world.scene.three.background = threeColor;
    // Sync the renderer clear color so PEN style (which bypasses the scene color pass)
    // uses the user-selected background
    const renderer = this.world?.renderer?.three;
    if (renderer) {
      renderer.setClearColor(threeColor, 1);
    }
    // Also sync the postproduction basePass clear color (defensive: the getter
    // throws until the composer's first render — P6 on-demand, see safeBasePass).
    const postRenderer = this.getPostproductionRenderer();
    const post = postRenderer?.postproduction;
    const basePass = post ? this.safeBasePass(post) : null;
    if (basePass) {
      basePass.clearColor = threeColor;
      basePass.clearAlpha = 1;
    }
    this.requestRender();
    if (updateStatus) this.setStatus(t('status.backgroundSet', { color: normalized }));
  }

  private syncVisualSettingsUi(): void {
    this.setRailPressed(this.dom.btnToggleGrid, this.gridVisible);
    this.setRailPressed(this.dom.btnTransparency, this.xrayEnabled);
    this.setRailPressed(this.dom.btnWireframe, this.edgesEnabled);
    setIcon(this.dom.btnThemeToggle, this.themeMode === 'dark' ? 'light_mode' : 'dark_mode');
    this.setActiveView(this.activeView);
  }

  private parseVisualStyle(value: string | null | undefined): VisualStyle {
    switch ((value || '').trim()) {
      case 'basic':
      case 'pen':
      case 'color-pen':
      case 'color-shadows':
      case 'color-pen-shadows':
        return value as VisualStyle;
      default:
        return 'color-pen-shadows';
    }
  }

  private getVisualStyleLabel(style: VisualStyle): string {
    switch (style) {
      case 'basic':
        return t('style.basic');
      case 'pen':
        return t('style.pen');
      case 'color-pen':
        return t('style.colorPen');
      case 'color-shadows':
        return t('style.colorShadows');
      case 'color-pen-shadows':
        return t('style.colorPenShadows');
      default:
        return t('style.colorPenShadows');
    }
  }

  /** Localized display for a stored issue priority enum (value stays English). */
  private localizeIssuePriority(priority: string): string {
    const keys: Record<string, Parameters<typeof t>[0]> = {
      Critical: 'shell.prioCritical',
      High: 'shell.prioHigh',
      Medium: 'shell.prioMedium',
      Low: 'shell.prioLow',
    };
    return keys[priority] ? t(keys[priority]) : priority;
  }

  /** Localized display for a stored issue status enum (value stays English). */
  private localizeIssueStatus(status: string): string {
    const keys: Record<string, Parameters<typeof t>[0]> = {
      Open: 'shell.statusOpen',
      'In Progress': 'shell.statusInProgress',
      Resolved: 'shell.statusResolved',
      Closed: 'shell.statusClosed',
    };
    return keys[status] ? t(keys[status]) : status;
  }

  private getPostproductionRenderer(): OBCF.PostproductionRenderer | null {
    const renderer = this.world?.renderer as unknown;
    if (!renderer || typeof renderer !== 'object') return null;
    if (!('postproduction' in renderer)) return null;
    return renderer as OBCF.PostproductionRenderer;
  }

  private configurePostproduction(style: VisualStyle): void {
    const postRenderer = this.getPostproductionRenderer();
    const post = postRenderer?.postproduction;
    if (!post) return;

    // Reset all effects
    post.enabled = true;
    post.outlinesEnabled = false;
    post.glossEnabled = false;
    post.excludedObjectsEnabled = false;
    post.smaaEnabled = true;

    // Ensure the base renderer clears between frames (prevents ghost lines in Pen mode)
    const baseRenderer = postRenderer?.three;
    if (baseRenderer) {
      baseRenderer.autoClear = true;
    }

    // Sync the postproduction basePass clear color with the user's background.
    // NOTE: `post.basePass` is a getter that THROWS ("Base pass not initialized")
    // until the composer has run once. Under P6 on-demand (MANUAL) rendering the
    // first render may not have happened yet at boot, so read it defensively.
    const basePass = this.safeBasePass(post);
    if (basePass) {
      basePass.clearColor = new THREE.Color(this.backgroundColor);
      basePass.clearAlpha = 1;
      basePass.clearDepth = true;
    }

    // Map directly to ThatOpen PostproductionAspect enum. The enum is a runtime
    // value from @thatopen, so the mapping lives in the dynamically-imported
    // engine module (W5.1) to keep @thatopen out of the initial shell.
    this.engineModule.applyPostproductionStyle(post, style);
  }

  /**
   * Safely reads the postproduction basePass. The library getter throws until
   * the composer is initialized (its first render) — under on-demand rendering
   * (P6) that may not have happened yet, so we swallow the throw and return null.
   */
  private safeBasePass(post: { basePass?: unknown }): {
    clearColor: THREE.Color;
    clearAlpha: number;
    clearDepth: boolean;
  } | null {
    try {
      return (post.basePass as {
        clearColor: THREE.Color;
        clearAlpha: number;
        clearDepth: boolean;
      }) ?? null;
    } catch {
      return null;
    }
  }

  private async setVisualStyle(style: VisualStyle, updateStatus: boolean, persist: boolean, resetToggles = false): Promise<void> {
    const resolvedStyle = this.parseVisualStyle(style);
    this.visualStyle = resolvedStyle;

    // A16: resetModelColors()/restoreOriginalLighting() were permanent no-ops —
    // their guard flags could never be set true after the W0.4 deletion of
    // applyHiddenLineColors()/applyConsistentLighting(). Removed with the flags.
    this.configurePostproduction(resolvedStyle);

    // F3: X-ray/edges are reset only on an explicit user style change —
    // model loads and state restores re-apply the current toggles instead
    // of wiping them.
    if (resetToggles) {
      this.xrayEnabled = false;
      this.edgesEnabled = false;
      this.setRailPressed(this.dom.btnTransparency, false);
      this.setRailPressed(this.dom.btnWireframe, false);
    }
    this.applyXRay();
    this.applyEdges();

    if (this.fragments.list.size > 0 || this.federatedModels.size > 0) {
      await this.updateFragments(true);
    }
    // Postprocessing style changed even when no model is loaded — repaint.
    this.requestRender();

    if (persist) this.persistLocalState();
    if (updateStatus) this.setStatus(t('status.visualStyle', { style: this.getVisualStyleLabel(resolvedStyle) }));
  }

  private clamp(value: number, min: number, max: number): number {
    return Math.min(max, Math.max(min, value));
  }

  // The model id passed to ifcLoader.load IS the fragments.list key and the
  // model.modelId (A6) — lookups are direct, no alias resolution needed.
  private getFragmentsModel(modelId: string): FragmentsModelLike | null {
    return (this.fragments?.list?.get(modelId) as FragmentsModelLike | undefined) ?? null;
  }

  private fireAndForget(task: Promise<unknown>, context: string): void {
    task.catch((error) => this.handleAsyncError(context, error));
  }

  /**
   * P6 (W5.3): updates the fragments engine AND re-arms the on-demand renderer,
   * so every visual mutation that touches geometry paints a frame. requestRender
   * is fired both now (the change is visible immediately) and again when the
   * async update settles (streamed meshes). Replaces bare fragments.core.update
   * calls so no visual change is left unpainted under MANUAL-mode rendering.
   */
  private updateFragments(force = false): Promise<void> {
    this.requestRender();
    return this.fragments.core.update(force).then(() => this.requestRender());
  }

  private handleAsyncError(context: string, error: unknown): void {
    const message = serializeError(error);
    if (isModelNotFoundError(error)) {
      // A9: benign engine race (operation vs model unload) — self-heals by
      // pruning the selection; logged for observability, not surfaced.
      console.debug(`Suppressed model-not-found race during "${context}":`, error);
      this.pruneSelectedItems();
      if (context !== 'Camera update') {
        this.setStatus(t('status.modelSyncUpdated'));
      }
      return;
    }
    // U4: every unexpected async failure surfaces as an error toast, not just
    // 11px status text (which is hidden on phones).
    this.showToast(t('status.contextFailed', { context, message }), 'error');
    this.setStatus(t('status.contextFailed', { context, message }));
    console.error(error);
  }

  private isLoadedModelId(modelId: string): boolean {
    return Boolean(modelId && this.fragments?.list?.has(modelId));
  }

  private getValidModelIdMap(input: OBC.ModelIdMap): OBC.ModelIdMap {
    const valid: OBC.ModelIdMap = {};
    for (const [modelId, ids] of Object.entries(input)) {
      if (ids.size === 0) continue;
      if (!this.isLoadedModelId(modelId)) continue;
      if (!valid[modelId]) valid[modelId] = new Set<number>();
      for (const localId of ids) valid[modelId].add(localId);
    }
    return valid;
  }

  private pruneSelectedItems(): void {
    const valid = this.getValidModelIdMap(this.selectedItems);
    clearMap(this.selectedItems);
    Object.assign(this.selectedItems, valid);
  }

  private async selectWholeModel(modelId: string): Promise<void> {
    const index = this.modelIndices.get(modelId);
    if (!index || index.allIds.size === 0) {
      this.setStatus(t('status.modelIndexNotReady'));
      return;
    }
    clearMap(this.selectedItems);
    this.selectedItems[modelId] = new Set(index.allIds);
    await this.refreshSelectionVisuals();
    await this.zoomToItems(this.selectedItems);
    this.setStatus(t('status.selectedFullModel', { count: index.allIds.size }));
  }

  private toggleModelVisibility(modelId: string): void {
    const model = this.federatedModels.get(modelId);
    if (!model) return;
    model.visible = !model.visible;
    model.object.visible = model.visible;
    if (!model.visible && this.activeGizmoModelId === modelId) this.detachModelGizmo();
    this.applyXRay();
    if (this.edgesEnabled) this.applyEdges();
    this.fireAndForget(this.updateFragments(true), 'Toggle visibility');
    this.fireAndForget(this.updateVisibilityCount(), 'Update visibility');
    this.renderModelBrowser();
    this.renderFederatedTree();
    this.persistLocalState();
    this.setStatus(model.visible ? t('status.modelShown', { name: model.fileName }) : t('status.modelHidden', { name: model.fileName }));
  }

  private applyModelOpacity(modelId: string, opacity: number): void {
    const model = this.federatedModels.get(modelId);
    if (!model) return;
    model.opacity = this.clamp(opacity, 0, 1);
    this.applyXRay();
    if (this.edgesEnabled) this.applyEdges();
    this.fireAndForget(this.updateFragments(true), 'Adjust opacity');
    this.persistLocalState();
  }

  private toggleModelGizmo(modelId: string): void {
    const model = this.federatedModels.get(modelId);
    if (!model || !this.transformControls) return;
    if (!model.visible) {
      this.setStatus(t('status.showModelBeforeGizmo'));
      return;
    }

    if (this.activeGizmoModelId === modelId && this.transformControlsHelper?.visible) {
      this.detachModelGizmo();
      this.setStatus(t('status.gizmoDetached'));
      return;
    }

    this.activeGizmoModelId = modelId;
    this.transformControls.attach(model.object);
    this.transformControls.setMode('translate');
    this.transformControls.enabled = true;
    if (this.transformControlsHelper) this.transformControlsHelper.visible = true;
    this.renderFederatedTree();
    this.setStatus(t('status.gizmoActive', { name: model.fileName }));
  }

  private detachModelGizmo(): void {
    if (!this.transformControls) return;
    this.transformControls.detach();
    if (this.transformControlsHelper) this.transformControlsHelper.visible = false;
    this.transformControls.enabled = false;
    this.activeGizmoModelId = null;
    this.gizmoDragging = false;
    this.world.camera.controls.enabled = true;
    this.renderFederatedTree();
  }

  private updateModelOffsetsFromObject(model: FederatedModelRecord): void {
    model.offsetPosition.x = model.object.position.x - model.basePosition.x;
    model.offsetPosition.y = model.object.position.y - model.basePosition.y;
    model.offsetPosition.z = model.object.position.z - model.basePosition.z;
    model.offsetRotation.x = THREE.MathUtils.radToDeg(model.object.rotation.x - model.baseRotation.x);
    model.offsetRotation.y = THREE.MathUtils.radToDeg(model.object.rotation.y - model.baseRotation.y);
    model.offsetRotation.z = THREE.MathUtils.radToDeg(model.object.rotation.z - model.baseRotation.z);
  }

  private applyTransformInput(input: HTMLInputElement): void {
    const modelId = input.dataset.modelId;
    const transform = input.dataset.transform;
    if (!modelId || !transform) return;

    const model = this.federatedModels.get(modelId);
    if (!model) return;

    const value = Number(input.value);
    if (!Number.isFinite(value)) {
      input.value = transform.startsWith('r')
        ? model.offsetRotation[transform.charAt(1).toLowerCase() as keyof TransformVector3].toFixed(1)
        : model.offsetPosition[transform.charAt(1).toLowerCase() as keyof TransformVector3].toFixed(2);
      return;
    }

    switch (transform) {
      case 'px':
        model.offsetPosition.x = value;
        break;
      case 'py':
        model.offsetPosition.y = value;
        break;
      case 'pz':
        model.offsetPosition.z = value;
        break;
      case 'rx':
        model.offsetRotation.x = value;
        break;
      case 'ry':
        model.offsetRotation.y = value;
        break;
      case 'rz':
        model.offsetRotation.z = value;
        break;
      default:
        return;
    }

    this.applyModelTransform(model);
    this.persistLocalState();
    this.setStatus(t('status.transformUpdated', { name: model.fileName }));
  }

  private applyModelTransform(model: FederatedModelRecord): void {
    model.object.position.set(
      model.basePosition.x + model.offsetPosition.x,
      model.basePosition.y + model.offsetPosition.y,
      model.basePosition.z + model.offsetPosition.z,
    );
    model.object.rotation.set(
      model.baseRotation.x + THREE.MathUtils.degToRad(model.offsetRotation.x),
      model.baseRotation.y + THREE.MathUtils.degToRad(model.offsetRotation.y),
      model.baseRotation.z + THREE.MathUtils.degToRad(model.offsetRotation.z),
    );
    model.object.updateMatrixWorld(true);
    if (this.activeGizmoModelId === model.modelId && this.transformControls && this.transformControls.object !== model.object) {
      this.transformControls.attach(model.object);
    }

    if (this.edgesEnabled) this.applyEdges();
    this.fireAndForget(this.updateFragments(true), 'Apply transform');
  }

  private resetModelOffsets(modelId: string): void {
    const model = this.federatedModels.get(modelId);
    if (!model) return;
    model.offsetPosition = { x: 0, y: 0, z: 0 };
    model.offsetRotation = { x: 0, y: 0, z: 0 };
    this.applyModelTransform(model);
    this.renderModelBrowser();
    this.renderFederatedTree();
    this.persistLocalState();
    this.setStatus(t('status.transformReset', { name: model.fileName }));
  }

  private fitToModelById(modelId: string): void {
    const model = this.federatedModels.get(modelId);
    if (!model) return;
    const bbox = new THREE.Box3().setFromObject(model.object);
    if (bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    const basis = new THREE.Quaternion();
    model.object.getWorldQuaternion(basis);
    this.navigateToDirection(center, maxDim, basis, VIEW_CUBE_HOME_VECTOR);
  }

  private async isolateLevelForModel(modelId: string, level: string): Promise<void> {
    const index = this.modelIndices.get(modelId);
    const ids = index?.levels.get(level);
    if (!ids || ids.size === 0) {
      this.setStatus(t('status.noElementsForLevel', { level }));
      return;
    }
    const map = this.getValidModelIdMap({ [modelId]: new Set(ids) });
    if (isMapEmpty(map)) {
      this.setStatus(t('status.levelUnavailable', { level }));
      return;
    }
    await this.hider.isolate(map);
    await this.updateVisibilityCount();
    this.setStatus(t('status.isolatedLevel', { level }));
  }

  private async isolateClassForModelLevel(modelId: string, level: string, className: string): Promise<void> {
    const ids = getClassIdsForModelLevel(this.modelIndices, modelId, level, className);
    if (ids.size === 0) {
      this.setStatus(t('status.noClassInLevel', { class: className, level }));
      return;
    }

    const map = this.getValidModelIdMap({ [modelId]: ids });
    if (isMapEmpty(map)) {
      this.setStatus(t('status.classInLevelUnavailable', { class: className, level }));
      return;
    }
    await this.hider.isolate(map);
    await this.updateVisibilityCount();
    this.setStatus(t('status.isolatedClassInLevel', { class: className, level }));
  }

  private collectCheckedValues(container: HTMLElement): string[] {
    const values: string[] = [];
    container.querySelectorAll<HTMLButtonElement>('.filter-chip.is-active[data-filter-value]').forEach((chip) => {
      if (chip.dataset.filterValue) values.push(chip.dataset.filterValue);
    });
    return values;
  }

  private mapFromClassFilters(selectedClasses: string[]): OBC.ModelIdMap {
    const map: OBC.ModelIdMap = {};
    for (const className of selectedClasses) {
      for (const [modelId, index] of this.modelIndices.entries()) {
        const localIds = index.classes.get(className);
        if (!localIds || localIds.size === 0) continue;
        if (!map[modelId]) map[modelId] = new Set<number>();
        for (const localId of localIds) map[modelId].add(localId);
      }
    }
    return map;
  }

  private mapFromLevelFilters(selectedLevels: string[]): OBC.ModelIdMap {
    const map: OBC.ModelIdMap = {};
    for (const levelName of selectedLevels) {
      for (const [modelId, index] of this.modelIndices.entries()) {
        const localIds = index.levels.get(levelName);
        if (!localIds || localIds.size === 0) continue;
        if (!map[modelId]) map[modelId] = new Set<number>();
        for (const localId of localIds) map[modelId].add(localId);
      }
    }
    return map;
  }

  private async applyFilters(): Promise<void> {
    const selectedClasses = this.collectCheckedValues(this.dom.classFilterList);
    const selectedLevels = this.collectCheckedValues(this.dom.levelFilterList);

    if (selectedClasses.length === 0 && selectedLevels.length === 0) {
      await this.hider.set(true);
      await this.updateVisibilityCount();
      this.setStatus(t('status.noFiltersSelected'));
      return;
    }

    const classMap = this.mapFromClassFilters(selectedClasses);
    const levelMap = this.mapFromLevelFilters(selectedLevels);

    let effectiveMap: OBC.ModelIdMap;
    if (!isMapEmpty(classMap) && !isMapEmpty(levelMap)) effectiveMap = intersectMaps(classMap, levelMap);
    else if (!isMapEmpty(classMap)) effectiveMap = classMap;
    else effectiveMap = levelMap;
    effectiveMap = this.getValidModelIdMap(effectiveMap);

    // F11: a disjoint class∩level combination used to silently hide the
    // entire model — warn and keep the current visibility instead.
    if (isMapEmpty(effectiveMap)) {
      this.showToast(t('toast.filtersNothingHidden'), 'warning');
      this.setStatus(t('status.filtersNoCommon'));
      return;
    }

    await this.hider.set(false);
    await this.hider.set(true, effectiveMap);
    await this.updateVisibilityCount();
    this.setStatus(t('status.filtersApplied'));
  }

  private async searchElements(term: string): Promise<void> {
    if (!term) {
      this.dom.elementResults.replaceChildren();
      this.dom.searchResultsGroup.hidden = true;
      return;
    }

    const escaped = escapeRegExp(term);
    const pattern = new RegExp(escaped, 'i');
    const results: SearchResult[] = [];

    for (const [modelKey, model] of this.fragments.list.entries()) {
      const modelId = String(modelKey);
      const ids = await model.getItemsByQuery({
        attributes: {
          aggregation: 'inclusive',
          queries: [{ name: /Name|GlobalId|ObjectType|PredefinedType/i, value: pattern }],
        },
      });

      if (!ids || ids.length === 0) continue;
      const trimmed = ids.slice(0, 60);
      const itemsData = await model.getItemsData(trimmed, {
        attributesDefault: true,
        attributes: ['Name', 'GlobalId', 'ObjectType', 'PredefinedType'],
      });

      for (let i = 0; i < trimmed.length; i += 1) {
        const localId = trimmed[i];
        const data = (itemsData[i] || {}) as Record<string, unknown>;
        // F1: fragments v3.3 returns {value,type} ItemAttribute objects —
        // unwrap to primitives before they reach escapeHtml/rendering.
        const name = readPrimitiveValue(data.Name) || t('label.elementFallback', { id: localId });
        const type = readPrimitiveValue(data.ObjectType) || readPrimitiveValue(data.PredefinedType) || t('label.itemFallback');
        const globalId = readPrimitiveValue(data.GlobalId) || '-';
        results.push({ modelId, localId, name, type, globalId });
      }
    }

    const capped = results.slice(0, 180);
    // U6: real <button> rows (keyboard-reachable), delegated selection.
    this.dom.elementResults.innerHTML = capped.length === 0
      ? `<div class="list-empty">${escapeHtml(t('empty.noMatches'))}</div>`
      : capped
          .map((result) => `
        <button type="button" class="result-row" data-model-id="${escapeHtml(result.modelId)}" data-local-id="${result.localId}">
          <span class="result-name">${escapeHtml(result.name)}</span>
          <span class="result-meta">${escapeHtml(result.type)} · ${escapeHtml(result.globalId)}</span>
        </button>
      `)
          .join('');
    this.dom.searchResultsGroup.hidden = false;

    this.dom.elementResults.querySelectorAll<HTMLButtonElement>('.result-row').forEach((item) => {
      item.addEventListener('click', () => {
        const modelId = item.dataset.modelId;
        const localId = Number(item.dataset.localId);
        if (!modelId || Number.isNaN(localId)) return;
        this.fireAndForget(this.selectSingleItem(modelId, localId, true), 'Select search result');
      });
    });

    this.setStatus(t('status.searchMatches', { count: results.length }));
  }

  private async loadIfcFiles(files: File[]): Promise<void> {
    const ifcFiles = files.filter((file) => file.name.toLowerCase().endsWith('.ifc'));
    if (ifcFiles.length === 0) {
      this.showToast(t('toast.onlyIfc'), 'warning');
      this.setStatus(t('toast.onlyIfc'));
      return;
    }
    if (this.isModelLoading) {
      this.showToast(t('status.modelAlreadyLoading'), 'warning');
      this.setStatus(t('status.modelAlreadyLoading'));
      return;
    }

    const failedFiles: File[] = [];
    let lastError = '';
    const batchTotal = ifcFiles.length;
    if (batchTotal > 1) this.suppressAutoFit = true;

    for (let i = 0; i < ifcFiles.length; i += 1) {
      const file = ifcFiles[i];
      const result = await this.loadIfcFile(file, i + 1, batchTotal);
      if (!result.success) {
        failedFiles.push(file);
        lastError = result.error ?? lastError;
      }
    }

    this.suppressAutoFit = false;

    if (batchTotal > 1 && this.modelObjects.length > 0) this.fitToModel();

    if (failedFiles.length === 0) {
      this.setStatus(batchTotal > 1 ? t('status.batchLoaded', { count: batchTotal }) : t('status.modelLoadedOk'));
      return;
    }

    // U4: failures surface as a toast + overlay error state with Retry.
    this.lastFailedLoadFiles = failedFiles;
    const failedNames = failedFiles.map((file) => file.name).join(', ');
    const summary = failedFiles.length === batchTotal
      ? t('status.batchFailedAll', { names: failedNames })
      : t('status.batchFailedSome', { ok: batchTotal - failedFiles.length, total: batchTotal, names: failedNames });
    this.showToast(summary, 'error', 6000);
    this.showLoadError(lastError ? `${summary} (${lastError})` : summary);
    this.setStatus(summary);
  }

  /** U4: switches the loading overlay into its error state (Retry/Dismiss). */
  private showLoadError(message: string): void {
    this.dom.loadingOverlay.hidden = false;
    this.dom.loadingOverlay.classList.add('is-error');
    this.dom.loadingText.textContent = message;
    this.dom.loadingProgress.style.width = '0%';
    this.dom.loadingErrorActions.hidden = false;
  }

  private hideLoadError(): void {
    this.dom.loadingOverlay.classList.remove('is-error');
    this.dom.loadingErrorActions.hidden = true;
    this.dom.loadingOverlay.hidden = true;
    this.dom.emptyState.hidden = this.modelObjects.length > 0;
  }

  private async loadIfcFile(
    file: File,
    batchIndex = 1,
    batchTotal = 1,
  ): Promise<{ success: boolean; error?: string }> {
    if (this.isModelLoading) {
      this.setStatus(t('status.modelAlreadyLoading'));
      return { success: false, error: 'A model is already loading' };
    }

    this.isModelLoading = true;
    const requestId = ++this.loadRequestId;
    this.dom.btnUpload.disabled = true;
    this.dom.btnUploadEmpty.disabled = true;
    this.dom.fileInput.disabled = true;

    this.dom.emptyState.hidden = true;
    this.dom.loadingOverlay.hidden = false;
    this.dom.loadingOverlay.classList.remove('is-error');
    this.dom.loadingErrorActions.hidden = true;
    this.dom.loadingText.textContent = batchTotal > 1
      ? t('load.parsingBatch', { index: batchIndex, total: batchTotal })
      : t('load.parsing');
    this.dom.loadingProgress.style.width = '8%';
    this.dom.loadingPct.textContent = '8%';
    this.setStatus(t('status.loadingModel'));

    let timeoutHandle: number | undefined;
    // Metadata is keyed by the id passed to ifcLoader.load — it becomes
    // model.modelId and the fragments.list key (A6). Duplicate file names get
    // a unique suffix so federated loads never collide.
    const modelId = this.modelRegistry.allocateModelId(file.name, [
      ...(this.fragments?.list?.keys() ?? []),
      ...this.federatedModels.keys(),
    ]);

    try {
      const start = performance.now();
      const buffer = await file.arrayBuffer();
      const data = new Uint8Array(buffer);
      // Deterministic guard before web-ifc (T13): reject non-IFC input fast and
      // identically on every platform, instead of relying on web-ifc's
      // undefined behavior for garbage (throw / hang / silent empty model).
      if (!isProbablyIfc(data)) {
        throw new Error(t('status.notValidIfc'));
      }
      this.modelRegistry.beginLoad(modelId, { fileName: file.name, sizeBytes: file.size });

      if (requestId === this.loadRequestId) {
        this.dom.loadingText.textContent = t('load.converting');
        this.dom.loadingProgress.style.width = '25%';
        this.dom.loadingPct.textContent = '25%';
      }

      // P4 (W5.4): convert IFC→fragments in the dedicated worker (off the main
      // thread), then stream-load the resulting `.frag` via fragments.core.load.
      // Behaviour matches the previous ifcLoader.load: addAllAttributes/Relations
      // for a full property/spatial index, same 25→95% progress mapping.
      const wasmPath = new URL(import.meta.env.BASE_URL, window.location.origin).href;
      const loadPromise = (async (): Promise<FragmentsModelLike> => {
        const fragBytes = await this.ifcConversionClient.convert(data, {
          wasmPath,
          addAllAttributes: true,
          addAllRelations: true,
          onProgress: (progress: number) => {
            if (requestId !== this.loadRequestId) return;
            const percentage = Math.round(25 + progress * 70);
            this.dom.loadingProgress.style.width = `${percentage}%`;
            this.dom.loadingPct.textContent = `${percentage}%`;
            this.dom.loadingText.textContent = t('load.buildingIndex');
          },
        });
        // C8: the converted bytes are already in hand — cache them now so a fresh
        // load is restorable without re-hashing later (onModelAdded reuses the key).
        const fragKey = hashFragBytes(fragBytes);
        this.restoreFragKeys.set(modelId, fragKey);
        this.fireAndForget(
          this.fragCache.put({
            key: fragKey,
            fileName: file.name,
            sizeBytes: fragBytes.byteLength,
            storedAt: Date.now(),
            bytes: fragBytes,
          }),
          'Cache fragments',
        );
        return this.fragments.core.load(fragBytes, { modelId });
      })();

      let timedOut = false;
      const timeoutPromise = new Promise<never>((_resolve, reject) => {
        timeoutHandle = window.setTimeout(() => {
          timedOut = true;
          reject(new Error('Model loading timed out. Please try again with a smaller IFC or reload the page.'));
        }, 120000);
      });

      let loadedModel: FragmentsModelLike;
      try {
        loadedModel = await Promise.race([loadPromise, timeoutPromise]);
      } catch (error) {
        if (timedOut) {
          // A10 + W2 (W5-fixups): the worker keeps converting after the race is
          // lost. Cancel it — terminate + null the worker to free the abandoned
          // web-ifc parse and reject the abandoned pending job, so a retry spawns
          // a fresh worker instead of queueing behind (or reusing) a busy one, and
          // stale progress callbacks stop mutating the loading overlay.
          this.modelRegistry.markStale(modelId);
          this.ifcConversionClient.cancel('Model loading timed out');
          loadPromise
            .then(() => this.disposeStaleModel(modelId))
            .catch(() => {
              // The abandoned load was cancelled/failed — nothing arrived to
              // dispose, so consume the stale flag (keeps disposeStaleModel sound).
              this.modelRegistry.consumeStale(modelId);
            });
        }
        throw error;
      }
      await this.registerModel(modelId, loadedModel);

      if (timeoutHandle !== undefined) {
        window.clearTimeout(timeoutHandle);
        timeoutHandle = undefined;
      }

      const elapsed = ((performance.now() - start) / 1000).toFixed(1);
      if (requestId !== this.loadRequestId) return { success: true };

      // A15: dedicated slot — the FPS monitor owns perfInfo.
      this.dom.loadInfo.textContent = t('label.loadInfo', { seconds: elapsed, size: (file.size / 1024 / 1024).toFixed(1) });
      this.dom.loadingProgress.style.width = '100%';
      this.setStatus(t('status.modelLoadedOk'));

      setTimeout(() => {
        if (requestId === this.loadRequestId) {
          this.dom.loadingOverlay.hidden = true;
        }
      }, 220);
      return { success: true };
    } catch (error) {
      // No-op when the id was already marked stale (timeout path).
      this.modelRegistry.failLoad(modelId);

      if (timeoutHandle !== undefined) {
        window.clearTimeout(timeoutHandle);
      }
      const message = serializeError(error);
      console.error(error);
      if (requestId !== this.loadRequestId) return { success: false, error: message };

      this.dom.loadingOverlay.hidden = true;
      this.dom.emptyState.hidden = this.modelObjects.length > 0;
      this.setStatus(t('status.loadFailed', { message }));
      return { success: false, error: message };
    } finally {
      if (requestId === this.loadRequestId) {
        this.isModelLoading = false;
        this.dom.btnUpload.disabled = false;
        this.dom.btnUploadEmpty.disabled = false;
        this.dom.fileInput.disabled = false;
      }
    }
    return { success: false };
  }

  private getModelBoundingBox(): THREE.Box3 | null {
    const bbox = new THREE.Box3();
    // Prefer each fragments model's data-driven box (A17): it is valid right
    // after load, whereas expandByObject(object) reads empty until the worker
    // has streamed meshes — which broke Fit/Section immediately post-load and
    // on slow GPUs. Transform the local box by the object's world matrix so
    // federation offsets/rotations are included. Fall back to the streamed
    // geometry bounds if a model exposes no usable data box.
    for (const record of this.federatedModels.values()) {
      if (!record.visible) continue;
      const model = this.getFragmentsModel(record.modelId);
      const dataBox = model?.box;
      if (dataBox && !dataBox.isEmpty()) {
        record.object.updateWorldMatrix(true, false);
        bbox.union(dataBox.clone().applyMatrix4(record.object.matrixWorld));
      } else {
        bbox.expandByObject(record.object);
      }
    }
    return bbox.isEmpty() ? null : bbox;
  }

  private getViewCubeAnchorModel(): FederatedModelRecord | null {
    if (this.activeGizmoModelId) {
      const activeModel = this.federatedModels.get(this.activeGizmoModelId);
      if (activeModel?.visible) return activeModel;
    }
    for (const model of this.federatedModels.values()) {
      if (model.visible) return model;
    }
    return this.federatedModels.values().next().value ?? null;
  }

  private getViewCubeBasisQuaternion(target = new THREE.Quaternion()): THREE.Quaternion {
    const anchorModel = this.getViewCubeAnchorModel();
    if (!anchorModel) return target.identity();
    anchorModel.object.getWorldQuaternion(target);
    return target.normalize();
  }

  private navigateToDirection(
    center: THREE.Vector3,
    maxDim: number,
    basis: THREE.Quaternion,
    directionTuple: readonly [number, number, number],
  ): void {
    if (!this.world?.camera) return;
    const localDirection = new THREE.Vector3(directionTuple[0], directionTuple[1], directionTuple[2]).normalize();
    const worldDirection = localDirection.clone().applyQuaternion(basis).normalize();
    const distance = getViewCubeNavigationDistance(maxDim, localDirection);
    const axes = getViewCubeAxes(basis);
    const up = resolveViewCubeCameraUp(localDirection, worldDirection, axes);
    const eye = center.clone().addScaledVector(worldDirection, distance);
    this.world.camera.three.up.copy(up);
    this.world.camera.three.updateMatrixWorld();
    void this.world.camera.controls.setLookAt(eye.x, eye.y, eye.z, center.x, center.y, center.z, true);
  }

  private navigateToView(vector: readonly [number, number, number], view: 'orbit' | 'front' | 'top'): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    this.navigateToDirection(center, maxDim, this.getViewCubeBasisQuaternion(new THREE.Quaternion()), vector);
    this.setActiveView(view);
  }

  private fitToModel(): void {
    this.navigateToView(VIEW_CUBE_HOME_VECTOR, 'orbit');
  }

  private setFrontView(): void {
    this.navigateToView(FRONT_VIEW_VECTOR, 'front');
  }

  private setTopView(): void {
    this.navigateToView(TOP_VIEW_VECTOR, 'top');
  }

  private setHomeView(): void {
    // A12: home view == fit-all (both frame the model along the home vector).
    this.fitToModel();
  }

  /** Reflects the active preset in the glass view controls + status label. */
  private setActiveView(view: 'orbit' | 'front' | 'top'): void {
    this.activeView = view;
    this.dom.cubeHome.classList.toggle('is-active', view === 'orbit');
    this.dom.btnFront.classList.toggle('is-active', view === 'front');
    this.dom.btnTop.classList.toggle('is-active', view === 'top');
    this.dom.viewLabel.textContent =
      view === 'orbit' ? t('label.viewOrbit') : view === 'front' ? t('label.viewFront') : t('label.viewTop');
  }

  private applySelectionMode(mode: SelectionMode): void {
    this.selectionMode = mode;
    this.setRailPressed(this.dom.btnSelectSingle, mode === 'single');
    this.setRailPressed(this.dom.btnSelectMulti, mode === 'multi');
    this.setStatus(mode === 'single' ? t('status.selectionModeSingle') : t('status.selectionModeMulti'));
    this.persistLocalState();
  }

  private applyNavigationMode(mode: NavigationMode): void {
    if (!this.world?.camera) return;
    this.navigationMode = mode;
    this.world.camera.set(mode);
    this.setRailPressed(this.dom.btnModeOrbit, mode === 'Orbit');
    this.setRailPressed(this.dom.btnModePlan, mode === 'Plan');
    this.setRailPressed(this.dom.btnModeFirstPerson, mode === 'FirstPerson');
    this.persistLocalState();
  }

  private setMeasureMode(mode: MeasureMode): void {
    this.measureMode = mode;
    this.lengthMeasurement.enabled = mode === 'length';
    this.areaMeasurement.enabled = mode === 'area';
    this.setRailPressed(this.dom.btnMeasureLength, mode === 'length');
    this.setRailPressed(this.dom.btnMeasureArea, mode === 'area');
    this.syncMeasureHint();
    this.setStatus(
      mode === 'none'
        ? t('status.measureDisabled')
        : t('status.measureEnabled', { mode: t(mode === 'length' ? 'measure.length' : 'measure.area') }),
    );
  }

  /** Shows/localizes the measure-hint overlay for the current mode. */
  private syncMeasureHint(): void {
    this.dom.measureHint.hidden = this.measureMode === 'none';
    if (this.measureMode !== 'none') {
      this.dom.measureHintText.textContent =
        this.measureMode === 'area' ? t('label.measureAreaHint') : t('label.measureLengthHint');
    }
  }

  private clearMeasurements(): void {
    this.lengthMeasurement.list.clear();
    this.areaMeasurement.list.clear();
    this.lengthMeasurement.cancelCreation();
    this.areaMeasurement.cancelCreation();
    this.setMeasureMode('none');
    // R3 (W5-fixups): removing the lines/labels mutates the scene with no camera
    // motion, so MANUAL mode needs an explicit repaint to clear the stale frame.
    this.requestRender();
  }

  /**
   * F10: the single clip-plane creation path — every caller (section buttons
   * AND viewpoint restore) gets the gizmo visibility fix.
   */
  private createClipPlane(normal: THREE.Vector3, point: THREE.Vector3): void {
    this.clipper.enabled = true;
    const planeId = this.clipper.createFromNormalAndCoplanarPoint(this.world, normal, point);
    // Exempt the gizmo from clipping and depth-test so the arrow is always visible
    // (renders on top of model geometry and not clipped by section planes)
    const plane = this.clipper.list.get(planeId);
    if (plane) {
      const renderer = this.world.renderer!.three;
      renderer.localClippingEnabled = true;
      const gizmoHelper = getClipperPlaneGizmoHelper(plane);
      gizmoHelper?.traverse((child: THREE.Object3D) => {
        const mesh = child as THREE.Mesh;
        if (mesh.material) {
          const mats = Array.isArray(mesh.material) ? mesh.material : [mesh.material];
          mats.forEach((m: THREE.Material) => {
            m.clippingPlanes = [];
            m.depthTest = false;
          });
        }
        child.renderOrder = 999;
      });
    }
    this.requestRender();
  }

  // Active single-plane section state, for the glass slider (position + normal).
  private activeSectionNormal: THREE.Vector3 | null = null;
  private activeSectionBox: THREE.Box3 | null = null;
  private activeSectionLabel = '';

  private addSectionPlane(normal: THREE.Vector3): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) {
      this.setStatus(t('status.noModelToSection'));
      return;
    }
    const center = bbox.getCenter(new THREE.Vector3());
    this.createClipPlane(normal, center);
    this.setStatus(t('status.sectionPlaneAdded'));
  }

  /** A12: shared toggle for the X/Y/Z axis section buttons. */
  private toggleSectionPlane(button: HTMLButtonElement, normal: THREE.Vector3): void {
    if (button.classList.contains('is-active')) {
      this.clearSections();
      return;
    }
    this.clearSections(false);
    this.addSectionPlane(normal);
    this.setRailPressed(button, true);
    // Wire the glass slider to this axis (design: bottom-center section control).
    this.activeSectionNormal = normal.clone();
    this.activeSectionBox = this.getModelBoundingBox();
    this.activeSectionLabel =
      button === this.dom.btnSectionX ? 'Section X' : button === this.dom.btnSectionY ? 'Section Y' : 'Section Z';
    this.dom.sectionLabel.textContent = this.activeSectionLabel;
    this.dom.sectionPos.value = '50';
    this.dom.sectionPosLabel.textContent = t('label.percent', { value: 50 });
    this.dom.sectionSlider.hidden = false;
  }

  /** Moves the active section plane along its normal to the given percentage. */
  private setSectionPosition(pct: number): void {
    if (!this.activeSectionNormal || !this.activeSectionBox || this.activeSectionBox.isEmpty()) return;
    const point = sectionPlanePoint(this.activeSectionBox, this.activeSectionNormal, pct);
    this.clipper.deleteAll();
    this.createClipPlane(this.activeSectionNormal.clone(), point);
    this.fireAndForget(this.updateFragments(true), 'Section move');
  }

  private createSectionBox(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) {
      this.setStatus(t('status.noModelToSection'));
      return;
    }
    this.clearSections(false);
    this.clipper.enabled = true;
    for (const { normal, point } of sectionBoxPlanes(bbox)) {
      this.clipper.createFromNormalAndCoplanarPoint(this.world, normal, point);
    }
    this.requestRender();
    this.setStatus(t('status.sectionBoxCreated'));
  }

  private clearSections(updateStatus = true): void {
    this.clipper.deleteAll();
    this.clipper.enabled = false;
    for (const btn of [this.dom.btnSectionX, this.dom.btnSectionY, this.dom.btnSectionZ, this.dom.btnSectionBox]) {
      this.setRailPressed(btn, false);
    }
    this.activeSectionNormal = null;
    this.activeSectionBox = null;
    this.dom.sectionSlider.hidden = true;
    this.requestRender();
    if (updateStatus) this.setStatus(t('status.sectionsCleared'));
  }

  private applyXRay(): void {
    for (const record of this.federatedModels.values()) {
      const object = record.object;
      object.visible = record.visible;
      const model = this.getFragmentsModel(record.modelId);
      if (!model) {
        this.appliedModelOpacity.delete(record.modelId);
        continue;
      }

      const targetOpacity = computeXrayOpacity(record.opacity, this.xrayEnabled);
      const previousOpacity = this.appliedModelOpacity.get(record.modelId) ?? 1;
      if (Math.abs(previousOpacity - targetOpacity) < 0.001) continue;

      this.appliedModelOpacity.set(record.modelId, targetOpacity);
      if (targetOpacity >= 0.999) {
        if (typeof model?.resetOpacity === 'function') {
          this.fireAndForget(model.resetOpacity(undefined), `Reset model opacity (${record.fileName})`);
        }
        continue;
      }
      if (typeof model?.setOpacity === 'function') {
        this.fireAndForget(model.setOpacity(undefined, targetOpacity), `Set model opacity (${record.fileName})`);
      }
    }
  }

  private applyEdges(): void {
    // P3: overlay geometries are owned by edgeGeometryCache (shared, reused),
    // so remove overlays from the scene but do NOT dispose their geometry here.
    for (const overlay of this.edgeOverlays) {
      this.world.scene.three.remove(overlay);
    }
    this.edgeOverlays.length = 0;
    if (!this.edgesEnabled) return;

    this.edgeOverlays = buildEdgeOverlays(this.modelObjects, this.edgeMaterial, this.edgeGeometryCache);
    for (const overlay of this.edgeOverlays) this.world.scene.three.add(overlay);
    this.requestRender();
  }

  private async onViewerClick(_event: MouseEvent): Promise<void> {
    await this.pickAndSelect();
  }

  /**
   * The canvas selection path: raycast (at `position` if given, else the last
   * pointer position), then apply measure/issue/single/multi selection exactly
   * as a real click would. Returns the hit item, or null. `position` is used by
   * the e2e test API (T6) to pick at explicit coordinates; production clicks
   * pass nothing and behave exactly as before.
   */
  private async pickAndSelect(position?: THREE.Vector2): Promise<{ modelId: string; localId: number } | null> {
    if (this.pointerDragged || this.gizmoDragging || !!this.activeGizmoModelId) return null;

    if (this.measureMode !== 'none') {
      if (this.measureMode === 'length') await this.lengthMeasurement.create();
      if (this.measureMode === 'area') await this.areaMeasurement.create();
      this.requestRender();
      return null;
    }

    const result = await this.raycaster.castRay(position ? { position } : undefined) as any;
    if (!result || !result.fragments || result.localId === undefined) {
      if (this.selectionMode === 'single' && !this.issuePinMode) await this.clearSelection();
      return null;
    }

    const modelId = String(result.fragments.modelId);
    if (!this.isLoadedModelId(modelId)) return null;
    const localId = result.localId as number;
    if (result.point) this.lastHitPoint = result.point.clone();

    if (this.issuePinMode) {
      if (result.point) this.pendingIssuePoint = result.point.clone();
      if (this.selectionMode === 'single') await this.selectSingleItem(modelId, localId, false);
      this.activateTab('issues');
      this.setStatus(t('status.issuePointCaptured'));
      return { modelId, localId };
    }

    if (this.selectionMode === 'single') {
      await this.selectSingleItem(modelId, localId, false);
      return { modelId, localId };
    }

    this.toggleSelectionItem(modelId, localId);
    await this.refreshSelectionVisuals();
    return { modelId, localId };
  }

  private async selectSingleItem(modelId: string, localId: number, zoomToItem: boolean): Promise<void> {
    if (!this.isLoadedModelId(modelId)) return;
    clearMap(this.selectedItems);
    this.selectedItems[modelId] = new Set([localId]);
    await this.refreshSelectionVisuals();
    if (zoomToItem) await this.zoomToItems(this.selectedItems);
  }

  private toggleSelectionItem(modelId: string, localId: number): void {
    if (!this.isLoadedModelId(modelId)) return;
    if (!this.selectedItems[modelId]) this.selectedItems[modelId] = new Set<number>();
    const set = this.selectedItems[modelId];
    if (set.has(localId)) set.delete(localId);
    else set.add(localId);
    if (set.size === 0) delete this.selectedItems[modelId];
  }

  private async clearSelection(): Promise<void> {
    clearMap(this.selectedItems);
    await this.refreshSelectionVisuals();
  }

  private async refreshSelectionVisuals(): Promise<void> {
    this.pruneSelectedItems();
    const validSelection = this.getValidModelIdMap(this.selectedItems);
    clearMap(this.selectedItems);
    Object.assign(this.selectedItems, validSelection);

    await this.fragments.resetHighlight();
    if (!isMapEmpty(validSelection)) {
      await this.fragments.highlight(
        { color: new THREE.Color(0xc8145c), transparent: true, opacity: 0.88 } as any,
        cloneMap(validSelection),
      );
    }
    await this.updateFragments(true);
    this.updateCounters();
    await this.updatePropertiesPanel();
  }

  private getFirstSelection(): { modelId: string; localId: number } | null {
    for (const [modelId, ids] of Object.entries(this.selectedItems)) {
      for (const localId of ids) return { modelId, localId };
    }
    return null;
  }

  /** Rows of the model's unit entities (IfcSIUnit/IfcConversionBasedUnit). */
  private async fetchUnitRows(model: FragmentsModelLike): Promise<unknown[]> {
    const categories = await model.getItemsOfCategories([/IFCSIUNIT/, /IFCCONVERSIONBASEDUNIT/]);
    const ids = Object.values(categories).flat();
    if (ids.length === 0) return [];
    return await model.getItemsData(ids, {
      attributesDefault: true,
      relationsDefault: { attributes: false, relations: false },
    });
  }

  private renderPropertySections(sections: PropertySectionData[]): void {
    // Localize the section headings by their stable id (C7). The property engine
    // stays pure/English; only the displayed title is swapped here.
    const titleKeys = {
      identity: 'prop.identity',
      type: 'prop.type',
      dimensions: 'prop.dimensions',
      location: 'prop.location',
      levels: 'prop.levels',
      materials: 'prop.materials',
      quantities: 'prop.quantities',
      relations: 'prop.relations',
      raw: 'prop.raw',
    } as const;
    const localized = sections.map((section) => ({ ...section, title: t(titleKeys[section.id]) }));
    this.dom.propSections.innerHTML = buildPropertySectionsMarkup(localized);
  }

  private applyPropertiesFilter(): void {
    const filter = this.propertyFilterText;
    const sections = Array.from(this.dom.propSections.querySelectorAll<HTMLElement>('[data-prop-section]'));

    for (const section of sections) {
      const sectionSearch = section.dataset.search || '';
      const sectionMatches = !filter || sectionSearch.includes(filter);
      const rows = Array.from(section.querySelectorAll<HTMLElement>('[data-prop-row]'));
      let visibleCount = 0;

      for (const row of rows) {
        const rowSearch = row.dataset.search || '';
        const visible = sectionMatches || !filter || rowSearch.includes(filter);
        row.hidden = !visible;
        if (visible) visibleCount += 1;
      }

      section.hidden = visibleCount === 0;
      const count = section.querySelector<HTMLElement>('[data-prop-count]');
      if (count) count.textContent = String(visibleCount);
    }
  }

  private async updatePropertiesPanel(): Promise<void> {
    if (isMapEmpty(this.selectedItems)) {
      this.dom.propsEmpty.hidden = false;
      this.dom.propsContent.hidden = true;
      this.dom.propSections.innerHTML = '';
      this.dom.selectionChip.hidden = true;
      return;
    }

    const firstSelection = this.getFirstSelection();
    if (!firstSelection) return;
    const model = this.getFragmentsModel(firstSelection.modelId);
    if (!model) return;
    // F5: property suffixes use the selected element's model units.
    this.activePropertyUnits = this.modelUnits.get(firstSelection.modelId) ?? DEFAULT_MODEL_UNITS;

    const itemData = await model.getItemsData([firstSelection.localId], {
      attributesDefault: true,
      relationsDefault: { attributes: true, relations: true },
    });
    const data = (itemData[0] || {}) as Record<string, unknown>;

    const typeValue = [
      toPropertyString(data.ObjectType, ''),
      toPropertyString(data.PredefinedType, ''),
      toPropertyString(data.type, ''),
      toPropertyString(data._category, ''),
    ].find((entry) => entry.length > 0) || '-';

    const nameValue = toPropertyString(data.Name, '');
    const displayName = nameValue || t('label.elementFallback', { id: firstSelection.localId });
    const storey = this.modelIndices.get(firstSelection.modelId)?.itemToLevel.get(firstSelection.localId) || '—';
    this.dom.propType.textContent = typeValue;
    this.dom.propName.textContent = displayName;
    this.dom.propGlobalId.textContent = toPropertyString(data.GlobalId, '—');
    this.dom.propDescription.textContent = toPropertyString(data.Description, '—');
    this.dom.propStory.textContent = storey;

    // Selection chip (glass overlay, top-left).
    const count = countMapItems(this.selectedItems);
    this.dom.selChipName.textContent = count > 1 ? t('status.elementsSelected', { count }) : displayName;
    this.dom.selChipMeta.textContent = count > 1 ? typeValue : `${typeValue} · ${storey}`;
    this.dom.selectionChip.hidden = false;

    let volumeText = '';
    let geometryProbe: GeometryProbe | null = null;
    try {
      const volume = await model.getItemsVolume([firstSelection.localId]);
      volumeText = `${volume.toFixed(3)} m\u00B3`;
    } catch {
      // optional
    }

    try {
      const boxes = await this.fragments.getBBoxes({ [firstSelection.modelId]: new Set([firstSelection.localId]) });
      if (boxes.length > 0) {
        const bbox = new THREE.Box3();
        for (const box of boxes) bbox.union(box);
        const center = bbox.getCenter(new THREE.Vector3());
        const size = bbox.getSize(new THREE.Vector3());
        geometryProbe = {
          center: { x: center.x, y: center.y, z: center.z },
          size: { x: size.x, y: size.y, z: size.z },
        };
      }
    } catch {
      // optional
    }

    const indexStorey = this.modelIndices.get(firstSelection.modelId)?.itemToLevel.get(firstSelection.localId);
    const sections = buildPropertySections(
      data,
      firstSelection.localId,
      this.activePropertyUnits,
      indexStorey,
      volumeText,
      geometryProbe,
    );
    this.renderPropertySections(sections);
    this.applyPropertiesFilter();

    this.dom.propsEmpty.hidden = true;
    this.dom.propsContent.hidden = false;
    this.setStatus(t('status.elementsSelected', { count: countMapItems(this.selectedItems) }));
  }

  private async zoomToItems(modelIdMap: OBC.ModelIdMap): Promise<void> {
    const validMap = this.getValidModelIdMap(modelIdMap);
    if (isMapEmpty(validMap)) return;
    const boxes = await this.fragments.getBBoxes(validMap);
    if (boxes.length === 0) return;
    const bbox = new THREE.Box3();
    for (const box of boxes) bbox.union(box);
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    await this.world.camera.controls.setLookAt(
      center.x + maxDim,
      center.y + maxDim * 0.8,
      center.z + maxDim,
      center.x,
      center.y,
      center.z,
      true,
    );
  }

  private async saveViewpoint(): Promise<void> {
    const name = this.dom.viewpointName.value.trim();
    if (!name) {
      this.setStatus(t('status.enterViewpointName'));
      return;
    }

    const cameraPos = this.world.camera.three.position.clone();
    const target = new THREE.Vector3();
    this.world.camera.controls.getTarget(target);
    const hiddenItems = await this.hider.getVisibilityMap(false);

    const clippingPlanes = [...this.clipper.list.values()].map((plane) => ({
      normal: { x: plane.normal.x, y: plane.normal.y, z: plane.normal.z },
      origin: { x: plane.origin.x, y: plane.origin.y, z: plane.origin.z },
    }));

    // F2: downscaled JPEG thumbnail rendered immediately before capture.
    const snapshot = this.captureViewpointThumbnail();

    const viewpoint: SavedViewpoint = {
      id: uniqueId(),
      name,
      createdAt: new Date().toISOString(),
      camera: {
        position: { x: cameraPos.x, y: cameraPos.y, z: cameraPos.z },
        target: { x: target.x, y: target.y, z: target.z },
        projection: this.world.camera.projection.current,
        mode: this.navigationMode,
      },
      clippingPlanes,
      hiddenItems,
      visualStyle: this.visualStyle,
      xray: this.xrayEnabled,
      edges: this.edgesEnabled,
      snapshot,
    };

    this.viewpoints.unshift(viewpoint);
    this.selectedViewpointId = viewpoint.id;
    this.updateViewpointList();
    this.persistLocalState();
    this.dom.viewpointName.value = '';
    this.setStatus(t('status.viewpointSaved', { name }));
  }

  private async applySelectedViewpoint(): Promise<void> {
    if (!this.selectedViewpointId) {
      this.setStatus(t('status.selectViewpointFirst'));
      return;
    }

    const viewpoint = this.viewpoints.find((entry) => entry.id === this.selectedViewpointId);
    if (!viewpoint) {
      this.setStatus(t('status.viewpointNotFound'));
      return;
    }

    this.applyNavigationMode(viewpoint.camera.mode);

    if (this.world.camera.projection.current !== viewpoint.camera.projection) {
      await this.world.camera.projection.set(viewpoint.camera.projection);
    }

    await this.world.camera.controls.setLookAt(
      viewpoint.camera.position.x,
      viewpoint.camera.position.y,
      viewpoint.camera.position.z,
      viewpoint.camera.target.x,
      viewpoint.camera.target.y,
      viewpoint.camera.target.z,
      true,
    );

    this.clearSections(false);
    // F10: restore clipping through the shared creation path so the plane
    // gizmos stay visible exactly like the section buttons' planes.
    for (const plane of viewpoint.clippingPlanes) {
      this.createClipPlane(
        new THREE.Vector3(plane.normal.x, plane.normal.y, plane.normal.z),
        new THREE.Vector3(plane.origin.x, plane.origin.y, plane.origin.z),
      );
    }
    this.clipper.enabled = viewpoint.clippingPlanes.length > 0;

    await this.hider.set(true);
    const hiddenMap = this.getValidModelIdMap(toSetMap(viewpoint.hiddenItems));
    if (!isMapEmpty(hiddenMap)) await this.hider.set(false, hiddenMap);

    const viewpointStyle = this.parseVisualStyle(viewpoint.visualStyle ?? 'color-pen-shadows');
    await this.setVisualStyle(viewpointStyle, false, false);
    this.xrayEnabled = viewpoint.xray;
    this.edgesEnabled = viewpoint.edges;
    this.setRailPressed(this.dom.btnTransparency, this.xrayEnabled);
    this.setRailPressed(this.dom.btnWireframe, this.edgesEnabled);
    this.applyXRay();
    this.applyEdges();

    await this.updateVisibilityCount();
    this.setStatus(t('status.viewpointApplied', { name: viewpoint.name }));
  }

  private async deleteSelectedViewpoint(): Promise<void> {
    if (!this.selectedViewpointId) {
      this.setStatus(t('status.selectViewpointFirst'));
      return;
    }

    const confirmed = await this.confirm(t('confirm.deleteViewpoint'), t('confirm.delete'), t('confirm.cancel'));
    if (!confirmed) return;

    this.viewpoints = this.viewpoints.filter((entry) => entry.id !== this.selectedViewpointId);
    this.selectedViewpointId = this.viewpoints[0]?.id ?? null;
    this.updateViewpointList();
    this.persistLocalState();
    this.showToast(t('status.viewpointDeleted'), 'success');
    this.setStatus(t('status.viewpointDeleted'));
  }

  private updateViewpointList(): void {
    // A11: row + action clicks handled by delegation bound once in
    // bindViewpointListEvents() — no per-item listeners to leak on re-render.
    const markup = buildViewpointListMarkup(this.viewpoints, this.selectedViewpointId, {
      apply: t('vp.apply'),
      deleteTitle: t('vp.deleteTitle'),
      formatDate: (iso) => formatDateTime(iso),
    });
    this.dom.viewpointList.innerHTML = markup
      ?? `<div class="list-empty">${escapeHtml(t('empty.noViewpoints'))}</div>`;
  }

  private createIssueFromCurrentContext(): void {
    const title = this.dom.issueTitle.value.trim();
    if (!title) {
      this.setStatus(t('status.issueTitleRequired'));
      return;
    }

    const selectedCount = countMapItems(this.selectedItems);
    if (selectedCount === 0 && !this.pendingIssuePoint && !this.lastHitPoint) {
      this.setStatus(t('status.issueNeedsContext'));
      return;
    }

    const firstSelection = this.getFirstSelection();
    const point = this.pendingIssuePoint ?? this.lastHitPoint;
    const issuePoint = point ? { x: point.x, y: point.y, z: point.z } : null;

    // F9: capture the whole multi-model selection; modelId/localIds keep the
    // first model for backwards compatibility with older exports.
    const elementsByModel: Record<string, number[]> = {};
    for (const [modelId, ids] of Object.entries(this.selectedItems)) {
      if (ids.size > 0) elementsByModel[modelId] = [...ids];
    }

    const issue: IssueRecord = {
      id: uniqueId(),
      title,
      description: this.dom.issueDescription.value.trim(),
      priority: this.dom.issuePriority.value as IssueRecord['priority'],
      status: this.dom.issueStatus.value as IssueRecord['status'],
      assignee: this.dom.issueAssignee.value.trim(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      modelId: firstSelection?.modelId ?? null,
      localIds: firstSelection ? [...this.selectedItems[firstSelection.modelId]] : [],
      elementsByModel: Object.keys(elementsByModel).length > 0 ? elementsByModel : undefined,
      point: issuePoint,
      comments: [],
    };

    this.issues.unshift(issue);
    this.activeIssueId = issue.id;
    this.createIssueMarker(issue);
    this.updateIssuesList();
    this.updateIssueComments();

    this.dom.issueTitle.value = '';
    this.dom.issueDescription.value = '';
    this.dom.issueAssignee.value = '';

    this.pendingIssuePoint = null;
    this.issuePinMode = false;
    this.dom.viewerHint.hidden = true;
    this.dom.btnIssuePinMode.classList.remove('active');

    this.persistLocalState();
    this.setStatus(t('status.issueCreated'));
  }

  /**
   * F9: a pin is only resolvable while at least one referenced model is
   * loaded (or, for point-only pins, while any model gives it context) —
   * orphan pins floating in an empty viewer are hidden, not deleted.
   */
  private isIssueMarkerResolvable(issue: IssueRecord): boolean {
    const referenced = issue.elementsByModel
      ? Object.keys(issue.elementsByModel)
      : (issue.modelId ? [issue.modelId] : []);
    if (referenced.length > 0) return referenced.some((modelId) => this.federatedModels.has(modelId));
    return this.federatedModels.size > 0;
  }

  private createIssueMarker(issue: IssueRecord): void {
    if (!issue.point) return;

    if (issue.markerId) {
      this.markerManager.delete(issue.markerId);
      issue.markerId = undefined;
    }

    if (!this.isIssueMarkerResolvable(issue)) return;

    const markerElement = document.createElement('button');
    markerElement.type = 'button';
    markerElement.className = `issue-marker issue-${issue.status.toLowerCase().replace(/\s+/g, '-')}`;
    markerElement.textContent = issue.priority[0];
    markerElement.title = t('label.issueMarker', { title: issue.title, status: this.localizeIssueStatus(issue.status) });

    markerElement.addEventListener('click', (event) => {
      event.stopPropagation();
      this.selectIssue(issue.id, true);
    });

    const markerId = this.markerManager.create(
      this.world,
      markerElement,
      new THREE.Vector3(issue.point.x, issue.point.y, issue.point.z),
      true,
    );
    if (markerId) issue.markerId = markerId;
  }

  private refreshIssueMarkers(): void {
    for (const issue of this.issues) {
      if (issue.markerId) {
        this.markerManager.delete(issue.markerId);
        issue.markerId = undefined;
      }
      this.createIssueMarker(issue);
    }
    this.requestRender();
  }

  private updateIssuesList(): void {
    // A11: row + action clicks handled by delegation bound once in
    // bindIssueListEvents() — no per-item listeners to leak on re-render.
    const markup = buildIssueListMarkup(this.issues, this.activeIssueId, {
      linked: (count, models) => t('issue.linked', { count, models }),
      noLink: t('issue.noLink'),
      localizePriority: (priority) => this.localizeIssuePriority(priority),
      localizeStatus: (status) => this.localizeIssueStatus(status),
      deleteTitle: t('issue.deleteTitle'),
    });
    this.dom.issuesList.innerHTML = markup ?? `<div class="list-empty">${escapeHtml(t('empty.noIssues'))}</div>`;
  }

  private selectIssue(issueId: string, focusView: boolean): void {
    const issue = this.issues.find((entry) => entry.id === issueId);
    if (!issue) return;

    this.activeIssueId = issueId;
    this.updateIssuesList();
    this.updateIssueComments();

    this.dom.issueTitle.value = issue.title;
    this.dom.issueDescription.value = issue.description;
    this.dom.issuePriority.value = issue.priority;
    this.dom.issueStatus.value = issue.status;
    this.dom.issueAssignee.value = issue.assignee;

    // F9: re-select the issue's elements across every loaded model.
    const elements = issue.elementsByModel
      ?? (issue.modelId && issue.localIds.length > 0 ? { [issue.modelId]: issue.localIds } : {});
    const loadedSelection: OBC.ModelIdMap = {};
    for (const [modelId, ids] of Object.entries(elements)) {
      if (ids.length > 0 && this.isLoadedModelId(modelId)) loadedSelection[modelId] = new Set(ids);
    }
    if (Object.keys(loadedSelection).length > 0) {
      clearMap(this.selectedItems);
      Object.assign(this.selectedItems, loadedSelection);
      this.fireAndForget(this.refreshSelectionVisuals(), 'Issue selection');
    }

    if (focusView && issue.point) {
      const point = new THREE.Vector3(issue.point.x, issue.point.y, issue.point.z);
      const offset = new THREE.Vector3(7, 5, 7);
      void this.world.camera.controls.setLookAt(
        point.x + offset.x,
        point.y + offset.y,
        point.z + offset.z,
        point.x,
        point.y,
        point.z,
        true,
      );
    }

    this.dom.btnDeleteIssue.hidden = false;
    this.dom.issueCommentsGroup.hidden = false;
    this.activateTab('issues');
    this.setStatus(t('status.selectedIssue', { title: issue.title }));
  }

  private async deleteSelectedIssue(): Promise<void> {
    if (!this.activeIssueId) {
      this.setStatus(t('status.selectIssueFirst'));
      return;
    }

    const confirmed = await this.confirm(t('confirm.deleteIssue'), t('confirm.delete'), t('confirm.cancel'));
    if (!confirmed) return;

    const issue = this.issues.find((entry) => entry.id === this.activeIssueId);
    if (issue?.markerId) this.markerManager.delete(issue.markerId);

    this.issues = this.issues.filter((entry) => entry.id !== this.activeIssueId);
    this.activeIssueId = this.issues[0]?.id ?? null;

    this.updateIssuesList();
    this.updateIssueComments();
    this.persistLocalState();
    this.showToast(t('status.issueDeleted'), 'success');
    this.setStatus(t('status.issueDeleted'));
  }

  private addCommentToActiveIssue(): void {
    if (!this.activeIssueId) {
      this.setStatus(t('status.selectIssueFirst'));
      return;
    }

    const text = this.dom.issueCommentInput.value.trim();
    if (!text) {
      this.setStatus(t('status.commentEmpty'));
      return;
    }

    const issue = this.issues.find((entry) => entry.id === this.activeIssueId);
    if (!issue) return;

    issue.comments.push({
      id: uniqueId(),
      text,
      author: 'Local User',
      createdAt: new Date().toISOString(),
    });
    issue.updatedAt = new Date().toISOString();

    this.dom.issueCommentInput.value = '';
    this.updateIssueComments();
    this.persistLocalState();
    this.setStatus(t('status.commentAdded'));
  }

  private updateIssueComments(): void {
    const issue = this.issues.find((entry) => entry.id === this.activeIssueId);
    const markup = buildIssueCommentsMarkup(issue, t('empty.noComments'), (iso) => formatDateTime(iso));
    this.dom.issueComments.innerHTML = markup
      ?? `<div class="comment-item">${escapeHtml(t('empty.selectIssueForComments'))}</div>`;
  }

  /**
   * F2: the WebGL drawing buffer is not preserved, so toDataURL() on a stale
   * canvas yields a blank transparent PNG. Rendering immediately before the
   * read guarantees a fresh frame (PostproductionRenderer.update() runs the
   * composer when postproduction is enabled).
   */
  private captureCanvas(type = 'image/png', quality?: number): string | null {
    const renderer = this.world?.renderer;
    if (!renderer) return null;
    // P6 (W5.3): under MANUAL-mode rendering, update() only paints when
    // needsUpdate is armed — force a fresh frame before the readback (F2).
    this.requestRender();
    renderer.update();
    return renderer.three.domElement.toDataURL(type, quality);
  }

  /** Downscaled JPEG thumbnail of the current view (≤320px, F2). */
  private captureViewpointThumbnail(): string | undefined {
    const renderer = this.world?.renderer;
    if (!renderer) return undefined;
    // P6 (W5.3): force a fresh frame under MANUAL-mode rendering (F2).
    this.requestRender();
    renderer.update();
    const source = renderer.three.domElement;
    const largest = Math.max(source.width, source.height);
    if (largest === 0) return undefined;
    const scale = Math.min(1, VIEWPOINT_THUMBNAIL_MAX_DIM / largest);
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(source.width * scale));
    canvas.height = Math.max(1, Math.round(source.height * scale));
    const context = canvas.getContext('2d');
    if (!context) return undefined;
    context.drawImage(source, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/jpeg', VIEWPOINT_THUMBNAIL_JPEG_QUALITY);
  }

  private exportScreenshot(): void {
    const dataUrl = this.captureCanvas('image/png');
    if (!dataUrl) return;
    fetch(dataUrl)
      .then((response) => response.blob())
      .then((blob) => {
        const name = `bim-view-${new Date().toISOString().replace(/[:.]/g, '-')}.png`;
        downloadBlob(name, blob);
        this.setStatus(t('status.screenshotExported'));
      })
      .catch((error: unknown) => {
        this.showToast(t('status.screenshotFailed', { error: serializeError(error) }), 'error');
        this.setStatus(t('status.screenshotFailed', { error: serializeError(error) }));
      });
  }

  /**
   * W4.5: exports the current visibility/isolation state to a binary GLB for
   * PowerPoint Insert → 3D Models. `onlyVisible` in the exporter honours
   * hide/isolate; section-clipped geometry is a render-time effect and stays in
   * the mesh, so we warn when clipping is active.
   */
  private async exportGlb(): Promise<void> {
    if (this.federatedModels.size === 0) {
      this.showToast(t('share.glbNoModel'), 'error');
      this.setStatus(t('share.glbNoModel'));
      return;
    }
    this.setStatus(t('share.glbExporting'));
    try {
      const buffer = await exportModelsToGlb(await this.collectExportModels());
      const blob = new Blob([buffer], { type: 'model/gltf-binary' });
      const name = `bim-model-${new Date().toISOString().replace(/[:.]/g, '-')}.glb`;
      downloadBlob(name, blob);
      if (hasActiveClipping(this.clipper.enabled, this.clipper.list.size)) {
        this.showToast(t('share.glbClippingWarning'), 'info', 6000);
      }
      this.setStatus(t('share.glbExported'));
    } catch (error) {
      this.showToast(t('share.glbFailed', { error: serializeError(error) }), 'error');
      this.setStatus(t('share.glbFailed', { error: serializeError(error) }));
    }
  }

  /**
   * Gathers each visible model + the local ids to export (all geometry ids minus
   * hidden), so GLB export honours hide/isolate. Hidden whole-models are skipped.
   */
  private async collectExportModels(): Promise<{ model: FragmentsModelLike; visibleIds: number[] }[]> {
    const hiddenMap = await this.hider.getVisibilityMap(false);
    const result: { model: FragmentsModelLike; visibleIds: number[] }[] = [];
    for (const [modelId, record] of this.federatedModels) {
      if (!record.visible) continue;
      const model = this.getFragmentsModel(modelId);
      if (!model) continue;
      const allIds = await model.getItemsIdsWithGeometry();
      const hiddenList = hiddenMap[modelId] ?? [];
      const hidden = new Set<number>(hiddenList);
      const visibleIds = hidden.size > 0 ? allIds.filter((id) => !hidden.has(id)) : allIds;
      if (visibleIds.length > 0) result.push({ model, visibleIds });
    }
    return result;
  }

  private exportViewerState(): void {
    const payload = this.getPersistedState();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const name = `viewer-state-${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
    downloadBlob(name, blob);
    this.setStatus(t('status.stateExported'));
  }

  // ── W4.4/W4.2: share & host ──────────────────────────────────────────────

  /** Instantiates the share-dialog controller with app-backed callbacks. */
  private initShareDialog(): void {
    this.shareController = new ShareDialogController(
      {
        dialog: this.dom.shareDialog,
        close: this.dom.shareClose,
        tabLink: this.dom.shareTabLink,
        tabPp: this.dom.shareTabPp,
        panelLink: this.dom.sharePanelLink,
        panelPp: this.dom.sharePanelPp,
        copyLink: this.dom.shareCopyLink,
        publish: this.dom.sharePublish,
        published: this.dom.sharePublished,
        embedUrl: this.dom.shareEmbedUrl,
        copyEmbed: this.dom.shareCopyEmbed,
        iframe: this.dom.shareIframe,
        copyIframe: this.dom.shareCopyIframe,
        qr: this.dom.shareQr,
        expiry: this.dom.shareExpiry,
        delete: this.dom.shareDelete,
        hostIntro: this.dom.shareHostIntro,
        ppGlb: this.dom.sharePpGlb,
      },
      {
        getFragForShare: () => this.getFragForShare(),
        publish: async (bytes, fileName) => {
          const result = await this.uploadClient.publish(bytes, fileName);
          // Record the hosted URL against the first model so the "Copy link to
          // view" deep link works after publishing a local file.
          const firstId = [...this.federatedModels.keys()][0];
          if (firstId) this.hostedModelUrls.set(firstId, result.fragUrl);
          return result;
        },
        deleteUpload: (id, token) => this.uploadClient.remove(id, token),
        buildCopyLink: () => this.buildCopyLink(),
        exportGlb: () => this.exportGlb(),
        toast: (message, type) => this.showToast(message, type),
        status: (message) => this.setStatus(message),
        confirm: (message) => this.confirm(message),
      },
    );
    this.shareController.init();
  }

  /** The first loaded model's `.frag` bytes + name for the share/host flow. */
  private async getFragForShare(): Promise<{ bytes: Uint8Array; fileName: string } | null> {
    const firstId = [...this.federatedModels.keys()][0];
    if (!firstId) return null;
    const model = this.getFragmentsModel(firstId);
    if (!model) return null;
    const buffer = await model.getBuffer(false);
    const record = this.federatedModels.get(firstId);
    const fileName = record?.fileName ?? `${firstId}.frag`;
    return { bytes: new Uint8Array(buffer), fileName };
  }

  /**
   * Builds a "Copy link to view" deep link for the current model + view, or null
   * when no hosted URL is known (a purely-local file must be published first).
   */
  private buildCopyLink(): string | null {
    const firstId = [...this.federatedModels.keys()][0];
    const hosted = firstId ? this.hostedModelUrls.get(firstId) : undefined;
    if (!hosted) return null;
    return buildShareUrl(`${window.location.origin}/`, {
      modelUrl: hosted,
      viewpoint: this.currentUrlViewpoint(),
    });
  }

  /** Snapshots the current camera/section/toggles as a URL viewpoint (deep links). */
  private currentUrlViewpoint(): UrlViewpointState {
    const position = this.world.camera.three.position.clone();
    const target = new THREE.Vector3();
    this.world.camera.controls.getTarget(target);
    return {
      camera: {
        position: { x: position.x, y: position.y, z: position.z },
        target: { x: target.x, y: target.y, z: target.z },
        projection: this.world.camera.projection.current,
        mode: this.navigationMode,
      },
      clippingPlanes: [...this.clipper.list.values()].map((plane) => ({
        normal: { x: plane.normal.x, y: plane.normal.y, z: plane.normal.z },
        origin: { x: plane.origin.x, y: plane.origin.y, z: plane.origin.z },
      })),
      visualStyle: this.visualStyle,
      xray: this.xrayEnabled,
      edges: this.edgesEnabled,
    };
  }

  /** W4.2: loads a hosted `.frag` (or `.ifc`) named in ?m= at boot, then applies ?vp=. */
  private async loadFromUrlParams(): Promise<void> {
    const urlState = decodeUrlState(window.location.search);
    if (!urlState.modelUrl) return;
    // S8: same allowlist as the embed — the full app must not auto-fetch an
    // arbitrary attacker `?m=` URL at boot either.
    const envHosts = (import.meta.env as Record<string, string | undefined>).VITE_ALLOWED_MODEL_HOSTS;
    if (!isAllowedModelUrl(urlState.modelUrl, {
      selfOrigin: window.location.origin,
      allowedOrigins: typeof envHosts === 'string' ? envHosts.split(',').map((s) => s.trim()).filter(Boolean) : [],
      allowVercelBlob: true,
    })) {
      this.showToast(t('embed.errorBlockedUrl'), 'error');
      return;
    }
    try {
      const response = await fetch(urlState.modelUrl, { mode: 'cors' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const bytes = new Uint8Array(await response.arrayBuffer());
      const isIfc = /\.ifc(\?|#|$)/i.test(urlState.modelUrl) || isProbablyIfc(bytes);
      if (isIfc) {
        const file = new File([bytes], urlState.modelUrl.split('/').pop() ?? 'model.ifc', { type: 'application/octet-stream' });
        await this.loadIfcFile(file);
      } else {
        const modelId = this.modelRegistry.allocateModelId(
          urlState.modelUrl.split('/').pop() ?? 'model.frag',
          [...(this.fragments?.list?.keys() ?? []), ...this.federatedModels.keys()],
        );
        this.modelRegistry.beginLoad(modelId, { fileName: modelId, sizeBytes: bytes.byteLength });
        await this.fragments.core.load(bytes, { modelId });
        this.hostedModelUrls.set(modelId, urlState.modelUrl);
      }
      // Remember the source URL so "Copy link to view" works for the first model.
      const firstId = [...this.federatedModels.keys()][0];
      if (firstId && !this.hostedModelUrls.has(firstId)) this.hostedModelUrls.set(firstId, urlState.modelUrl);
    } catch (error) {
      this.showToast(t('status.loadFailed', { message: serializeError(error) }), 'error');
    }
  }

  private async importViewerState(file: File): Promise<void> {
    try {
      const text = await file.text();
      // A7: one validated apply path shared with restoreLocalState — a
      // minimal `{"version":1}` import is crash-free.
      const state = normalizePersistedState(JSON.parse(text));
      if (!state) throw new Error('Unsupported or invalid viewer state file');
      // File import restores view state only — the referenced model `.frag`
      // bytes live in the exporting browser's IndexedDB, not this one (C8).
      await this.applyPersistedState(state, false);
      this.persistLocalState();
      this.setStatus(t('status.stateImported'));
    } catch (error) {
      this.showToast(t('status.importFailed', { error: serializeError(error) }), 'error');
      this.setStatus(t('status.importFailed', { error: serializeError(error) }));
    }
  }

  /**
   * Projects the current viewer state into the persisted v2 schema via the pure
   * serializer in core/persistence.ts (the W3.5-deferred persistence-serializer
   * extraction, folded into W5.2). Gathers the C8 full-session fields (loaded
   * models + per-model modifications, camera, section, selection) here.
   */
  private getPersistedState(): PersistedViewerState {
    return buildPersistedState({
      selectionMode: this.selectionMode,
      navigationMode: this.navigationMode,
      visualStyle: this.visualStyle,
      xray: this.xrayEnabled,
      edges: this.edgesEnabled,
      gridVisible: this.gridVisible,
      backgroundColor: this.backgroundColor,
      backgroundByTheme: { ...this.backgroundByTheme },
      theme: this.themeMode,
      language: getLanguage(),
      viewpoints: this.viewpoints,
      issues: this.issues,
      models: this.collectPersistedModels(),
      camera: this.currentPersistedCamera(),
      sectionPlanes: this.currentSectionPlanes(),
      selection: this.currentSelectionRecord(),
      activeTab: this.activeTab,
      maxSnapshotChars: MAX_PERSISTED_SNAPSHOT_CHARS,
    });
  }

  /** C8: the loaded models + per-model modifications (only cacheable ones). */
  private collectPersistedModels(): PersistedModelRecord[] {
    const records: PersistedModelRecord[] = [];
    for (const record of this.federatedModels.values()) {
      // A model with no cached fragKey can't be restored — skip it.
      if (!record.fragKey) continue;
      records.push({
        modelId: record.modelId,
        fileName: record.fileName,
        sizeBytes: record.sizeBytes,
        fragKey: record.fragKey,
        offsetPosition: { ...record.offsetPosition },
        offsetRotation: { ...record.offsetRotation },
        opacity: record.opacity,
        visible: record.visible,
        hiddenIds: this.hiddenIdsByModel.get(record.modelId) ?? [],
      });
    }
    return records;
  }

  /** C8: current camera pose (skipped before the engine exists). */
  private currentPersistedCamera(): PersistedCamera | undefined {
    if (!this.world) return undefined;
    const position = this.world.camera.three.position;
    const target = new THREE.Vector3();
    this.world.camera.controls.getTarget(target);
    return {
      position: { x: position.x, y: position.y, z: position.z },
      target: { x: target.x, y: target.y, z: target.z },
      projection: this.world.camera.projection.current,
    };
  }

  /** C8: active section planes (normal + coplanar origin). */
  private currentSectionPlanes(): Array<{ normal: Vector3Record; origin: Vector3Record }> {
    if (!this.clipper) return [];
    return [...this.clipper.list.values()].map((plane) => ({
      normal: { x: plane.normal.x, y: plane.normal.y, z: plane.normal.z },
      origin: { x: plane.origin.x, y: plane.origin.y, z: plane.origin.z },
    }));
  }

  /** C8: the current selection as a plain per-model id record. */
  private currentSelectionRecord(): Record<string, number[]> {
    const selection: Record<string, number[]> = {};
    for (const [modelId, ids] of Object.entries(this.selectedItems)) {
      const list = [...ids];
      if (list.length > 0) selection[modelId] = list;
    }
    return selection;
  }

  private persistLocalState(): void {
    // C8: don't churn localStorage while a restore is re-applying state — the
    // final state is persisted once when the restore completes.
    if (this.restoringSession) return;
    const payload = this.getPersistedState();
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // Quota exceeded — retry once without snapshots (F2 size guard).
      try {
        const stripped = {
          ...payload,
          viewpoints: payload.viewpoints.map((viewpoint) => ({ ...viewpoint, snapshot: undefined })),
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(stripped));
        this.setStatus(t('status.savedWithoutThumbnails'));
      } catch (retryError) {
        this.showToast(t('status.persistFailed', { error: serializeError(retryError) }), 'error');
        this.setStatus(t('status.persistFailed', { error: serializeError(retryError) }));
      }
    }
  }

  private async restoreLocalState(): Promise<void> {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const state = normalizePersistedState(JSON.parse(raw));
      if (!state) return;
      await this.applyPersistedState(state);
    } catch (error) {
      this.showToast(t('status.restoreFailed', { error: serializeError(error) }), 'error');
      this.setStatus(t('status.restoreFailed', { error: serializeError(error) }));
    }
  }

  /**
   * C8 (W5.2): explicit "Save session" — persists the full session (state is
   * already written on every change; this is the obvious, user-facing affordance)
   * and confirms with a toast. Model `.frag` bytes are already cached in IDB as
   * they load, so this is a fast metadata write.
   */
  private saveSession(): void {
    try {
      this.persistLocalState();
      const count = this.collectPersistedModels().length;
      this.showToast(t('status.sessionSaved', { count }), 'success');
      this.setStatus(t('status.sessionSaved', { count }));
    } catch (error) {
      this.showToast(t('status.sessionSaveFailed', { error: serializeError(error) }), 'error');
    }
  }

  /**
   * C8 (W5.2): explicit "Restore session" — re-applies the saved session from
   * localStorage (models reload from the IDB `.frag` cache, no re-conversion).
   * Same path as the boot auto-restore.
   */
  private async restoreSavedSession(): Promise<void> {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      this.showToast(t('status.noSessionToRestore'), 'info');
      return;
    }
    await this.restoreLocalState();
  }

  /**
   * A7: the single apply path for validated persisted state — used by both
   * restoreLocalState (localStorage) and importViewerState (file import).
   * Expects a state produced by normalizePersistedState.
   */
  private async applyPersistedState(state: PersistedViewerState, restoreModels = true): Promise<void> {
    // C7: restore the UI language first so any status text produced below is
    // already localized. setLanguage no-ops when unchanged, and re-hydrates +
    // re-renders (harmless during boot). Only applied when the state carries one.
    if (state.language) setLanguage(state.language);
    this.selectionMode = state.selectionMode;
    this.navigationMode = state.navigationMode;
    this.visualStyle = this.parseVisualStyle(state.visualStyle ?? 'color-pen-shadows');
    const restoredXray = state.xray;
    const restoredEdges = state.edges;
    this.gridVisible = state.gridVisible ?? false;
    this.themeMode = state.theme ?? 'dark';
    if (state.backgroundByTheme) {
      this.backgroundByTheme.dark = this.normalizeHexColor(state.backgroundByTheme.dark, DEFAULT_BACKGROUND_COLOR);
      this.backgroundByTheme.light = this.normalizeHexColor(state.backgroundByTheme.light, DEFAULT_LIGHT_BACKGROUND_COLOR);
    }
    this.backgroundColor = this.normalizeHexColor(
      state.backgroundColor ?? this.backgroundByTheme[this.themeMode],
      this.backgroundByTheme[this.themeMode],
    );
    this.viewpoints = state.viewpoints;
    this.issues = state.issues.map((issue) => ({ ...issue }));

    document.documentElement.setAttribute('data-theme', this.themeMode);

    this.applySelectionMode(this.selectionMode);
    this.applyNavigationMode(this.navigationMode);

    await this.setVisualStyle(this.visualStyle, false, false);
    this.xrayEnabled = restoredXray;
    this.edgesEnabled = restoredEdges;
    this.setRailPressed(this.dom.btnTransparency, this.xrayEnabled);
    this.setRailPressed(this.dom.btnWireframe, this.edgesEnabled);
    this.setGridVisible(this.gridVisible, false);
    this.setBackgroundColor(this.backgroundColor, false);
    this.syncVisualSettingsUi();
    this.applyXRay();
    this.applyEdges();

    this.updateViewpointList();
    this.updateIssuesList();
    this.updateIssueComments();
    this.refreshIssueMarkers();

    if (state.activeTab) this.activateTab(state.activeTab);

    // C8 (W5.2): restore the full session — reload the cached models (no
    // re-conversion), re-apply per-model modifications, then camera/section/
    // selection. Only on boot / manual restore (restoreModels), not on file
    // import: an imported state references `.frag` bytes cached in the ORIGINATING
    // browser's IndexedDB, which this browser does not have, so a file import
    // restores view state (viewpoints/issues/theme/…) only — as it did pre-C8.
    if (restoreModels && this.fragments && state.models.length > 0) {
      await this.restoreSession(state);
    }
  }

  /**
   * C8 (W5.2): reloads each persisted model from the IndexedDB `.frag` cache via
   * the fragments loader (NO re-conversion), then re-applies every stored
   * per-model modification and the view state (camera, section, selection).
   * Models whose cached bytes are gone are skipped with a toast.
   */
  private async restoreSession(state: PersistedViewerState): Promise<void> {
    this.restoringSession = true;
    this.suppressAutoFit = true;
    let restored = 0;
    let missing = 0;
    try {
      for (const persisted of state.models) {
        // Already loaded (a ?m= boot load, or a manual re-restore): don't reload
        // the fragments, just re-apply the persisted per-model modifications.
        if (this.federatedModels.has(persisted.modelId)) {
          this.applyPersistedModelModifications(persisted);
          restored += 1;
          continue;
        }
        const entry = await this.fragCache.get(persisted.fragKey).catch(() => null);
        if (!entry) {
          missing += 1;
          continue;
        }
        // Reload from cached fragments bytes — the loader re-adds the model and
        // fires onModelLoaded → registerModel → onModelAdded (which reuses the
        // known fragKey via restoreFragKeys, so no re-cache).
        this.restoreFragKeys.set(persisted.modelId, persisted.fragKey);
        this.modelRegistry.beginLoad(persisted.modelId, {
          fileName: persisted.fileName,
          sizeBytes: persisted.sizeBytes,
        });
        try {
          const loaded = await this.fragments.core.load(new Uint8Array(entry.bytes), {
            modelId: persisted.modelId,
          });
          // Register explicitly (mirrors loadIfcFile) rather than racing the
          // onModelLoaded event, so onModelAdded has completed before we modify.
          await this.registerModel(persisted.modelId, loaded);
          this.applyPersistedModelModifications(persisted);
          restored += 1;
        } catch (error) {
          this.restoreFragKeys.delete(persisted.modelId);
          this.modelRegistry.failLoad(persisted.modelId);
          console.error(`Failed to restore model ${persisted.modelId}`, error);
          missing += 1;
        }
      }

      await this.applyRestoredViewState(state);
    } finally {
      this.restoringSession = false;
      this.suppressAutoFit = false;
    }

    this.dom.emptyState.hidden = this.federatedModels.size > 0;
    if (restored > 0) {
      this.setStatus(t('status.sessionRestored', { count: restored }));
    }
    if (missing > 0) {
      this.showToast(t('status.sessionModelsMissing', { count: missing }), 'warning');
    }
  }

  /** C8: re-applies one persisted model's transform / opacity / visibility / hide. */
  private applyPersistedModelModifications(persisted: PersistedModelRecord): void {
    const record = this.federatedModels.get(persisted.modelId);
    if (!record) return;
    record.offsetPosition = { ...persisted.offsetPosition };
    record.offsetRotation = { ...persisted.offsetRotation };
    record.opacity = Math.min(1, Math.max(0, persisted.opacity));
    record.visible = persisted.visible;
    record.object.visible = persisted.visible;
    this.applyModelTransform(record);
    if (persisted.hiddenIds.length > 0) {
      this.hiddenIdsByModel.set(persisted.modelId, [...persisted.hiddenIds]);
      this.fireAndForget(
        this.hider.set(false, { [persisted.modelId]: new Set(persisted.hiddenIds) }),
        'Restore hidden ids',
      );
    }
  }

  /** C8: re-applies camera pose, section planes and selection after models load. */
  private async applyRestoredViewState(state: PersistedViewerState): Promise<void> {
    // Section planes.
    if (state.sectionPlanes.length > 0) {
      this.clipper.enabled = true;
      for (const plane of state.sectionPlanes) {
        this.createClipPlane(
          new THREE.Vector3(plane.normal.x, plane.normal.y, plane.normal.z),
          new THREE.Vector3(plane.origin.x, plane.origin.y, plane.origin.z),
        );
      }
    }

    // Selection.
    if (Object.keys(state.selection).length > 0) {
      clearMap(this.selectedItems);
      for (const [modelId, ids] of Object.entries(state.selection)) {
        if (this.federatedModels.has(modelId) && ids.length > 0) {
          this.selectedItems[modelId] = new Set(ids);
        }
      }
      await this.refreshSelectionVisuals();
      this.updateCounters();
    }

    // Opacity/edges reflect the restored per-model state.
    this.applyXRay();
    this.applyEdges();
    await this.updateVisibilityCount();

    // Camera pose — restored last so it isn't overwritten by a fit.
    if (state.camera && this.world) {
      if (this.world.camera.projection.current !== state.camera.projection) {
        await this.world.camera.projection.set(state.camera.projection);
      }
      await this.world.camera.controls.setLookAt(
        state.camera.position.x,
        state.camera.position.y,
        state.camera.position.z,
        state.camera.target.x,
        state.camera.target.y,
        state.camera.target.z,
        false,
      );
    } else if (this.federatedModels.size > 0) {
      this.fitToModel();
    }

    this.renderModelBrowser();
    this.renderFederatedTree();
    this.renderClassFilters();
    this.renderLevelFilters();
    this.updateElementCounter();
    await this.updateFragments(true);
  }

  /** U7: real tab activation — aria-selected + roving tabindex + panel title. */
  private activateTab(tab: string): void {
    this.activeTab = tab;
    this.dom.tabStripButtons.forEach((button) => {
      const active = button.dataset.tab === tab;
      button.classList.toggle('is-active', active);
      button.setAttribute('aria-selected', String(active));
      button.tabIndex = active ? 0 : -1;
    });
    this.dom.tabPanels.forEach((panel) => {
      const active = panel.id === `panel-${tab}`;
      panel.classList.toggle('is-active', active);
      panel.hidden = !active;
    });
    this.updateActiveTabTitle();
  }

  /** Sets the side-panel title from the active tab (re-run on language change). */
  private updateActiveTabTitle(): void {
    const titleKeys = {
      explorer: 'panel.explorer',
      models: 'panel.models',
      properties: 'panel.properties',
      viewpoints: 'panel.viewpoints',
      issues: 'panel.issues',
      help: 'panel.help',
    } as const;
    const key = titleKeys[this.activeTab as keyof typeof titleKeys] ?? 'panel.explorer';
    this.dom.panelTitle.textContent = t(key);
  }

  private async updateVisibilityCount(): Promise<void> {
    const visibleMap = await this.hider.getVisibilityMap(true);
    let visibleCount = 0;
    for (const [modelId, ids] of Object.entries(visibleMap)) {
      const model = this.federatedModels.get(modelId);
      if (model && !model.visible) continue;
      visibleCount += ids.length;
    }
    this.dom.visibleCount.textContent = t('label.visible', { count: visibleCount });
    // C8 (W5.2): refresh the per-model hidden-id snapshot so the (synchronous)
    // session serializer persists the current hide/isolate state.
    await this.refreshHiddenIdsSnapshot();
  }

  /** C8: caches hidden element ids per model from the hider (for persistence). */
  private async refreshHiddenIdsSnapshot(): Promise<void> {
    try {
      const hiddenMap = await this.hider.getVisibilityMap(false);
      this.hiddenIdsByModel.clear();
      for (const [modelId, ids] of Object.entries(hiddenMap)) {
        if (ids.length > 0) this.hiddenIdsByModel.set(modelId, [...ids]);
      }
    } catch (error) {
      console.debug('Hidden-id snapshot refresh failed', error);
    }
  }

  private updateCounters(): void {
    this.dom.selectionCount.textContent = t('label.selected', { count: countMapItems(this.selectedItems) });
  }

  /**
   * The keyboard action bindings the extracted router dispatches to (W5.3). Each
   * mirrors the original inline onKeyDown branch exactly — the router only owns
   * the form-target guard + key→action mapping (input/keyboard-router.ts).
   */
  private readonly keyboardActions: KeyboardActions = {
    cancel: () => {
      this.setMeasureMode('none');
      this.issuePinMode = false;
      this.dom.viewerHint.hidden = true;
      this.setRailPressed(this.dom.btnIssuePinMode, false);
      if (this.activeGizmoModelId) this.detachModelGizmo();
      if (this.dom.root.classList.contains('sheet-open')) this.closeSheet();
      else if (this.isSmallScreen() && this.dom.root.classList.contains('panel-open')) this.closePanel();
      this.setStatus(t('status.toolCanceled'));
    },
    fitToModel: () => this.fitToModel(),
    setNavigationMode: (mode) => this.applyNavigationMode(mode),
    toggleSelectionMode: () => this.applySelectionMode(this.selectionMode === 'single' ? 'multi' : 'single'),
    toggleMeasure: (mode) => this.setMeasureMode(this.measureMode === mode ? 'none' : mode),
    toggleGrid: () => this.dom.btnToggleGrid.click(),
    toggleXray: () => this.dom.btnTransparency.click(),
    edgesOrGizmoRotate: () => {
      if (this.activeGizmoModelId) {
        this.transformControls?.setMode('rotate');
        this.setStatus(t('status.gizmoModeRotate'));
      } else {
        this.dom.btnWireframe.click();
      }
    },
    gizmoTranslate: () => {
      if (this.activeGizmoModelId) {
        this.transformControls?.setMode('translate');
        this.setStatus(t('status.gizmoModeTranslate'));
      }
    },
    gizmoReset: () => {
      if (this.activeGizmoModelId) {
        this.resetModelOffsets(this.activeGizmoModelId);
        this.setStatus(t('status.gizmoTransformReset'));
      }
    },
    toggleIssuePin: () => this.dom.btnIssuePinMode.click(),
    deleteSelectedIssue: () => this.fireAndForget(this.deleteSelectedIssue(), 'Delete issue'),
    finishAreaMeasurement: () => {
      if (this.measureMode !== 'area') return;
      this.areaMeasurement.endCreation();
      // R3 (W5-fixups): committing the area fill mutates the scene with no camera
      // motion — repaint so the finished fill shows without a camera nudge.
      this.requestRender();
    },
  };

  private onKeyDown(event: KeyboardEvent): void {
    routeKeyboardEvent(event, this.keyboardActions);
  }

  private setStatus(message: string): void {
    this.dom.statusText.textContent = message;
  }

  /**
   * U4/U11: floating toast, reachable on every viewport (region is fixed and
   * clear of the view controls). Auto-dismisses; assertive for errors.
   */
  private showToast(message: string, type: 'success' | 'error' | 'warning' | 'info' = 'info', durationMs = 4000): void {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.setAttribute('role', type === 'error' ? 'alert' : 'status');
    const iconSpan = document.createElement('span');
    const iconName: IconName = type === 'error' ? 'error_outline' : type === 'success' ? 'visibility' : 'info';
    setIcon(iconSpan, iconName, 18);
    const text = document.createElement('span');
    text.textContent = message;
    toast.append(iconSpan, text);
    this.dom.toastRegion.append(toast);
    setTimeout(() => {
      toast.addEventListener('transitionend', () => toast.remove(), { once: true });
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 400);
    }, durationMs);
  }

  /**
   * U8: modal confirm via native <dialog>.showModal() — the browser provides
   * the focus trap and Escape handling; we focus Cancel (not the destructive
   * button) and restore focus to the invoker on close.
   */
  private confirm(message: string, confirmLabel = 'Confirm', cancelLabel = 'Cancel'): Promise<boolean> {
    return new Promise((resolve) => {
      const dialog = this.dom.confirmDialog;
      const invoker = document.activeElement as HTMLElement | null;
      this.dom.confirmMessage.textContent = message;
      this.dom.confirmOk.textContent = confirmLabel;
      this.dom.confirmCancel.textContent = cancelLabel;
      const onClose = (): void => {
        dialog.removeEventListener('close', onClose);
        invoker?.focus?.();
        resolve(dialog.returnValue === 'confirm');
      };
      dialog.addEventListener('close', onClose);
      dialog.returnValue = 'cancel';
      dialog.showModal();
      // Focus the non-destructive action (U8), not the confirm button.
      this.dom.confirmCancel.focus();
    });
  }

  /**
   * Builds the frozen e2e contract (T6). Only reached in VITE_E2E builds.
   * Every member is a stable, documented accessor over viewer state/behavior;
   * e2e must not reach past this object into private fields.
   */
  private buildTestApi(): ViewerTestApi {
    const api: ViewerTestApi = {
      version: 1,
      modelCount: () => this.federatedModels.size,
      indexedModelCount: () => this.modelIndices.size,
      findItemByName: (keyword: string): TestItemRef | null => {
        const needle = keyword.toLowerCase();
        for (const [modelId, index] of this.modelIndices.entries()) {
          for (const [localId, name] of index.itemNames.entries()) {
            if (!index.allIds.has(localId) || !name) continue;
            if (name.toLowerCase().includes(needle)) return { modelId, localId };
          }
        }
        return null;
      },
      firstModelId: (): string | null => {
        for (const modelId of this.federatedModels.keys()) return modelId;
        return null;
      },
      allModelIds: (): string[] => [...this.federatedModels.keys()],
      findDisjointClassLevel: () => {
        const index = this.modelIndices.values().next().value;
        if (!index) return null;
        for (const [className, classIds] of index.classes.entries()) {
          for (const [levelName, levelIds] of index.levels.entries()) {
            let overlaps = false;
            for (const id of classIds) {
              if (levelIds.has(id)) { overlaps = true; break; }
            }
            if (!overlaps) return { className, levelName };
          }
        }
        return null;
      },
      firstModelContext: () => {
        const firstEntry = this.modelIndices.entries().next().value;
        if (!firstEntry) return null;
        const [modelId, index] = firstEntry;
        let firstNamed: { id: number; name: string } | null = null;
        for (const [id, name] of index.itemNames.entries()) {
          if (!index.allIds.has(id) || !name) continue;
          firstNamed = { id, name };
          break;
        }
        return {
          modelId,
          searchTerm: firstNamed?.name || index.classes.keys().next().value || '',
          firstItemId: firstNamed?.id ?? index.allIds.values().next().value ?? 0,
        };
      },
      selectionCount: () => countMapItems(this.selectedItems),
      isXrayEnabled: () => this.xrayEnabled,
      isEdgesEnabled: () => this.edgesEnabled,
      activeGizmoModelId: () => this.activeGizmoModelId,
      viewpointCount: () => this.viewpoints.length,
      issueCount: () => this.issues.length,
      firstViewpointSnapshot: () => this.viewpoints[0]?.snapshot ?? null,
      firstIssueLinkedCount: () => this.issues[0]?.localIds.length ?? 0,
      firstIssueModelCount: () => Object.keys(this.issues[0]?.elementsByModel ?? {}).length,
      firstIssueHasLegacyModelId: () => typeof this.issues[0]?.modelId === 'string',
      engineModelState: () => ({
        fragmentsCount: this.fragments.list.size,
        federatedCount: this.federatedModels.size,
        indexCount: this.modelIndices.size,
        objectCount: this.modelObjects.length,
      }),
      cameraState: () => {
        const camera = this.world.camera.three;
        const target = new THREE.Vector3();
        this.world.camera.controls.getTarget(target);
        return {
          position: { x: camera.position.x, y: camera.position.y, z: camera.position.z },
          target: { x: target.x, y: target.y, z: target.z },
        };
      },
      anchorDirectionForCube: (localVector) => {
        const anchor = this.getViewCubeAnchorModel();
        if (!anchor) return null;
        const basis = this.getViewCubeBasisQuaternion(new THREE.Quaternion());
        const direction = new THREE.Vector3(localVector[0], localVector[1], localVector[2])
          .normalize()
          .applyQuaternion(basis)
          .normalize();
        return { x: direction.x, y: direction.y, z: direction.z };
      },
      selectItem: (modelId: string, localId: number, zoom = false): Promise<void> =>
        this.selectSingleItem(modelId, localId, zoom),
      selectFirstItemPerModel: async (): Promise<void> => {
        clearMap(this.selectedItems);
        for (const [modelId, index] of this.modelIndices.entries()) {
          const firstId = index.allIds.values().next().value;
          if (typeof firstId === 'number') this.selectedItems[modelId] = new Set([firstId]);
        }
        await this.refreshSelectionVisuals();
      },
      setVisualStyle: (style: string): Promise<void> =>
        this.setVisualStyle(this.parseVisualStyle(style), false, false),
      clickCanvasAt: async (clientX: number, clientY: number): Promise<TestItemRef | null> => {
        // ThatOpen's raycaster expects the pick position in normalized device
        // coordinates ([-1,1], y-up) — the same shape SimpleMouse.position
        // returns — NOT CSS pixels. Convert here.
        const rect = this.dom.viewerContainer.getBoundingClientRect();
        const ndc = new THREE.Vector2(
          ((clientX - rect.left) / rect.width) * 2 - 1,
          -((clientY - rect.top) / rect.height) * 2 + 1,
        );
        return this.pickAndSelect(ndc);
      },
      exportGlbBytes: async (): Promise<{ byteLength: number; valid: boolean }> => {
        // Geometry comes from model.getItemsGeometry (CPU-side), so no render
        // frame is required — works even under headless software WebGL.
        const buffer = await exportModelsToGlb(await this.collectExportModels());
        return { byteLength: buffer.byteLength, valid: isValidGlb(buffer) };
      },
      // ---- C8 (W5.2) ----
      modelModifications: (modelId: string) => {
        const record = this.federatedModels.get(modelId);
        if (!record) return null;
        return {
          offsetPosition: { ...record.offsetPosition },
          offsetRotation: { ...record.offsetRotation },
          opacity: record.opacity,
          visible: record.visible,
          fragKey: record.fragKey ?? null,
        };
      },
      persistedModelCount: (): number => {
        try {
          const raw = localStorage.getItem(STORAGE_KEY);
          if (!raw) return 0;
          const state = normalizePersistedState(JSON.parse(raw));
          return state?.models.length ?? 0;
        } catch {
          return 0;
        }
      },
      setModelOpacity: (modelId: string, opacity: number): void => {
        this.applyModelOpacity(modelId, opacity);
      },
      setModelOffset: (modelId: string, x: number, y: number, z: number): void => {
        const record = this.federatedModels.get(modelId);
        if (!record) return;
        record.offsetPosition = { x, y, z };
        this.applyModelTransform(record);
        this.persistLocalState();
      },
      saveSession: (): void => this.saveSession(),
      renderRequestCount: (): number => this.renderRequestCount,
    };
    return Object.freeze(api);
  }

  /**
   * P6 (W5.3): switch the world's renderer to on-demand (MANUAL) rendering.
   * Called at the end of init() — after postproduction is configured and the
   * initial AUTO frames built the composer/basePass — so the `enabled` setter's
   * basePass read never throws.
   */
  private enableOnDemandRendering(): void {
    if (!this.engineHandles) return;
    this.onDemandRender = this.engineModule.enableOnDemandRendering(this.engineHandles);
    const rawRequestRender = this.onDemandRender.requestRender;
    // W5-fixups: count every render request so the e2e can assert render parity.
    this.requestRender = () => {
      this.renderRequestCount += 1;
      rawRequestRender();
    };
  }

  private startFpsMonitor(): void {
    // A15 (W5.3): count real rendered frames via the renderer's onAfterUpdate
    // (fires only on an actual paint) rather than the rAF cadence.
    const renderer = this.world?.renderer as unknown as
      | { onAfterUpdate: { add(cb: () => void): void; remove(cb: () => void): void } }
      | undefined;
    this.fpsMonitor = this.engineModule.createFpsMonitor(this.dom.perfInfo, renderer);
  }
}

const viewerApp = new ViewerApp();
void viewerApp.init();

// A5: on HMR reload, tear the previous instance down (aborts listeners, cancels
// rAF, disposes THREE/engine resources, restores console.warn) so dev reloads
// don't leak a second viewer/render loop. No-op in production builds.
if (import.meta.hot) {
  import.meta.hot.dispose(() => viewerApp.destroy());
}
