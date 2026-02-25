import * as THREE from 'three';
import * as WebIFC from 'web-ifc';
import * as OBC from '@thatopen/components';
import * as OBCF from '@thatopen/components-front';
import { TransformControls } from 'three/examples/jsm/controls/TransformControls.js';
import type { SpatialTreeItem } from '@thatopen/fragments';

type SelectionMode = 'single' | 'multi';
type MeasureMode = 'none' | 'length' | 'area';
type NavigationMode = 'Orbit' | 'Plan' | 'FirstPerson';
type VisualStyle = 'basic' | 'pen' | 'color-pen' | 'color-shadows' | 'color-pen-shadows';

interface IssueCommentRecord {
  id: string;
  text: string;
  author: string;
  createdAt: string;
}

interface IssueRecord {
  id: string;
  title: string;
  description: string;
  priority: 'Critical' | 'High' | 'Medium' | 'Low';
  status: 'Open' | 'In Progress' | 'Resolved' | 'Closed';
  assignee: string;
  createdAt: string;
  updatedAt: string;
  modelId: string | null;
  localIds: number[];
  point: { x: number; y: number; z: number } | null;
  markerId?: string;
  comments: IssueCommentRecord[];
}

interface SavedViewpoint {
  id: string;
  name: string;
  createdAt: string;
  camera: {
    position: { x: number; y: number; z: number };
    target: { x: number; y: number; z: number };
    projection: OBC.CameraProjection;
    mode: NavigationMode;
  };
  clippingPlanes: Array<{
    normal: { x: number; y: number; z: number };
    origin: { x: number; y: number; z: number };
  }>;
  hiddenItems: Record<string, number[]>;
  visualStyle?: VisualStyle;
  xray: boolean;
  edges: boolean;
  snapshot?: string;
}

interface PersistedViewerState {
  version: 1;
  selectionMode: SelectionMode;
  navigationMode: NavigationMode;
  visualStyle?: VisualStyle;
  xray: boolean;
  edges: boolean;
  gridVisible?: boolean;
  backgroundColor?: string;
  viewpoints: SavedViewpoint[];
  issues: Omit<IssueRecord, 'markerId'>[];
}

interface SearchResult {
  modelId: string;
  localId: number;
  name: string;
  type: string;
  globalId: string;
}

interface ModelIndex {
  modelId: string;
  allIds: Set<number>;
  classes: Map<string, Set<number>>;
  levels: Map<string, Set<number>>;
  itemToLevel: Map<number, string>;
  itemNames: Map<number, string>;
  spatialRoot: BrowserTreeNode | null;
}

interface BrowserTreeNode {
  category: string;
  localId: number | null;
  label: string;
  geometryCount: number;
  children: BrowserTreeNode[];
}

interface TransformVector3 {
  x: number;
  y: number;
  z: number;
}

interface FederatedModelRecord {
  modelId: string;
  fileName: string;
  sizeBytes: number;
  elementCount: number;
  visible: boolean;
  opacity: number;
  object: THREE.Object3D;
  basePosition: TransformVector3;
  baseRotation: TransformVector3;
  offsetPosition: TransformVector3;
  offsetRotation: TransformVector3;
}

const STORAGE_KEY = 'bim_for_field_viewer_state_v1';
const DEFAULT_BACKGROUND_COLOR = '#0b1220';
const MAX_PROPERTY_ROWS = 280;
const MAX_PROPERTY_DEPTH = 4;
const MAX_PROPERTY_VALUE_LENGTH = 220;
const MAX_PROPERTY_ARRAY_PREVIEW = 6;
const MAX_BROWSER_LEVELS = 120;
const MAX_BROWSER_CLASSES_PER_LEVEL = 28;
const MAX_BROWSER_ELEMENTS_PER_CLASS = 26;
const MAX_BROWSER_SPATIAL_DEPTH = 7;
const MAX_BROWSER_SPATIAL_CHILDREN = 80;

const toSetMap = (plain: Record<string, number[] | Set<number>>): OBC.ModelIdMap => {
  const result: OBC.ModelIdMap = {};
  for (const [modelId, ids] of Object.entries(plain)) {
    const set = ids instanceof Set ? ids : new Set(ids);
    if (set.size > 0) result[modelId] = set;
  }
  return result;
};

const cloneMap = (map: OBC.ModelIdMap): OBC.ModelIdMap => {
  const copy: OBC.ModelIdMap = {};
  for (const [modelId, ids] of Object.entries(map)) copy[modelId] = new Set(ids);
  return copy;
};

const clearMap = (map: OBC.ModelIdMap): void => {
  for (const key of Object.keys(map)) delete map[key];
};

const isMapEmpty = (map: OBC.ModelIdMap): boolean => {
  for (const ids of Object.values(map)) {
    if (ids.size > 0) return false;
  }
  return true;
};

const countMapItems = (map: OBC.ModelIdMap): number => {
  let count = 0;
  for (const ids of Object.values(map)) count += ids.size;
  return count;
};

const intersectMaps = (a: OBC.ModelIdMap, b: OBC.ModelIdMap): OBC.ModelIdMap => {
  const result: OBC.ModelIdMap = {};
  for (const [modelId, idsA] of Object.entries(a)) {
    const idsB = b[modelId];
    if (!idsB) continue;
    const intersection = new Set<number>();
    for (const id of idsA) {
      if (idsB.has(id)) intersection.add(id);
    }
    if (intersection.size > 0) result[modelId] = intersection;
  }
  return result;
};

const serializeError = (error: unknown): string => {
  if (error instanceof Error) return error.message;
  return String(error);
};

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const escapeHtml = (value: string): string => value
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

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

const required = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing required DOM element #${id}`);
  return element as T;
};

class ViewerApp {
  private readonly dom = {
    viewerContainer: required<HTMLDivElement>('viewer-container'),
    btnUpload: required<HTMLButtonElement>('btnUpload'),
    btnUploadEmpty: required<HTMLButtonElement>('btnUploadEmpty'),
    fileInput: required<HTMLInputElement>('fileInput'),
    btnExportScreenshot: required<HTMLButtonElement>('btnExportScreenshot'),
    btnExportState: required<HTMLButtonElement>('btnExportState'),
    btnImportState: required<HTMLButtonElement>('btnImportState'),
    importStateInput: required<HTMLInputElement>('importStateInput'),
    loadingOverlay: required<HTMLDivElement>('loadingOverlay'),
    loadingText: required<HTMLDivElement>('loadingText'),
    loadingProgress: required<HTMLDivElement>('loadingProgress'),
    emptyState: required<HTMLDivElement>('emptyState'),
    viewerHint: required<HTMLDivElement>('viewerHint'),
    modelName: required<HTMLDivElement>('modelName'),
    headerModeLabel: required<HTMLDivElement>('headerModeLabel'),
    statusText: required<HTMLSpanElement>('statusText'),
    selectionCount: required<HTMLSpanElement>('selectionCount'),
    elementCount: required<HTMLSpanElement>('elementCount'),
    visibleCount: required<HTMLSpanElement>('visibleCount'),
    perfInfo: required<HTMLSpanElement>('perfInfo'),
    btnModeOrbit: required<HTMLButtonElement>('btnModeOrbit'),
    btnModePlan: required<HTMLButtonElement>('btnModePlan'),
    btnModeFirstPerson: required<HTMLButtonElement>('btnModeFirstPerson'),
    btnFitAll: required<HTMLButtonElement>('btnFitAll'),
    btnFront: required<HTMLButtonElement>('btnFront'),
    btnTop: required<HTMLButtonElement>('btnTop'),
    btnSelectSingle: required<HTMLButtonElement>('btnSelectSingle'),
    btnSelectMulti: required<HTMLButtonElement>('btnSelectMulti'),
    btnIsolate: required<HTMLButtonElement>('btnIsolate'),
    btnHide: required<HTMLButtonElement>('btnHide'),
    btnShow: required<HTMLButtonElement>('btnShow'),
    btnResetVisibility: required<HTMLButtonElement>('btnResetVisibility'),
    btnSectionX: required<HTMLButtonElement>('btnSectionX'),
    btnSectionY: required<HTMLButtonElement>('btnSectionY'),
    btnSectionZ: required<HTMLButtonElement>('btnSectionZ'),
    btnSectionBox: required<HTMLButtonElement>('btnSectionBox'),
    btnClearSections: required<HTMLButtonElement>('btnClearSections'),
    btnMeasureLength: required<HTMLButtonElement>('btnMeasureLength'),
    btnMeasureArea: required<HTMLButtonElement>('btnMeasureArea'),
    btnClearMeasurements: required<HTMLButtonElement>('btnClearMeasurements'),
    btnTransparency: required<HTMLButtonElement>('btnTransparency'),
    btnWireframe: required<HTMLButtonElement>('btnWireframe'),
    btnIssuePinMode: required<HTMLButtonElement>('btnIssuePinMode'),
    cubeTop: required<HTMLButtonElement>('cubeTop'),
    cubeFront: required<HTMLButtonElement>('cubeFront'),
    cubeRight: required<HTMLButtonElement>('cubeRight'),
    cubeLeft: required<HTMLButtonElement>('cubeLeft'),
    cubeBack: required<HTMLButtonElement>('cubeBack'),
    cubeBottom: required<HTMLButtonElement>('cubeBottom'),
    viewCubeBody: required<HTMLDivElement>('viewCubeBody'),
    viewerDock: required<HTMLDivElement>('viewerDock'),
    modelBrowserTree: required<HTMLDivElement>('modelBrowserTree'),
    federationTree: required<HTMLDivElement>('federationTree'),
    toggleGrid: required<HTMLInputElement>('toggleGrid'),
    visualStyleSelect: required<HTMLSelectElement>('visualStyleSelect'),
    backgroundColorInput: required<HTMLInputElement>('backgroundColorInput'),
    backgroundPresetButtons: Array.from(document.querySelectorAll<HTMLButtonElement>('[data-bg-preset]')),
    tabButtons: Array.from(document.querySelectorAll<HTMLButtonElement>('.tab-btn')),
    searchInput: required<HTMLInputElement>('searchInput'),
    btnSearch: required<HTMLButtonElement>('btnSearch'),
    btnClearSearch: required<HTMLButtonElement>('btnClearSearch'),
    classFilterList: required<HTMLDivElement>('classFilterList'),
    levelFilterList: required<HTMLDivElement>('levelFilterList'),
    btnApplyFilters: required<HTMLButtonElement>('btnApplyFilters'),
    btnClearFilters: required<HTMLButtonElement>('btnClearFilters'),
    elementResults: required<HTMLDivElement>('elementResults'),
    spatialTree: required<HTMLDivElement>('spatialTree'),
    propsEmpty: required<HTMLDivElement>('propsEmpty'),
    propsContent: required<HTMLDivElement>('propsContent'),
    propType: required<HTMLSpanElement>('propType'),
    propName: required<HTMLSpanElement>('propName'),
    propGlobalId: required<HTMLSpanElement>('propGlobalId'),
    propDescription: required<HTMLSpanElement>('propDescription'),
    propStory: required<HTMLSpanElement>('propStory'),
    propAttributes: required<HTMLDivElement>('propAttributes'),
    viewpointName: required<HTMLInputElement>('viewpointName'),
    btnSaveViewpoint: required<HTMLButtonElement>('btnSaveViewpoint'),
    btnApplySelectedViewpoint: required<HTMLButtonElement>('btnApplySelectedViewpoint'),
    btnDeleteSelectedViewpoint: required<HTMLButtonElement>('btnDeleteSelectedViewpoint'),
    viewpointList: required<HTMLDivElement>('viewpointList'),
    issueTitle: required<HTMLInputElement>('issueTitle'),
    issueDescription: required<HTMLTextAreaElement>('issueDescription'),
    issuePriority: required<HTMLSelectElement>('issuePriority'),
    issueStatus: required<HTMLSelectElement>('issueStatus'),
    issueAssignee: required<HTMLInputElement>('issueAssignee'),
    btnCreateIssue: required<HTMLButtonElement>('btnCreateIssue'),
    btnDeleteIssue: required<HTMLButtonElement>('btnDeleteIssue'),
    issuesList: required<HTMLDivElement>('issuesList'),
    issueCommentInput: required<HTMLInputElement>('issueCommentInput'),
    btnAddIssueComment: required<HTMLButtonElement>('btnAddIssueComment'),
    issueComments: required<HTMLDivElement>('issueComments'),
  };

  private components!: OBC.Components;
  private world!: OBC.SimpleWorld<OBC.SimpleScene, OBC.OrthoPerspectiveCamera, OBCF.PostproductionRenderer>;
  private ifcLoader!: OBC.IfcLoader;
  private fragments!: OBC.FragmentsManager;
  private clipper!: OBC.Clipper;
  private hider!: OBC.Hider;
  private raycaster!: any;
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
  private gridVisible = true;
  private backgroundColor = DEFAULT_BACKGROUND_COLOR;
  private gridHelper: THREE.Object3D | null = null;
  private readonly appliedModelOpacity = new Map<string, number>();
  private hiddenLineColorOverride = false;
  private consistentLightOverride = false;
  private savedLightStates: { light: THREE.Light; visible: boolean; intensity: number }[] = [];

  private edgeOverlays: THREE.LineSegments[] = [];
  private readonly edgeMaterial = new THREE.LineBasicMaterial({ color: 0xc8145c, transparent: true, opacity: 0.65 });
  private modelObjects: THREE.Object3D[] = [];
  private readonly federatedModels = new Map<string, FederatedModelRecord>();
  private readonly modelIdAliases = new Map<string, string>();
  private readonly registeringModelIds = new Set<string>();
  private modelIndices = new Map<string, ModelIndex>();
  private readonly pendingModelMetaQueue: Array<{ fileName: string; sizeBytes: number }> = [];
  private lastHitPoint: THREE.Vector3 | null = null;
  private pendingIssuePoint: THREE.Vector3 | null = null;
  private viewpoints: SavedViewpoint[] = [];
  private selectedViewpointId: string | null = null;
  private issues: IssueRecord[] = [];
  private activeIssueId: string | null = null;
  private lastPointerDown = { x: 0, y: 0 };
  private pointerDragged = false;
  private frameCount = 0;
  private fpsLastTs = performance.now();
  private isModelLoading = false;
  private loadRequestId = 0;
  private suppressAutoFit = false;
  private activeGizmoModelId: string | null = null;
  private gizmoDragging = false;
  private readonly cubeTarget = new THREE.Vector3();
  private readonly cubeCamDir = new THREE.Vector3();
  private shaderWarningFilterInstalled = false;

  constructor() {
    this.bindUiEvents();
  }

  async init(): Promise<void> {
    try {
      this.setStatus('Initializing BTC IFC Viewer...');
      this.installShaderWarningFilter();
      await this.initEngine();
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
      this.setStatus('Ready - load IFC model(s)');
      this.startFpsMonitor();
    } catch (error) {
      this.setStatus(`Initialization failed: ${serializeError(error)}`);
      console.error(error);
    }
  }

  private installShaderWarningFilter(): void {
    if (this.shaderWarningFilterInstalled) return;
    const originalWarn = console.warn.bind(console);
    console.warn = (...args: unknown[]) => {
      const header = typeof args[0] === 'string' ? args[0] : '';
      const payload = args
        .map((entry) => (typeof entry === 'string' ? entry : ''))
        .join(' ');
      const isThreeProgramLog = header.includes('THREE.WebGLProgram: Program Info Log:');
      const isKnownNoise = payload.includes('dyn_index_vec4_float4_int');
      if (isThreeProgramLog && isKnownNoise) return;
      originalWarn(...args);
    };
    this.shaderWarningFilterInstalled = true;
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

    this.dom.btnExportScreenshot.addEventListener('click', () => this.exportScreenshot());
    this.dom.btnExportState.addEventListener('click', () => this.exportViewerState());
    this.dom.btnImportState.addEventListener('click', () => this.dom.importStateInput.click());
    this.dom.importStateInput.addEventListener('change', (event) => {
      const file = (event.target as HTMLInputElement).files?.[0];
      if (!file) return;
      this.fireAndForget(this.importViewerState(file), 'Import viewer state');
      this.dom.importStateInput.value = '';
    });

    this.dom.tabButtons.forEach((button) => {
      button.addEventListener('click', () => this.activateTab(button.dataset.tab || 'explorer'));
    });
    this.bindDockEvents();
    this.bindModelBrowserEvents();
    this.bindFederationTreeEvents();

    this.dom.btnModeOrbit.addEventListener('click', () => this.applyNavigationMode('Orbit'));
    this.dom.btnModePlan.addEventListener('click', () => this.applyNavigationMode('Plan'));
    this.dom.btnModeFirstPerson.addEventListener('click', () => this.applyNavigationMode('FirstPerson'));
    this.dom.btnFitAll.addEventListener('click', () => this.fitToModel());
    this.dom.btnFront.addEventListener('click', () => this.setFrontView());
    this.dom.btnTop.addEventListener('click', () => this.setTopView());
    this.dom.cubeFront.addEventListener('click', () => this.setFrontView());
    this.dom.cubeTop.addEventListener('click', () => this.setTopView());
    this.dom.cubeRight.addEventListener('click', () => this.setRightView());
    this.dom.cubeLeft.addEventListener('click', () => this.setLeftView());
    this.dom.cubeBack.addEventListener('click', () => this.setBackView());
    this.dom.cubeBottom.addEventListener('click', () => this.setBottomView());

    this.dom.btnSelectSingle.addEventListener('click', () => this.applySelectionMode('single'));
    this.dom.btnSelectMulti.addEventListener('click', () => this.applySelectionMode('multi'));

    this.dom.btnIsolate.addEventListener('click', () => {
      const selectionMap = this.getValidModelIdMap(this.selectedItems);
      if (isMapEmpty(selectionMap)) {
        this.setStatus('No selection to isolate');
        return;
      }
      this.fireAndForget((async () => {
        await this.hider.isolate(cloneMap(selectionMap));
        await this.updateVisibilityCount();
        this.setStatus('Selection isolated');
      })(), 'Isolate selection');
    });

    this.dom.btnHide.addEventListener('click', () => {
      const selectionMap = this.getValidModelIdMap(this.selectedItems);
      if (isMapEmpty(selectionMap)) {
        this.setStatus('No selection to hide');
        return;
      }
      this.fireAndForget((async () => {
        await this.hider.set(false, cloneMap(selectionMap));
        await this.updateVisibilityCount();
        this.setStatus('Selection hidden');
      })(), 'Hide selection');
    });

    this.dom.btnShow.addEventListener('click', () => {
      const selectionMap = this.getValidModelIdMap(this.selectedItems);
      if (isMapEmpty(selectionMap)) {
        this.setStatus('No selection to show');
        return;
      }
      this.fireAndForget((async () => {
        await this.hider.set(true, cloneMap(selectionMap));
        await this.updateVisibilityCount();
        this.setStatus('Selection shown');
      })(), 'Show selection');
    });

    this.dom.btnResetVisibility.addEventListener('click', () => {
      this.fireAndForget((async () => {
        await this.hider.set(true);
        this.clearFilterChecks();
        await this.updateVisibilityCount();
        this.setStatus('Visibility reset');
      })(), 'Reset visibility');
    });

    this.dom.btnSectionX.addEventListener('click', () => {
      this.addSectionPlane(new THREE.Vector3(1, 0, 0));
      this.dom.btnSectionX.classList.toggle('active');
    });
    this.dom.btnSectionY.addEventListener('click', () => {
      this.addSectionPlane(new THREE.Vector3(0, 1, 0));
      this.dom.btnSectionY.classList.toggle('active');
    });
    this.dom.btnSectionZ.addEventListener('click', () => {
      this.addSectionPlane(new THREE.Vector3(0, 0, 1));
      this.dom.btnSectionZ.classList.toggle('active');
    });
    this.dom.btnSectionBox.addEventListener('click', () => {
      this.createSectionBox();
      this.dom.btnSectionBox.classList.add('active');
    });
    this.dom.btnClearSections.addEventListener('click', () => {
      this.clearSections();
      this.setStatus('Sections cleared');
    });

    this.dom.btnMeasureLength.addEventListener('click', () => this.setMeasureMode(this.measureMode === 'length' ? 'none' : 'length'));
    this.dom.btnMeasureArea.addEventListener('click', () => this.setMeasureMode(this.measureMode === 'area' ? 'none' : 'area'));
    this.dom.btnClearMeasurements.addEventListener('click', () => {
      this.clearMeasurements();
      this.setStatus('Measurements cleared');
    });

    this.dom.btnTransparency.addEventListener('click', () => {
      this.xrayEnabled = !this.xrayEnabled;
      this.dom.btnTransparency.classList.toggle('active', this.xrayEnabled);
      this.applyXRay();
      this.fireAndForget(this.fragments.core.update(true), 'Toggle x-ray');
      this.persistLocalState();
      this.setStatus(this.xrayEnabled ? 'X-ray enabled' : 'X-ray disabled');
    });

    this.dom.btnWireframe.addEventListener('click', () => {
      this.edgesEnabled = !this.edgesEnabled;
      this.dom.btnWireframe.classList.toggle('active', this.edgesEnabled);
      this.applyEdges();
      this.persistLocalState();
      this.setStatus(this.edgesEnabled ? 'Edge overlay enabled' : 'Edge overlay disabled');
    });

    this.dom.btnIssuePinMode.addEventListener('click', () => {
      this.issuePinMode = !this.issuePinMode;
      this.dom.btnIssuePinMode.classList.toggle('active', this.issuePinMode);
      this.dom.viewerHint.hidden = !this.issuePinMode;
      this.setStatus(this.issuePinMode ? 'Issue pin mode active' : 'Issue pin mode disabled');
    });

    this.dom.toggleGrid.addEventListener('change', () => {
      this.setGridVisible(this.dom.toggleGrid.checked, true);
      this.persistLocalState();
    });

    this.dom.backgroundColorInput.addEventListener('input', () => {
      this.setBackgroundColor(this.dom.backgroundColorInput.value, false);
    });

    this.dom.backgroundColorInput.addEventListener('change', () => {
      this.setBackgroundColor(this.dom.backgroundColorInput.value, true);
      this.persistLocalState();
    });

    this.dom.backgroundPresetButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const color = button.dataset.bgPreset;
        if (!color) return;
        this.setBackgroundColor(color, true);
        this.persistLocalState();
      });
    });

    this.dom.visualStyleSelect.addEventListener('change', () => {
      const style = this.parseVisualStyle(this.dom.visualStyleSelect.value);
      this.fireAndForget(this.setVisualStyle(style, true, true), 'Set visual style');
    });

    this.dom.btnSearch.addEventListener('click', () => {
      this.fireAndForget(this.searchElements(this.dom.searchInput.value.trim()), 'Search');
    });
    this.dom.searchInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') this.fireAndForget(this.searchElements(this.dom.searchInput.value.trim()), 'Search');
    });
    this.dom.btnClearSearch.addEventListener('click', () => {
      this.dom.searchInput.value = '';
      this.dom.elementResults.innerHTML = '';
      this.setStatus('Search cleared');
    });

    this.dom.btnApplyFilters.addEventListener('click', () => {
      this.fireAndForget(this.applyFilters(), 'Apply filters');
    });
    this.dom.btnClearFilters.addEventListener('click', () => {
      this.fireAndForget((async () => {
        this.clearFilterChecks();
        await this.hider.set(true);
        await this.updateVisibilityCount();
        this.setStatus('Filters reset');
      })(), 'Reset filters');
    });

    this.dom.btnSaveViewpoint.addEventListener('click', () => {
      this.fireAndForget(this.saveViewpoint(), 'Save viewpoint');
    });
    this.dom.btnApplySelectedViewpoint.addEventListener('click', () => {
      this.fireAndForget(this.applySelectedViewpoint(), 'Apply viewpoint');
    });
    this.dom.btnDeleteSelectedViewpoint.addEventListener('click', () => this.deleteSelectedViewpoint());

    this.dom.btnCreateIssue.addEventListener('click', () => {
      this.fireAndForget(this.createIssueFromCurrentContext(), 'Create issue');
    });
    this.dom.btnDeleteIssue.addEventListener('click', () => this.deleteSelectedIssue());
    this.dom.btnAddIssueComment.addEventListener('click', () => this.addCommentToActiveIssue());

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
      this.dom.viewerContainer.style.outline = '2px dashed #c8145c';
    });

    this.dom.viewerContainer.addEventListener('dragleave', () => {
      this.dom.viewerContainer.style.outline = 'none';
    });

    this.dom.viewerContainer.addEventListener('drop', (event) => {
      event.preventDefault();
      this.dom.viewerContainer.style.outline = 'none';
      const files = Array.from(event.dataTransfer?.files ?? []);
      if (files.length === 0) return;
      const ifcFiles = files.filter((file) => file.name.toLowerCase().endsWith('.ifc'));
      if (ifcFiles.length === 0) {
        this.setStatus('Only IFC files are supported');
        return;
      }
      this.fireAndForget(this.loadIfcFiles(ifcFiles), 'Load IFC files');
    });
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

  private bindDockEvents(): void {
    const toggles = Array.from(this.dom.viewerDock.querySelectorAll<HTMLButtonElement>('[data-dock-toggle]'));
    toggles.forEach((toggle) => {
      toggle.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        const group = toggle.closest<HTMLElement>('.dock-group');
        if (!group) return;
        const shouldOpen = !group.classList.contains('open');
        this.closeDockGroups();
        if (shouldOpen) group.classList.add('open');
      });
    });

    this.dom.viewerDock.querySelectorAll<HTMLButtonElement>('.dock-tool-btn').forEach((button) => {
      button.addEventListener('click', () => this.closeDockGroups());
    });

    document.addEventListener('pointerdown', (event) => {
      const target = event.target as HTMLElement | null;
      if (!target?.closest('#viewerDock')) this.closeDockGroups();
    });
  }

  private closeDockGroups(): void {
    this.dom.viewerDock.querySelectorAll<HTMLElement>('.dock-group.open').forEach((group) => {
      group.classList.remove('open');
    });
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
      this.applyModelOpacity(modelId, opacity);
      const card = opacityInput.closest<HTMLElement>('.federated-opacity');
      const valueLabel = card?.querySelector<HTMLElement>('[data-opacity-value]');
      if (valueLabel) valueLabel.textContent = `${Math.round(this.clamp(opacity, 0, 1) * 100)}%`;
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

  private async initEngine(): Promise<void> {
    this.components = new OBC.Components();
    const worlds = this.components.get(OBC.Worlds);

    this.world = worlds.create<OBC.SimpleScene, OBC.OrthoPerspectiveCamera, OBCF.PostproductionRenderer>();
    this.world.scene = new OBC.SimpleScene(this.components);
    this.world.scene.setup();
    this.world.scene.three.background = new THREE.Color(this.backgroundColor);

    this.world.renderer = new OBCF.PostproductionRenderer(this.components, this.dom.viewerContainer);

    // Hook into the render loop to clear composer targets before each frame in PEN mode.
    // PEN mode (PostproductionAspect.PEN) skips the BasePass, so the EffectComposer's
    // read/write buffers never get cleared — causing ghost lines from previous frames.
    this.world.renderer.onBeforeUpdate.add(() => {
      const postRenderer = this.getPostproductionRenderer();
      const post = postRenderer?.postproduction;
      if (!post?.enabled || !post.composer) return;
      // PEN = 1, PEN_SHADOWS = 2 — these use EdgeDetectionPass without a prior clear
      const isPenStyle = post.style === OBCF.PostproductionAspect.PEN
        || post.style === OBCF.PostproductionAspect.PEN_SHADOWS;
      if (!isPenStyle) return;
      const renderer = postRenderer!.three;
      const bgColor = new THREE.Color(this.backgroundColor);
      renderer.setClearColor(bgColor, 1);
      renderer.setRenderTarget(post.composer.renderTarget1);
      renderer.clear();
      renderer.setRenderTarget(post.composer.renderTarget2);
      renderer.clear();
      renderer.setRenderTarget(null);
    });
    this.world.camera = new OBC.OrthoPerspectiveCamera(this.components);
    await this.world.camera.controls.setLookAt(18, 18, 18, 0, 0, 0);

    this.components.init();

    // Temporary debug exposure — remove in production
    (window as any).__viewer = this;
    (window as any).__components = this.components;
    (window as any).__world = this.world;
    (window as any).__THREE = THREE;

    const grids = this.components.get(OBC.Grids);
    const grid = grids.create(this.world);
    this.gridHelper = grid as unknown as THREE.Object3D;
    this.gridHelper.visible = this.gridVisible;
    if (grid.material?.uniforms?.uColor) grid.material.uniforms.uColor.value = new THREE.Color(0x25334a);

    this.ifcLoader = this.components.get(OBC.IfcLoader);
    await this.ifcLoader.setup({
      autoSetWasm: false,
      wasm: {
        path: 'https://unpkg.com/web-ifc@0.0.74/',
        absolute: true,
      },
    });
    this.ifcLoader.settings.webIfc.CIRCLE_SEGMENTS = 24;

    this.fragments = this.components.get(OBC.FragmentsManager);
    const workerUrl = 'https://thatopen.github.io/engine_fragment/resources/worker.mjs';
    const fetchedWorker = await fetch(workerUrl);
    const workerBlob = await fetchedWorker.blob();
    const workerFile = new File([workerBlob], 'worker.mjs', { type: 'text/javascript' });
    this.fragments.init(URL.createObjectURL(workerFile));
    this.fragments.core.settings.graphicsQuality = 1;

    this.world.camera.controls.addEventListener('update', () => {
      this.fireAndForget(this.fragments.core.update(), 'Camera update');
      this.updateViewCubeFromCamera();
    });

    this.fragments.core.onModelLoaded.add((model) => {
      const modelId = this.getModelInternalId(model, '');
      this.fireAndForget(this.onModelAdded(modelId || String(model?.modelId ?? ''), model), 'Register model');
    });

    this.fragments.core.models.materials.list.onItemSet.add(({ value: material }) => {
      if (!('isLodMaterial' in material && (material as unknown as { isLodMaterial: boolean }).isLodMaterial)) {
        const cast = material as THREE.Material & {
          polygonOffset?: boolean;
          polygonOffsetFactor?: number;
          polygonOffsetUnits?: number;
        };
        cast.polygonOffset = true;
        cast.polygonOffsetFactor = 1;
        cast.polygonOffsetUnits = 1;
      }
    });

    this.clipper = this.components.get(OBC.Clipper);
    this.clipper.enabled = false;
    this.hider = this.components.get(OBC.Hider);
    const raycasters = this.components.get(OBC.Raycasters);
    this.raycaster = raycasters.get(this.world);

    this.lengthMeasurement = this.components.get(OBCF.LengthMeasurement);
    this.lengthMeasurement.world = this.world;
    this.lengthMeasurement.enabled = false;

    this.areaMeasurement = this.components.get(OBCF.AreaMeasurement);
    this.areaMeasurement.world = this.world;
    this.areaMeasurement.enabled = false;

    this.markerManager = this.components.get(OBCF.Marker);
    this.markerManager.threshold = 64;
    this.markerManager.autoCluster = true;

    this.transformControls = new TransformControls(this.world.camera.three, this.world.renderer.three.domElement);
    this.transformControls.setSize(0.75);
    this.transformControls.setSpace('world');
    this.transformControls.enabled = false;
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
      this.fireAndForget(this.fragments.core.update(true), 'Gizmo update');
    });
    this.transformControls.addEventListener('mouseUp', () => {
      if (!this.activeGizmoModelId) return;
      const model = this.federatedModels.get(this.activeGizmoModelId);
      if (!model) return;
      this.updateModelOffsetsFromObject(model);
      this.renderModelBrowser();
      this.renderFederatedTree();
      this.setStatus(`Gizmo updated: ${model.fileName}`);
    });
    this.transformControlsHelper = this.transformControls.getHelper();
    this.transformControlsHelper.visible = false;
    this.world.scene.three.add(this.transformControlsHelper);

    this.updateViewCubeFromCamera();
    await this.updateVisibilityCount();
  }

  private async onModelAdded(modelId: string, model: any): Promise<void> {
    const resolvedModelId = this.getModelInternalId(model, String(modelId || ''));
    if (!resolvedModelId) return;
    if (this.federatedModels.has(resolvedModelId) || this.registeringModelIds.has(resolvedModelId)) return;

    this.registeringModelIds.add(resolvedModelId);
    try {
      await this.waitForModelReady(model);
      model.useCamera(this.world.camera.three);
      if (typeof model?.graphicsQuality === 'number') model.graphicsQuality = 1;
      this.world.scene.three.add(model.object);

      const modelObject = model.object as THREE.Object3D;
      if (!this.modelObjects.includes(modelObject)) this.modelObjects.push(modelObject);

      const ids = await model.getItemsIdsWithGeometry();
      this.dom.emptyState.hidden = true;

      const meta = this.pendingModelMetaQueue.shift();
      const fileName = meta?.fileName || String(model?.modelId ?? resolvedModelId);
      this.registerModelAlias(resolvedModelId, resolvedModelId);
      this.registerModelAlias(String(modelId || ''), resolvedModelId);
      this.registerModelAlias(fileName, resolvedModelId);
      this.registerModelAlias(String(model?.modelId ?? ''), resolvedModelId);
      const modelRecord: FederatedModelRecord = {
        modelId: resolvedModelId,
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
      this.federatedModels.set(resolvedModelId, modelRecord);
      this.updateElementCounter();
      this.updateHeaderModelName(fileName);
      this.renderModelBrowser();
      this.renderFederatedTree();

      await this.indexModel(resolvedModelId, model);
      this.renderSpatialTree();
      this.renderModelBrowser();
      this.renderFederatedTree();

      await this.setVisualStyle(this.visualStyle, false, false);

      if (!this.suppressAutoFit) this.fitToModel();
      await this.updateVisibilityCount();

      this.renderClassFilters();
      this.renderLevelFilters();

      this.refreshIssueMarkers();
      this.setStatus(`Model loaded: ${fileName}`);
    } finally {
      this.registeringModelIds.delete(resolvedModelId);
    }
  }

  private async waitForModelReady(model: any, timeoutMs = 10000): Promise<void> {
    const startedAt = performance.now();
    while (performance.now() - startedAt < timeoutMs) {
      if (!model?.isBusy) return;
      await new Promise((resolve) => window.setTimeout(resolve, 40));
    }
  }

  private async indexModel(modelId: string, model: any): Promise<void> {
    const resolvedModelId = this.getModelInternalId(model, String(modelId || ''));
    if (!resolvedModelId) return;
    const itemIds = await model.getItemsIdsWithGeometry() as number[];
    const idsSet = new Set(itemIds);

    const classes = new Map<string, Set<number>>();
    const itemClassById = new Map<number, string>();
    const categories = await model.getItemsWithGeometryCategories() as Array<string | null>;
    for (let i = 0; i < categories.length; i += 1) {
      const category = categories[i] ?? 'Unknown';
      const id = itemIds[i];
      if (typeof id !== 'number') continue;
      if (!classes.has(category)) classes.set(category, new Set<number>());
      classes.get(category)?.add(id);
      itemClassById.set(id, category);
    }

    const itemNames = new Map<number, string>();
    const itemToLevel = new Map<number, string>();
    const levels = new Map<string, Set<number>>();
    const spatial = await model.getSpatialStructure() as SpatialTreeItem;

    // Read element names + level assignment from ContainedInStructure relation.
    const chunkSize = 360;
    for (let start = 0; start < itemIds.length; start += chunkSize) {
      const chunk = itemIds.slice(start, start + chunkSize);
      const itemsData = await model.getItemsData(chunk, {
        attributesDefault: true,
        relations: {
          ContainedInStructure: { attributes: true, relations: true },
        },
        relationsDefault: { attributes: false, relations: false },
      });

      for (let i = 0; i < chunk.length; i += 1) {
        const localId = chunk[i];
        const data = (itemsData[i] || {}) as Record<string, unknown>;
        const category = itemClassById.get(localId) ?? 'Element';
        itemNames.set(localId, this.getModelTreeItemLabel(data, localId, category));
        const levelName = this.extractStoreyNameFromItemData(data);
        if (!levelName) continue;
        itemToLevel.set(localId, levelName);
        if (!levels.has(levelName)) levels.set(levelName, new Set<number>());
        levels.get(levelName)?.add(localId);
      }
    }

    // Spatial fallback: ensure storey names are loaded and assign ungrouped items.
    const storeyIds = new Set<number>();
    const collectStoreys = (node: SpatialTreeItem): void => {
      const category = (node.category ?? '').toUpperCase();
      if (category.includes('IFCBUILDINGSTOREY') && node.localId !== null) storeyIds.add(node.localId);
      for (const child of node.children ?? []) collectStoreys(child);
    };
    collectStoreys(spatial);

    const unknownStoreyIds = [...storeyIds].filter((id) => !itemNames.has(id));
    for (let start = 0; start < unknownStoreyIds.length; start += chunkSize) {
      const chunk = unknownStoreyIds.slice(start, start + chunkSize);
      const rows = await model.getItemsData(chunk, {
        attributesDefault: true,
        relationsDefault: { attributes: false, relations: false },
      });
      for (let i = 0; i < chunk.length; i += 1) {
        const localId = chunk[i];
        const data = (rows[i] || {}) as Record<string, unknown>;
        itemNames.set(localId, this.getModelTreeItemLabel(data, localId, 'Storey'));
      }
    }

    const walkSpatial = (node: SpatialTreeItem, activeStorey: string | null): void => {
      const category = (node.category ?? '').toUpperCase();
      let nextStorey = activeStorey;
      if (category.includes('IFCBUILDINGSTOREY') && node.localId !== null) {
        nextStorey = itemNames.get(node.localId) ?? `Storey ${node.localId}`;
      }
      if (node.localId !== null && nextStorey && idsSet.has(node.localId) && !itemToLevel.has(node.localId)) {
        itemToLevel.set(node.localId, nextStorey);
        if (!levels.has(nextStorey)) levels.set(nextStorey, new Set<number>());
        levels.get(nextStorey)?.add(node.localId);
      }
      for (const child of node.children ?? []) walkSpatial(child, nextStorey);
    };
    walkSpatial(spatial, null);

    // Load names for non-geometry nodes used by the browser spatial tree.
    const spatialIds = new Set<number>();
    this.collectSpatialTreeIds(spatial, spatialIds);
    const missingSpatialIds = [...spatialIds].filter((id) => !itemNames.has(id));
    for (let start = 0; start < missingSpatialIds.length; start += chunkSize) {
      const chunk = missingSpatialIds.slice(start, start + chunkSize);
      const rows = await model.getItemsData(chunk, {
        attributesDefault: true,
        relationsDefault: { attributes: false, relations: false },
      });
      for (let i = 0; i < chunk.length; i += 1) {
        const localId = chunk[i];
        const data = (rows[i] || {}) as Record<string, unknown>;
        itemNames.set(localId, this.getModelTreeItemLabel(data, localId, 'Item'));
      }
    }

    const spatialRoot = this.buildSpatialBrowserTree(spatial, itemNames, idsSet);

    this.modelIndices.set(resolvedModelId, {
      modelId: resolvedModelId,
      allIds: new Set(itemIds),
      classes,
      levels,
      itemToLevel,
      itemNames,
      spatialRoot,
    });
  }

  private renderSpatialTree(): void {
    this.dom.spatialTree.innerHTML = '';
    const rows: string[] = [];
    for (const [modelId, index] of this.modelIndices.entries()) {
      rows.push(`<div class="tree-item"><strong>${modelId}</strong></div>`);
      for (const [level, ids] of index.levels.entries()) {
        rows.push(`<div class="tree-item" data-model-id="${modelId}" data-level="${level}">Level: ${level} (${ids.size})</div>`);
      }
      if (index.levels.size === 0) rows.push('<div class="tree-item">No storeys detected</div>');
    }
    this.dom.spatialTree.innerHTML = rows.join('');

    this.dom.spatialTree.querySelectorAll<HTMLElement>('[data-model-id][data-level]').forEach((element) => {
      element.addEventListener('click', () => {
        const modelId = element.dataset.modelId;
        const level = element.dataset.level;
        if (!modelId || !level) return;
        const index = this.modelIndices.get(modelId);
        const ids = index?.levels.get(level);
        if (!ids || ids.size === 0) return;
        const map = this.getValidModelIdMap({ [modelId]: new Set(ids) });
        if (isMapEmpty(map)) return;
        this.fireAndForget((async () => {
          await this.hider.isolate(map);
          await this.updateVisibilityCount();
          this.setStatus(`Isolated level ${level}`);
        })(), 'Isolate spatial level');
      });
    });
  }

  private renderClassFilters(): void {
    const classes = new Set<string>();
    for (const index of this.modelIndices.values()) {
      for (const className of index.classes.keys()) classes.add(className);
    }
    const sorted = [...classes].sort((a, b) => a.localeCompare(b));
    if (sorted.length === 0) {
      this.dom.classFilterList.innerHTML = '<div class="filter-item">No classes detected</div>';
      return;
    }
    this.dom.classFilterList.innerHTML = sorted
      .map((className) => `
        <label class="filter-item">
          <input type="checkbox" data-filter-type="class" value="${className}" />
          <span>${className}</span>
        </label>
      `)
      .join('');
  }

  private renderLevelFilters(): void {
    const levels = new Set<string>();
    for (const index of this.modelIndices.values()) {
      for (const levelName of index.levels.keys()) levels.add(levelName);
    }
    const sorted = [...levels].sort((a, b) => a.localeCompare(b));
    if (sorted.length === 0) {
      this.dom.levelFilterList.innerHTML = '<div class="filter-item">No levels detected</div>';
      return;
    }
    this.dom.levelFilterList.innerHTML = sorted
      .map((levelName) => `
        <label class="filter-item">
          <input type="checkbox" data-filter-type="level" value="${levelName}" />
          <span>${levelName}</span>
        </label>
      `)
      .join('');
  }

  private clearFilterChecks(): void {
    this.dom.classFilterList.querySelectorAll<HTMLInputElement>('input[type="checkbox"]').forEach((checkbox) => {
      checkbox.checked = false;
    });
    this.dom.levelFilterList.querySelectorAll<HTMLInputElement>('input[type="checkbox"]').forEach((checkbox) => {
      checkbox.checked = false;
    });
  }

  private updateElementCounter(): void {
    let total = 0;
    for (const model of this.federatedModels.values()) total += model.elementCount;
    this.dom.elementCount.textContent = `${total} elements`;
  }

  private updateHeaderModelName(lastLoadedName?: string): void {
    const count = this.federatedModels.size;
    if (count === 0) {
      this.dom.modelName.textContent = 'No model loaded';
      return;
    }
    if (count === 1) {
      const only = this.federatedModels.values().next().value as FederatedModelRecord | undefined;
      this.dom.modelName.textContent = only?.fileName ?? '1 model loaded';
      return;
    }
    if (lastLoadedName) {
      this.dom.modelName.textContent = `${count} models loaded (latest: ${lastLoadedName})`;
      return;
    }
    this.dom.modelName.textContent = `${count} models loaded`;
  }

  private getClassIdsForModelLevel(modelId: string, level: string, className: string): Set<number> {
    const index = this.modelIndices.get(modelId);
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

  private getLevelClassEntries(modelId: string, level: string): Array<{ className: string; count: number }> {
    const index = this.modelIndices.get(modelId);
    const levelIds = index?.levels.get(level);
    if (!index || !levelIds || levelIds.size === 0) return [];

    const entries: Array<{ className: string; count: number }> = [];
    for (const className of index.classes.keys()) {
      const ids = this.getClassIdsForModelLevel(modelId, level, className);
      if (ids.size === 0) continue;
      entries.push({ className, count: ids.size });
    }
    entries.sort((a, b) => a.className.localeCompare(b.className));
    return entries;
  }

  private readPrimitiveValue(value: unknown): string {
    const normalized = this.unwrapIfcValue(value);
    if (normalized === null || normalized === undefined) return '';
    if (typeof normalized === 'string') return normalized.trim();
    if (typeof normalized === 'number' || typeof normalized === 'boolean' || typeof normalized === 'bigint') return String(normalized);
    return '';
  }

  private getRecordValueCaseInsensitive(record: Record<string, unknown>, preferredKey: string): unknown {
    if (Object.prototype.hasOwnProperty.call(record, preferredKey)) return record[preferredKey];
    const lowerPreferred = preferredKey.toLowerCase();
    const matched = Object.keys(record).find((key) => key.toLowerCase() === lowerPreferred);
    return matched ? record[matched] : undefined;
  }

  private findNameLikeValue(value: unknown, visited: WeakSet<object>, depth: number): string {
    if (depth > 5) return '';
    const normalized = this.unwrapIfcValue(value);
    if (normalized === null || normalized === undefined) return '';
    if (typeof normalized === 'string') return normalized.trim();
    if (typeof normalized === 'number' || typeof normalized === 'boolean' || typeof normalized === 'bigint') {
      return String(normalized);
    }
    if (Array.isArray(normalized)) {
      for (const entry of normalized) {
        const found = this.findNameLikeValue(entry, visited, depth + 1);
        if (found) return found;
      }
      return '';
    }
    if (typeof normalized !== 'object') return '';

    const record = normalized as Record<string, unknown>;
    if (visited.has(record)) return '';
    visited.add(record);

    const preferredKeys = ['Name', 'LongName', 'ObjectType', 'PredefinedType', 'Tag'];
    for (const key of preferredKeys) {
      const primitive = this.readPrimitiveValue(this.getRecordValueCaseInsensitive(record, key));
      if (primitive) return primitive;
    }

    for (const [key, entry] of Object.entries(record)) {
      const lower = key.toLowerCase();
      if (lower === 'name' || lower.endsWith('name') || lower.endsWith('objecttype')) {
        const primitive = this.readPrimitiveValue(entry);
        if (primitive) return primitive;
      }
    }

    for (const nested of Object.values(record)) {
      const found = this.findNameLikeValue(nested, visited, depth + 1);
      if (found) return found;
    }
    return '';
  }

  private getModelTreeItemLabel(data: Record<string, unknown>, localId: number, defaultCategory: string): string {
    const label = this.readPrimitiveValue(this.getRecordValueCaseInsensitive(data, 'Name'))
      || this.readPrimitiveValue(this.getRecordValueCaseInsensitive(data, 'LongName'))
      || this.readPrimitiveValue(this.getRecordValueCaseInsensitive(data, 'ObjectType'))
      || this.readPrimitiveValue(this.getRecordValueCaseInsensitive(data, 'PredefinedType'))
      || this.findNameLikeValue(data, new WeakSet<object>(), 0);
    if (label) return label;
    const category = this.toPropertyString(data._category, defaultCategory) || defaultCategory;
    return `${category} ${localId}`;
  }

  private extractStoreyNameFromItemData(data: Record<string, unknown>): string | null {
    const visited = new WeakSet<object>();
    const fromContained = this.findStoreyNameInValue(data.ContainedInStructure, visited, 0);
    if (fromContained) return fromContained;
    const fromDecomposes = this.findStoreyNameInValue(data.Decomposes, visited, 0);
    if (fromDecomposes) return fromDecomposes;
    return null;
  }

  private findStoreyNameInValue(value: unknown, visited: WeakSet<object>, depth: number): string | null {
    if (depth > 8) return null;
    const unwrapped = this.unwrapIfcValue(value);
    if (unwrapped === null || unwrapped === undefined) return null;

    if (Array.isArray(unwrapped)) {
      for (const entry of unwrapped) {
        const name = this.findStoreyNameInValue(entry, visited, depth + 1);
        if (name) return name;
      }
      return null;
    }

    if (typeof unwrapped !== 'object') return null;
    const record = unwrapped as Record<string, unknown>;
    if (visited.has(record)) return null;
    visited.add(record);

    const category = this.toPropertyString(record._category, '').toUpperCase();
    if (category.includes('IFCBUILDINGSTOREY')) {
      let storeyName = this.readPrimitiveValue(this.getRecordValueCaseInsensitive(record, 'Name'))
        || this.readPrimitiveValue(this.getRecordValueCaseInsensitive(record, 'LongName'));
      if (!storeyName) {
        for (const nested of Object.values(record)) {
          const nestedName = this.findNameLikeValue(nested, visited, depth + 1);
          if (nestedName) {
            storeyName = nestedName;
            break;
          }
        }
      }
      if (storeyName) return storeyName;
      return null;
    }

    for (const nested of Object.values(record)) {
      const name = this.findStoreyNameInValue(nested, visited, depth + 1);
      if (name) return name;
    }
    return null;
  }

  private collectSpatialTreeIds(node: SpatialTreeItem, target: Set<number>): void {
    if (node.localId !== null) target.add(node.localId);
    for (const child of node.children ?? []) this.collectSpatialTreeIds(child, target);
  }

  private buildSpatialBrowserTree(
    node: SpatialTreeItem,
    itemNames: Map<number, string>,
    geometryIds: Set<number>,
  ): BrowserTreeNode {
    const localId = node.localId;
    const category = node.category ?? 'Group';
    const children = (node.children ?? []).map((child) => this.buildSpatialBrowserTree(child, itemNames, geometryIds));

    let geometryCount = localId !== null && geometryIds.has(localId) ? 1 : 0;
    for (const child of children) geometryCount += child.geometryCount;

    const label = localId !== null
      ? (itemNames.get(localId) ?? `${category} ${localId}`)
      : (category || 'Structure');

    return {
      category,
      localId,
      label,
      geometryCount,
      children,
    };
  }

  private renderSpatialBrowserNode(modelId: string, index: ModelIndex, node: BrowserTreeNode, depth: number): string {
    const hasChildren = node.children.length > 0;
    const escapedModelId = escapeHtml(modelId);
    const categoryUpper = node.category.toUpperCase();
    const isStoreyNode = categoryUpper.includes('IFCBUILDINGSTOREY') && index.levels.has(node.label);
    const isElement = node.localId !== null && index.allIds.has(node.localId);
    const countText = node.geometryCount > 0 ? String(node.geometryCount) : (node.localId ?? '-').toString();

    if (!hasChildren || depth >= MAX_BROWSER_SPATIAL_DEPTH) {
      const leafContent = isElement && node.localId !== null
        ? `
          <button
            type="button"
            class="browser-action"
            data-browser-action="select-item"
            data-model-id="${escapedModelId}"
            data-local-id="${node.localId}"
            title="Select ${escapeHtml(node.label)}"
          >
            ${escapeHtml(node.label)}
          </button>
        `
        : `<span>${escapeHtml(node.label)}</span>`;
      const icon = isElement ? 'view_in_ar' : 'subdirectory_arrow_right';
      return `
        <div class="browser-leaf">
          <span class="material-icons-round">${icon}</span>
          ${leafContent}
          <span class="browser-count">${countText}</span>
        </div>
      `;
    }

    const visibleChildren = node.children.slice(0, MAX_BROWSER_SPATIAL_CHILDREN);
    const childrenMarkup = visibleChildren
      .map((child) => this.renderSpatialBrowserNode(modelId, index, child, depth + 1))
      .join('');
    const moreMarkup = node.children.length > MAX_BROWSER_SPATIAL_CHILDREN
      ? `<div class="browser-leaf"><span class="material-icons-round">more_horiz</span><span>${node.children.length - MAX_BROWSER_SPATIAL_CHILDREN} more nodes</span><span class="browser-count">+</span></div>`
      : '';

    const labelMarkup = isStoreyNode
      ? `
        <button
          type="button"
          class="browser-action"
          data-browser-action="isolate-level"
          data-model-id="${escapedModelId}"
          data-level="${escapeHtml(node.label)}"
          title="Isolate ${escapeHtml(node.label)}"
        >
          ${escapeHtml(node.label)}
        </button>
      `
      : `<span>${escapeHtml(node.label)}</span>`;

    return `
      <details class="browser-node">
        <summary class="browser-summary">
          <span class="browser-twist material-icons-round">chevron_right</span>
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

  private renderModelBrowser(): void {
    if (this.federatedModels.size === 0) {
      this.dom.modelBrowserTree.innerHTML = '<div class="tree-item">No models loaded yet</div>';
      return;
    }

    const modelMarkup = [...this.federatedModels.values()]
      .map((record) => {
        const modelId = String(record.modelId);
        const escapedModelId = escapeHtml(modelId);
        const index = this.modelIndices.get(modelId);
        const visibilitySuffix = record.visible ? '' : ' (Hidden)';

        if (!index) {
          return `
            <details class="browser-node is-model" open>
              <summary class="browser-summary">
                <span class="browser-twist material-icons-round">chevron_right</span>
                <button
                  type="button"
                  class="browser-action"
                  data-browser-action="select-model"
                  data-model-id="${escapedModelId}"
                  title="Select full model"
                >
                  ${escapeHtml(record.fileName)}${visibilitySuffix}
                </button>
                <span class="browser-count">${record.elementCount}</span>
              </summary>
              <div class="browser-children">
                <div class="browser-leaf">
                  <span class="material-icons-round">hourglass_top</span>
                  <span>Building model tree...</span>
                  <span class="browser-count">-</span>
                </div>
              </div>
            </details>
          `;
        }

        const levelEntries = [...index.levels.entries()].sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true }));
        const levelMarkup = levelEntries.slice(0, MAX_BROWSER_LEVELS).map(([levelName, ids]) => {
          const classEntries = this.getLevelClassEntries(modelId, levelName).slice(0, MAX_BROWSER_CLASSES_PER_LEVEL);
          const classMarkup = classEntries.map(({ className }) => {
            const classIds = [...this.getClassIdsForModelLevel(modelId, levelName, className)];
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
                  <span class="material-icons-round">view_in_ar</span>
                  <button
                    type="button"
                    class="browser-action"
                    data-browser-action="select-item"
                    data-model-id="${escapedModelId}"
                    data-local-id="${localId}"
                    title="Select ${escapeHtml(label)}"
                  >
                    ${escapeHtml(label)}
                  </button>
                  <span class="browser-count">${localId}</span>
                </div>
              `;
            }).join('');
            const hiddenCount = classIds.length - visibleIds.length;
            const moreElementsMarkup = hiddenCount > 0
              ? `<div class="browser-leaf"><span class="material-icons-round">more_horiz</span><span>${hiddenCount} more elements</span><span class="browser-count">+</span></div>`
              : '';

            return `
              <details class="browser-node">
                <summary class="browser-summary">
                  <span class="browser-twist material-icons-round">chevron_right</span>
                  <button
                    type="button"
                    class="browser-action"
                    data-browser-action="isolate-class-level"
                    data-model-id="${escapedModelId}"
                    data-level="${escapeHtml(levelName)}"
                    data-class="${escapeHtml(className)}"
                    title="Isolate ${escapeHtml(className)} in ${escapeHtml(levelName)}"
                  >
                    ${escapeHtml(className)}
                  </button>
                  <span class="browser-count">${classIds.length}</span>
                </summary>
                <div class="browser-children">
                  ${elementsMarkup || '<div class="browser-leaf"><span class="material-icons-round">hide_source</span><span>No elements</span><span class="browser-count">0</span></div>'}
                  ${moreElementsMarkup}
                </div>
              </details>
            `;
          }).join('');

          return `
            <details class="browser-node">
              <summary class="browser-summary">
                <span class="browser-twist material-icons-round">chevron_right</span>
                <button
                  type="button"
                  class="browser-action"
                  data-browser-action="isolate-level"
                  data-model-id="${escapedModelId}"
                  data-level="${escapeHtml(levelName)}"
                  title="Isolate level ${escapeHtml(levelName)}"
                >
                  ${escapeHtml(levelName)}
                </button>
                <span class="browser-count">${ids.size}</span>
              </summary>
              <div class="browser-children">
                ${classMarkup || '<div class="browser-leaf"><span class="material-icons-round">category</span><span>No classes</span><span class="browser-count">0</span></div>'}
              </div>
            </details>
          `;
        }).join('');

        const levelMoreMarkup = levelEntries.length > MAX_BROWSER_LEVELS
          ? `<div class="browser-leaf"><span class="material-icons-round">more_horiz</span><span>${levelEntries.length - MAX_BROWSER_LEVELS} more levels</span><span class="browser-count">+</span></div>`
          : '';

        const spatialRootNodes = index.spatialRoot?.children?.length ? index.spatialRoot.children : (index.spatialRoot ? [index.spatialRoot] : []);
        const spatialVisible = spatialRootNodes.slice(0, MAX_BROWSER_SPATIAL_CHILDREN);
        const spatialMarkup = spatialVisible
          .map((node) => this.renderSpatialBrowserNode(modelId, index, node, 0))
          .join('');
        const spatialMoreMarkup = spatialRootNodes.length > MAX_BROWSER_SPATIAL_CHILDREN
          ? `<div class="browser-leaf"><span class="material-icons-round">more_horiz</span><span>${spatialRootNodes.length - MAX_BROWSER_SPATIAL_CHILDREN} more nodes</span><span class="browser-count">+</span></div>`
          : '';

        return `
          <details class="browser-node is-model" open>
            <summary class="browser-summary">
              <span class="browser-twist material-icons-round">chevron_right</span>
              <button
                type="button"
                class="browser-action"
                data-browser-action="select-model"
                data-model-id="${escapedModelId}"
                title="Select full model"
              >
                ${escapeHtml(record.fileName)}${visibilitySuffix}
              </button>
              <span class="browser-count">${record.elementCount}</span>
            </summary>
            <div class="browser-children">
              <div class="browser-leaf">
                <span class="material-icons-round">my_location</span>
                <button
                  type="button"
                  class="browser-action"
                  data-browser-action="fit-model"
                  data-model-id="${escapedModelId}"
                  title="Fit camera to model"
                >
                  Default
                </button>
                <span class="browser-count">${levelEntries.length} lvls</span>
              </div>

              <details class="browser-node" ${levelEntries.length > 0 ? 'open' : ''}>
                <summary class="browser-summary">
                  <span class="browser-twist material-icons-round">chevron_right</span>
                  <span>Levels</span>
                  <span class="browser-count">${levelEntries.length}</span>
                </summary>
                <div class="browser-children">
                  ${levelMarkup || '<div class="browser-leaf"><span class="material-icons-round">folder_open</span><span>No levels detected</span><span class="browser-count">-</span></div>'}
                  ${levelMoreMarkup}
                </div>
              </details>

              <details class="browser-node">
                <summary class="browser-summary">
                  <span class="browser-twist material-icons-round">chevron_right</span>
                  <span>Spatial Structure</span>
                  <span class="browser-count">${spatialRootNodes.length}</span>
                </summary>
                <div class="browser-children">
                  ${spatialMarkup || '<div class="browser-leaf"><span class="material-icons-round">account_tree</span><span>No spatial tree data</span><span class="browser-count">-</span></div>'}
                  ${spatialMoreMarkup}
                </div>
              </details>
            </div>
          </details>
        `;
      })
      .join('');

    this.dom.modelBrowserTree.innerHTML = modelMarkup;
  }

  private renderFederatedTree(): void {
    if (this.federatedModels.size === 0) {
      this.dom.federationTree.innerHTML = '<div class="tree-item">No models loaded yet</div>';
      return;
    }

    const cards = [...this.federatedModels.values()]
      .map((record) => {
        const modelId = String(record.modelId);
        const escapedModelId = escapeHtml(modelId);
        const opacityPct = Math.round(this.clamp(record.opacity, 0, 1) * 100);
        const visibilityLabel = record.visible ? 'Hide' : 'Show';
        const visibilityStateClass = record.visible ? '' : 'is-off';
        const gizmoStateClass = this.activeGizmoModelId === modelId ? 'is-active' : '';
        const levels = this.modelIndices.get(record.modelId)?.levels;
        const levelEntries = levels ? [...levels.entries()] : [];
        const levelMarkup = levelEntries.length === 0
          ? '<div class="federated-level-empty">No storeys found</div>'
          : levelEntries
            .slice(0, 80)
            .map(([levelName, ids]) => `
              <button
                class="federated-level-btn"
                type="button"
                data-model-id="${escapedModelId}"
                data-level="${escapeHtml(levelName)}"
                title="Isolate level ${escapeHtml(levelName)}"
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
                  title="Select full model"
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
                ${escapedModelId} | ${record.elementCount} elements | ${this.formatModelSize(record.sizeBytes)}
              </div>
            </div>

            <div class="federated-opacity">
              <div class="federated-opacity-head">
                <span>Opacity</span>
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

            <div class="federated-transform-title">Offset XYZ</div>
            <div class="federated-transform-grid">
              <label>X<input type="number" step="0.1" value="${record.offsetPosition.x.toFixed(2)}" data-model-id="${escapedModelId}" data-transform="px" /></label>
              <label>Y<input type="number" step="0.1" value="${record.offsetPosition.y.toFixed(2)}" data-model-id="${escapedModelId}" data-transform="py" /></label>
              <label>Z<input type="number" step="0.1" value="${record.offsetPosition.z.toFixed(2)}" data-model-id="${escapedModelId}" data-transform="pz" /></label>
            </div>

            <div class="federated-transform-title">Rotation XYZ (deg)</div>
            <div class="federated-transform-grid">
              <label>Rx<input type="number" step="1" value="${record.offsetRotation.x.toFixed(1)}" data-model-id="${escapedModelId}" data-transform="rx" /></label>
              <label>Ry<input type="number" step="1" value="${record.offsetRotation.y.toFixed(1)}" data-model-id="${escapedModelId}" data-transform="ry" /></label>
              <label>Rz<input type="number" step="1" value="${record.offsetRotation.z.toFixed(1)}" data-model-id="${escapedModelId}" data-transform="rz" /></label>
            </div>

            <div class="federated-model-actions">
              <button class="federated-model-btn" type="button" data-model-id="${escapedModelId}" data-model-action="select-model">Select</button>
              <button class="federated-model-btn ${gizmoStateClass}" type="button" data-model-id="${escapedModelId}" data-model-action="toggle-gizmo">Gizmo</button>
              <button class="federated-model-btn" type="button" data-model-id="${escapedModelId}" data-model-action="fit">Fit</button>
              <button class="federated-model-btn" type="button" data-model-id="${escapedModelId}" data-model-action="reset">Reset</button>
            </div>

            <div class="federated-levels">
              <details>
                <summary>Levels (${levelEntries.length})</summary>
                <div class="federated-level-list">${levelMarkup}</div>
              </details>
            </div>
          </div>
        `;
      })
      .join('');

    this.dom.federationTree.innerHTML = cards;
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

  private updateBackgroundPresetState(): void {
    for (const button of this.dom.backgroundPresetButtons) {
      const preset = this.normalizeHexColor(button.dataset.bgPreset, '');
      button.classList.toggle('active', preset === this.backgroundColor);
    }
  }

  private setGridVisible(visible: boolean, updateStatus: boolean): void {
    this.gridVisible = visible;
    if (this.gridHelper) this.gridHelper.visible = visible;
    this.dom.toggleGrid.checked = visible;
    if (updateStatus) this.setStatus(visible ? 'Grid enabled' : 'Grid hidden');
  }

  private setBackgroundColor(color: string, updateStatus: boolean): void {
    const normalized = this.normalizeHexColor(color);
    this.backgroundColor = normalized;
    const threeColor = new THREE.Color(normalized);
    if (this.world?.scene?.three) this.world.scene.three.background = threeColor;
    // Sync the renderer clear color so PEN style (which bypasses the scene color pass)
    // uses the user-selected background
    const renderer = this.world?.renderer?.three;
    if (renderer) {
      renderer.setClearColor(threeColor, 1);
    }
    // Also sync the postproduction basePass clear color
    const postRenderer = this.getPostproductionRenderer();
    const post = postRenderer?.postproduction;
    if (post?.basePass) {
      post.basePass.clearColor = threeColor;
      post.basePass.clearAlpha = 1;
    }
    this.dom.backgroundColorInput.value = normalized;
    this.updateBackgroundPresetState();
    if (updateStatus) this.setStatus(`Background color set to ${normalized}`);
  }

  private syncVisualSettingsUi(): void {
    this.dom.toggleGrid.checked = this.gridVisible;
    this.dom.visualStyleSelect.value = this.visualStyle;
    this.dom.backgroundColorInput.value = this.backgroundColor;
    this.updateBackgroundPresetState();
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
        return 'Basic';
      case 'pen':
        return 'Pen';
      case 'color-pen':
        return 'Color Pen';
      case 'color-shadows':
        return 'Color Shadows';
      case 'color-pen-shadows':
        return 'Color Pen Shadows';
      default:
        return 'Color Pen Shadows';
    }
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

    // Sync the postproduction basePass clear color with the user's background
    if (post.basePass) {
      post.basePass.clearColor = new THREE.Color(this.backgroundColor);
      post.basePass.clearAlpha = 1;
      post.basePass.clearDepth = true;
    }

    // Map directly to ThatOpen PostproductionAspect enum
    switch (style) {
      case 'basic':
        post.style = OBCF.PostproductionAspect.COLOR;
        break;
      case 'pen':
        post.style = OBCF.PostproductionAspect.PEN;
        post.outlinesEnabled = true;
        break;
      case 'color-pen':
        post.style = OBCF.PostproductionAspect.COLOR_PEN;
        post.outlinesEnabled = true;
        break;
      case 'color-shadows':
        post.style = OBCF.PostproductionAspect.COLOR_SHADOWS;
        post.glossEnabled = true;
        break;
      case 'color-pen-shadows':
        post.style = OBCF.PostproductionAspect.COLOR_PEN_SHADOWS;
        post.outlinesEnabled = true;
        post.glossEnabled = true;
        break;
      default:
        post.style = OBCF.PostproductionAspect.COLOR_PEN_SHADOWS;
        post.outlinesEnabled = true;
        post.glossEnabled = true;
        break;
    }
  }

  private async resetModelColors(): Promise<void> {
    if (!this.hiddenLineColorOverride) return;
    const tasks: Promise<unknown>[] = [];
    for (const model of this.federatedModels.values()) {
      const fragmentsModel = this.getFragmentsModel(model.modelId);
      if (!fragmentsModel || typeof fragmentsModel?.resetColor !== 'function') continue;
      tasks.push(fragmentsModel.resetColor(undefined));
    }
    if (tasks.length > 0) await Promise.allSettled(tasks);
    this.hiddenLineColorOverride = false;
  }

  private async applyHiddenLineColors(): Promise<void> {
    const tasks: Promise<unknown>[] = [];
    const white = new THREE.Color(0xffffff);
    for (const model of this.federatedModels.values()) {
      const fragmentsModel = this.getFragmentsModel(model.modelId);
      if (!fragmentsModel || typeof fragmentsModel?.setColor !== 'function') continue;
      tasks.push(fragmentsModel.setColor(undefined, white));
    }
    if (tasks.length > 0) await Promise.allSettled(tasks);
    this.hiddenLineColorOverride = tasks.length > 0;
  }

  private applyConsistentLighting(): void {
    const scene = this.world?.scene as OBC.SimpleScene | undefined;
    if (!scene?.three) return;

    // Save original light intensities
    this.savedLightStates = [];

    // Disable directional lights (remove directional shading / shadows)
    if (scene.directionalLights) {
      for (const [, light] of scene.directionalLights) {
        this.savedLightStates.push({ light, visible: light.visible, intensity: light.intensity });
        light.intensity = 0;
      }
    }

    // Boost ambient lights for flat, even illumination preserving original colors
    if (scene.ambientLights) {
      for (const [, light] of scene.ambientLights) {
        this.savedLightStates.push({ light, visible: light.visible, intensity: light.intensity });
        light.intensity = 2.0;
      }
    }

    this.consistentLightOverride = true;
  }

  private restoreOriginalLighting(): void {
    if (!this.consistentLightOverride) return;

    // Restore all lights to their original intensities
    for (const state of this.savedLightStates) {
      state.light.visible = state.visible;
      state.light.intensity = state.intensity;
    }
    this.savedLightStates = [];
    this.consistentLightOverride = false;
  }

  private async setVisualStyle(style: VisualStyle, updateStatus: boolean, persist: boolean): Promise<void> {
    const resolvedStyle = this.parseVisualStyle(style);
    this.visualStyle = resolvedStyle;
    this.dom.visualStyleSelect.value = resolvedStyle;

    await this.resetModelColors();
    this.restoreOriginalLighting();
    this.configurePostproduction(resolvedStyle);

    // Reset independent toggles when switching styles
    this.xrayEnabled = false;
    this.edgesEnabled = false;
    this.dom.btnTransparency.classList.toggle('active', false);
    this.dom.btnWireframe.classList.toggle('active', false);
    this.applyXRay();
    this.applyEdges();

    await this.fragments.core.update(true);

    if (persist) this.persistLocalState();
    if (updateStatus) this.setStatus(`Visual style: ${this.getVisualStyleLabel(resolvedStyle)}`);
  }

  private formatModelSize(sizeBytes: number): string {
    if (!sizeBytes || sizeBytes <= 0) return '-';
    return `${(sizeBytes / 1024 / 1024).toFixed(1)} MB`;
  }

  private clamp(value: number, min: number, max: number): number {
    return Math.min(max, Math.max(min, value));
  }

  private registerModelAlias(alias: string, modelId: string): void {
    const normalizedAlias = alias.trim();
    const normalizedModelId = modelId.trim();
    if (!normalizedAlias || !normalizedModelId) return;
    this.modelIdAliases.set(normalizedAlias, normalizedModelId);
    this.modelIdAliases.set(normalizedAlias.toLowerCase(), normalizedModelId);
  }

  private getModelInternalId(model: any, fallback = ''): string {
    const fromModel = String(model?.modelId ?? '').trim();
    if (fromModel) return fromModel;
    return fallback.trim();
  }

  private findFragmentsEntryByInternalId(modelId: string): [string, any] | null {
    if (!this.fragments?.list) return null;
    const normalizedId = modelId.trim();
    if (!normalizedId) return null;
    for (const [key, entry] of this.fragments.list.entries()) {
      const internalId = this.getModelInternalId(entry, key);
      if (internalId === normalizedId) return [String(key), entry];
    }
    return null;
  }

  private getFragmentsModel(modelIdOrAlias: string): any | null {
    if (!this.fragments?.list) return null;
    const resolvedModelId = this.resolveModelId(modelIdOrAlias);
    if (!resolvedModelId) return null;

    const direct = this.fragments.list.get(resolvedModelId);
    if (direct) return direct;

    const byInternal = this.findFragmentsEntryByInternalId(resolvedModelId);
    if (byInternal) return byInternal[1];
    return null;
  }

  private resolveModelId(candidate: string): string | null {
    if (!this.fragments?.list) return null;
    const value = candidate.trim();
    if (!value) return null;

    const directByKey = this.fragments?.list?.get(value);
    if (directByKey) {
      const internalId = this.getModelInternalId(directByKey, value);
      this.registerModelAlias(value, internalId);
      this.registerModelAlias(internalId, internalId);
      return internalId;
    }

    const aliasResolved = this.modelIdAliases.get(value) ?? this.modelIdAliases.get(value.toLowerCase());
    if (aliasResolved) {
      const aliasByKey = this.fragments.list.get(aliasResolved);
      if (aliasByKey) {
        const internalId = this.getModelInternalId(aliasByKey, aliasResolved);
        this.registerModelAlias(value, internalId);
        this.registerModelAlias(aliasResolved, internalId);
        return internalId;
      }
      const aliasByInternal = this.findFragmentsEntryByInternalId(aliasResolved);
      if (aliasByInternal) {
        const [entryKey, entry] = aliasByInternal;
        const internalId = this.getModelInternalId(entry, entryKey);
        this.registerModelAlias(value, internalId);
        this.registerModelAlias(entryKey, internalId);
        return internalId;
      }
    }

    const lowerValue = value.toLowerCase();
    for (const [entryKey, entry] of this.fragments.list.entries()) {
      const internalId = this.getModelInternalId(entry, String(entryKey));
      if (internalId.toLowerCase() === lowerValue || String(entryKey).toLowerCase() === lowerValue) {
        this.registerModelAlias(value, internalId);
        this.registerModelAlias(String(entryKey), internalId);
        return internalId;
      }
    }

    for (const [modelId, record] of this.federatedModels.entries()) {
      if (record.fileName.toLowerCase() === lowerValue) {
        this.registerModelAlias(value, modelId);
        return modelId;
      }
    }

    return null;
  }

  private fireAndForget(task: Promise<unknown>, context: string): void {
    task.catch((error) => this.handleAsyncError(context, error));
  }

  private handleAsyncError(context: string, error: unknown): void {
    const message = serializeError(error);
    const normalized = message.toLowerCase();
    if (normalized.includes('model not found')) {
      this.pruneSelectedItems();
      if (context !== 'Camera update') {
        this.setStatus('Model synchronization updated. Please reselect element if needed.');
      }
      return;
    } else {
      this.setStatus(`${context} failed: ${message}`);
    }
    console.error(error);
  }

  private isLoadedModelId(modelId: string): boolean {
    return this.resolveModelId(modelId) !== null;
  }

  private getValidModelIdMap(input: OBC.ModelIdMap): OBC.ModelIdMap {
    const valid: OBC.ModelIdMap = {};
    for (const [modelId, ids] of Object.entries(input)) {
      if (ids.size === 0) continue;
      const resolvedModelId = this.resolveModelId(modelId);
      if (!resolvedModelId) continue;
      if (!valid[resolvedModelId]) valid[resolvedModelId] = new Set<number>();
      for (const localId of ids) valid[resolvedModelId].add(localId);
    }
    return valid;
  }

  private pruneSelectedItems(): void {
    const valid = this.getValidModelIdMap(this.selectedItems);
    clearMap(this.selectedItems);
    Object.assign(this.selectedItems, valid);
  }

  private async selectWholeModel(modelId: string): Promise<void> {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const index = this.modelIndices.get(resolvedModelId);
    if (!index || index.allIds.size === 0) {
      this.setStatus('Model index not ready yet');
      return;
    }
    clearMap(this.selectedItems);
    this.selectedItems[resolvedModelId] = new Set(index.allIds);
    await this.refreshSelectionVisuals();
    await this.zoomToItems(this.selectedItems);
    this.setStatus(`Selected full model (${index.allIds.size} elements)`);
  }

  private toggleModelVisibility(modelId: string): void {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const model = this.federatedModels.get(resolvedModelId);
    if (!model) return;
    model.visible = !model.visible;
    model.object.visible = model.visible;
    if (!model.visible && this.activeGizmoModelId === resolvedModelId) this.detachModelGizmo();
    this.applyXRay();
    if (this.edgesEnabled) this.applyEdges();
    this.fireAndForget(this.fragments.core.update(true), 'Toggle visibility');
    this.fireAndForget(this.updateVisibilityCount(), 'Update visibility');
    this.renderModelBrowser();
    this.renderFederatedTree();
    this.setStatus(`${model.visible ? 'Shown' : 'Hidden'}: ${model.fileName}`);
  }

  private applyModelOpacity(modelId: string, opacity: number): void {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const model = this.federatedModels.get(resolvedModelId);
    if (!model) return;
    model.opacity = this.clamp(opacity, 0, 1);
    this.applyXRay();
    if (this.edgesEnabled) this.applyEdges();
    this.fireAndForget(this.fragments.core.update(true), 'Adjust opacity');
  }

  private toggleModelGizmo(modelId: string): void {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const model = this.federatedModels.get(resolvedModelId);
    if (!model || !this.transformControls) return;
    if (!model.visible) {
      this.setStatus('Show the model before enabling gizmo');
      return;
    }

    if (this.activeGizmoModelId === resolvedModelId && this.transformControlsHelper?.visible) {
      this.detachModelGizmo();
      this.setStatus('Model gizmo detached');
      return;
    }

    this.activeGizmoModelId = resolvedModelId;
    this.transformControls.attach(model.object);
    this.transformControls.setMode('translate');
    this.transformControls.enabled = true;
    if (this.transformControlsHelper) this.transformControlsHelper.visible = true;
    this.renderFederatedTree();
    this.setStatus(`Model gizmo active: ${model.fileName} (W move, E rotate, R reset transform)`);
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

    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const model = this.federatedModels.get(resolvedModelId);
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
    this.setStatus(`Updated transform: ${model.fileName}`);
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
    this.fireAndForget(this.fragments.core.update(true), 'Apply transform');
  }

  private resetModelOffsets(modelId: string): void {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const model = this.federatedModels.get(resolvedModelId);
    if (!model) return;
    model.offsetPosition = { x: 0, y: 0, z: 0 };
    model.offsetRotation = { x: 0, y: 0, z: 0 };
    this.applyModelTransform(model);
    this.renderModelBrowser();
    this.renderFederatedTree();
    this.setStatus(`Reset transform: ${model.fileName}`);
  }

  private fitToModelById(modelId: string): void {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const model = this.federatedModels.get(resolvedModelId);
    if (!model) return;
    const bbox = new THREE.Box3().setFromObject(model.object);
    if (bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(
      center.x + maxDim,
      center.y + maxDim * 0.75,
      center.z + maxDim,
      center.x,
      center.y,
      center.z,
      true,
    );
  }

  private async isolateLevelForModel(modelId: string, level: string): Promise<void> {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const index = this.modelIndices.get(resolvedModelId);
    const ids = index?.levels.get(level);
    if (!ids || ids.size === 0) {
      this.setStatus(`No elements found for level ${level}`);
      return;
    }
    const map = this.getValidModelIdMap({ [resolvedModelId]: new Set(ids) });
    if (isMapEmpty(map)) {
      this.setStatus(`Level ${level} is not available for current loaded model IDs`);
      return;
    }
    await this.hider.isolate(map);
    await this.updateVisibilityCount();
    this.setStatus(`Isolated ${level}`);
  }

  private async isolateClassForModelLevel(modelId: string, level: string, className: string): Promise<void> {
    const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
    const ids = this.getClassIdsForModelLevel(resolvedModelId, level, className);
    if (ids.size === 0) {
      this.setStatus(`No ${className} elements found in ${level}`);
      return;
    }

    const map = this.getValidModelIdMap({ [resolvedModelId]: ids });
    if (isMapEmpty(map)) {
      this.setStatus(`${className} in ${level} is not available for current loaded model IDs`);
      return;
    }
    await this.hider.isolate(map);
    await this.updateVisibilityCount();
    this.setStatus(`Isolated ${className} in ${level}`);
  }

  private collectCheckedValues(container: HTMLElement): string[] {
    const values: string[] = [];
    container.querySelectorAll<HTMLInputElement>('input[type="checkbox"]:checked').forEach((checkbox) => {
      values.push(checkbox.value);
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
      this.setStatus('No filters selected. Showing all elements');
      return;
    }

    const classMap = this.mapFromClassFilters(selectedClasses);
    const levelMap = this.mapFromLevelFilters(selectedLevels);

    let effectiveMap: OBC.ModelIdMap;
    if (!isMapEmpty(classMap) && !isMapEmpty(levelMap)) effectiveMap = intersectMaps(classMap, levelMap);
    else if (!isMapEmpty(classMap)) effectiveMap = classMap;
    else effectiveMap = levelMap;
    effectiveMap = this.getValidModelIdMap(effectiveMap);

    await this.hider.set(false);
    if (!isMapEmpty(effectiveMap)) await this.hider.set(true, effectiveMap);
    await this.updateVisibilityCount();
    this.setStatus('Filters applied');
  }

  private async searchElements(term: string): Promise<void> {
    if (!term) {
      this.dom.elementResults.innerHTML = '';
      this.setStatus('Enter a term to search');
      return;
    }

    const escaped = escapeRegExp(term);
    const pattern = new RegExp(escaped, 'i');
    const results: SearchResult[] = [];

    for (const [modelKey, model] of this.fragments.list.entries()) {
      const modelId = this.getModelInternalId(model, String(modelKey));
      this.registerModelAlias(String(modelKey), modelId);
      this.registerModelAlias(modelId, modelId);
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
        const data = itemsData[i] as Record<string, unknown>;
        const name = (data?.Name as string | undefined) || `Element ${localId}`;
        const type = (data?.ObjectType as string | undefined) || (data?.PredefinedType as string | undefined) || 'Item';
        const globalId = (data?.GlobalId as string | undefined) || '-';
        results.push({ modelId, localId, name, type, globalId });
      }
    }

    const capped = results.slice(0, 180);
    this.dom.elementResults.innerHTML = capped
      .map((result) => `
        <div class="result-item" data-model-id="${result.modelId}" data-local-id="${result.localId}">
          <div><strong>${result.name}</strong></div>
          <div>${result.type}</div>
          <div>${result.globalId}</div>
        </div>
      `)
      .join('');

    this.dom.elementResults.querySelectorAll<HTMLElement>('.result-item').forEach((item) => {
      item.addEventListener('click', () => {
        const modelId = item.dataset.modelId;
        const localId = Number(item.dataset.localId);
        if (!modelId || Number.isNaN(localId)) return;
        this.fireAndForget(this.selectSingleItem(modelId, localId, true), 'Select search result');
      });
    });

    this.setStatus(`Search found ${results.length} matches`);
  }

  private async loadIfcFiles(files: File[]): Promise<void> {
    const ifcFiles = files.filter((file) => file.name.toLowerCase().endsWith('.ifc'));
    if (ifcFiles.length === 0) {
      this.setStatus('Only IFC files are supported');
      return;
    }
    if (this.isModelLoading) {
      this.setStatus('A model is already loading. Please wait...');
      return;
    }

    const failedFiles: string[] = [];
    const batchTotal = ifcFiles.length;
    if (batchTotal > 1) this.suppressAutoFit = true;

    for (let i = 0; i < ifcFiles.length; i += 1) {
      const file = ifcFiles[i];
      const success = await this.loadIfcFile(file, i + 1, batchTotal);
      if (!success) failedFiles.push(file.name);
    }

    this.suppressAutoFit = false;

    if (batchTotal > 1 && this.modelObjects.length > 0) this.fitToModel();

    if (failedFiles.length === 0) {
      this.setStatus(batchTotal > 1 ? `Loaded ${batchTotal} IFC models` : 'Model loaded successfully');
      return;
    }

    if (failedFiles.length === batchTotal) {
      this.setStatus('Failed to load selected IFC files');
      return;
    }
    this.setStatus(`Loaded ${batchTotal - failedFiles.length}/${batchTotal} IFC files`);
  }

  private async loadIfcFile(file: File, batchIndex = 1, batchTotal = 1): Promise<boolean> {
    if (this.isModelLoading) {
      this.setStatus('A model is already loading. Please wait...');
      return false;
    }

    this.isModelLoading = true;
    const requestId = ++this.loadRequestId;
    this.dom.btnUpload.disabled = true;
    this.dom.btnUploadEmpty.disabled = true;
    this.dom.fileInput.disabled = true;

    this.dom.emptyState.hidden = true;
    this.dom.loadingOverlay.hidden = false;
    this.dom.loadingText.textContent = batchTotal > 1
      ? `Reading IFC file ${batchIndex}/${batchTotal}...`
      : 'Reading IFC file...';
    this.dom.loadingProgress.style.width = '8%';
    this.dom.modelName.textContent = batchTotal > 1
      ? `Loading ${batchIndex}/${batchTotal}: ${file.name}`
      : file.name;
    this.setStatus('Loading model...');

    let timeoutHandle: number | undefined;

    try {
      const start = performance.now();
      const buffer = await file.arrayBuffer();
      const data = new Uint8Array(buffer);
      this.pendingModelMetaQueue.push({ fileName: file.name, sizeBytes: file.size });

      if (requestId === this.loadRequestId) {
        this.dom.loadingText.textContent = 'Converting IFC to fragments...';
        this.dom.loadingProgress.style.width = '25%';
      }

      const loadPromise = this.ifcLoader.load(data, true, file.name, {
        instanceCallback: (importer: any) => {
          if (typeof importer?.addAllAttributes === 'function') importer.addAllAttributes();
          if (typeof importer?.addAllRelations === 'function') importer.addAllRelations();
        },
        processData: {
          progressCallback: (progress: number) => {
            if (requestId !== this.loadRequestId) return;
            const percentage = Math.round(25 + progress * 70);
            this.dom.loadingProgress.style.width = `${percentage}%`;
            this.dom.loadingText.textContent = `Processing ${Math.round(progress * 100)}%`;
          },
        },
      });

      const timeoutPromise = new Promise<never>((_resolve, reject) => {
        timeoutHandle = window.setTimeout(() => {
          reject(new Error('Model loading timed out. Please try again with a smaller IFC or reload the page.'));
        }, 120000);
      });

      const loadedModel = await Promise.race([loadPromise, timeoutPromise]) as any;
      await this.ensureModelRegistered(loadedModel);

      if (timeoutHandle !== undefined) {
        window.clearTimeout(timeoutHandle);
        timeoutHandle = undefined;
      }

      // Workaround for web-ifc #1833: extract real material colors and apply them
      // TEMPORARILY DISABLED — testing native fragment colors without override
      // if (requestId === this.loadRequestId) {
      //   this.dom.loadingText.textContent = 'Applying material colors...';
      //   try {
      //     await this.extractAndApplyMaterialColors(data, loadedModel);
      //   } catch (colorErr) {
      //     console.warn('[Viewer] Material color extraction failed (non-fatal):', colorErr);
      //   }
      // }

      const elapsed = ((performance.now() - start) / 1000).toFixed(1);
      if (requestId !== this.loadRequestId) return false;

      this.dom.perfInfo.textContent = `Loaded in ${elapsed}s | ${(file.size / 1024 / 1024).toFixed(1)}MB`;
      this.dom.loadingProgress.style.width = '100%';
      this.setStatus('Model loaded successfully');

      setTimeout(() => {
        if (requestId === this.loadRequestId) {
          this.dom.loadingOverlay.hidden = true;
        }
      }, 220);
      return true;
    } catch (error) {
      const staleIndex = this.pendingModelMetaQueue.findIndex(
        (entry) => entry.fileName === file.name && entry.sizeBytes === file.size,
      );
      if (staleIndex !== -1) this.pendingModelMetaQueue.splice(staleIndex, 1);

      if (timeoutHandle !== undefined) {
        window.clearTimeout(timeoutHandle);
      }
      if (requestId !== this.loadRequestId) return false;

      this.dom.loadingOverlay.hidden = true;
      this.dom.emptyState.hidden = this.modelObjects.length > 0;
      this.setStatus(`Failed to load IFC: ${serializeError(error)}`);
      console.error(error);
      return false;
    } finally {
      if (requestId === this.loadRequestId) {
        this.isModelLoading = false;
        this.dom.btnUpload.disabled = false;
        this.dom.btnUploadEmpty.disabled = false;
        this.dom.fileInput.disabled = false;
      }
    }
    return false;
  }

  /**
   * Workaround for web-ifc issue #1833:
   * web-ifc 0.0.74+ ignores IfcMaterialDefinitionRepresentation colors and only
   * uses geometry-level IfcStyledItem colors (which default to gray in Revit exports).
   * This method independently parses the IFC data to extract the real material colors
   * and applies them to the loaded fragment model.
   */
  private async extractAndApplyMaterialColors(ifcData: Uint8Array, model: any): Promise<void> {
    const ifcApi = new WebIFC.IfcAPI();
    // Load WASM from public folder (web-ifc.wasm is served by Vite with correct MIME type)
    ifcApi.SetWasmPath('/', true);
    await ifcApi.Init();
    const modelId = ifcApi.OpenModel(ifcData);

    try {
      // IFC type constants
      const IFCINDEXEDCOLOURMAP = 3570813810;
      const IFCCOLOURRGBLIST = 3285139300;
      const IFCPRODUCTDEFINITIONSHAPE = 673634403;
      const IFCSHAPEREPRESENTATION = 4240577450;
      const IFCMAPPEDITEM = 2347385850;
      const IFCREPRESENTATIONMAP = 1660063152;

      const expressIdToColor = new Map<number, THREE.Color>();

      // ──────────────────────────────────────────────────────────────────────
      // Strategy A: IFC4 Reference View — IfcIndexedColourMap + IfcColourRgbList
      // Used by Revit IFC4 exports with tessellated geometry (IfcPolygonalFaceSet)
      // ──────────────────────────────────────────────────────────────────────
      const icmIds = ifcApi.GetLineIDsWithType(modelId, IFCINDEXEDCOLOURMAP);
      console.log(`[Viewer] Found ${icmIds.size()} IfcIndexedColourMap entities`);

      if (icmIds.size() > 0) {
        // Build faceSetId → primary color map from IfcIndexedColourMap
        const faceSetToColor = new Map<number, THREE.Color>();

        for (let i = 0; i < icmIds.size(); i++) {
          try {
            const icm = ifcApi.GetLine(modelId, icmIds.get(i), false);
            if (!icm) continue;

            // MappedTo → IfcPolygonalFaceSet expressId
            const mappedTo = icm.MappedTo?.value ?? icm.MappedTo;
            if (typeof mappedTo !== 'number') continue;

            // Colours → IfcColourRgbList expressId
            const coloursRef = icm.Colours?.value ?? icm.Colours;
            if (typeof coloursRef !== 'number') continue;

            const colourList = ifcApi.GetLine(modelId, coloursRef, false);
            if (!colourList?.ColourList || !Array.isArray(colourList.ColourList) || colourList.ColourList.length === 0) continue;

            // Extract the first colour from the list (primary colour for this face set)
            const firstColour = colourList.ColourList[0];
            if (!Array.isArray(firstColour) || firstColour.length < 3) continue;

            const r = firstColour[0]?._representationValue ?? firstColour[0]?.value ?? firstColour[0] ?? 0.5;
            const g = firstColour[1]?._representationValue ?? firstColour[1]?.value ?? firstColour[1] ?? 0.5;
            const b = firstColour[2]?._representationValue ?? firstColour[2]?.value ?? firstColour[2] ?? 0.5;

            const color = new THREE.Color(
              typeof r === 'number' ? r : 0.5,
              typeof g === 'number' ? g : 0.5,
              typeof b === 'number' ? b : 0.5,
            );
            faceSetToColor.set(mappedTo, color);
          } catch (e) {
            // Skip
          }
        }

        console.log(`[Viewer] Built faceSet→color map with ${faceSetToColor.size} entries`);

        // Now iterate ALL products to find their face sets and look up colors
        // Instead of hardcoding product type constants (which can be incorrect across web-ifc versions),
        // find all products via IfcRelContainedInSpatialStructure traversal
        const IFCRELCONTAINEDINSPATIALSTRUCTURE = 3242617779;

        const allProductIds = new Set<number>();

        // Collect products from IfcRelContainedInSpatialStructure
        try {
          const relContained = ifcApi.GetLineIDsWithType(modelId, IFCRELCONTAINEDINSPATIALSTRUCTURE);
          for (let i = 0; i < relContained.size(); i++) {
            const rel = ifcApi.GetLine(modelId, relContained.get(i), false);
            if (rel?.RelatedElements) {
              const elems = Array.isArray(rel.RelatedElements) ? rel.RelatedElements : [rel.RelatedElements];
              for (const e of elems) {
                const eId = e?.value ?? e;
                if (typeof eId === 'number') allProductIds.add(eId);
              }
            }
          }
        } catch { /* skip */ }

        console.log(`[Viewer] Found ${allProductIds.size} products via spatial containment`);

        for (const prodExpressId of allProductIds) {
          try {
            const prod = ifcApi.GetLine(modelId, prodExpressId, false);
            if (!prod?.Representation) continue;

            const repRef = prod.Representation?.value ?? prod.Representation;
            if (typeof repRef !== 'number') continue;

            const prodDefShape = ifcApi.GetLine(modelId, repRef, false);
            if (!prodDefShape?.Representations) continue;

            const reps = Array.isArray(prodDefShape.Representations)
              ? prodDefShape.Representations : [prodDefShape.Representations];

            let foundColor: THREE.Color | null = null;

            for (const shapeRepRef of reps) {
              if (foundColor) break;
              const shapeRepId = shapeRepRef?.value ?? shapeRepRef;
              if (typeof shapeRepId !== 'number') continue;

              const shapeRep = ifcApi.GetLine(modelId, shapeRepId, false);
              if (!shapeRep?.Items) continue;

              const items = Array.isArray(shapeRep.Items) ? shapeRep.Items : [shapeRep.Items];

              for (const itemRef of items) {
                if (foundColor) break;
                const itemId = itemRef?.value ?? itemRef;
                if (typeof itemId !== 'number') continue;

                // Check if this item is directly a face set with a color
                if (faceSetToColor.has(itemId)) {
                  foundColor = faceSetToColor.get(itemId)!;
                  break;
                }

                // It might be an IfcMappedItem → IfcRepresentationMap → MappedRepresentation → Items
                try {
                  const item = ifcApi.GetLine(modelId, itemId, false);
                  if (item?.type === IFCMAPPEDITEM) {
                    const mapRef = item.MappingSource?.value ?? item.MappingSource;
                    if (typeof mapRef === 'number') {
                      const repMap = ifcApi.GetLine(modelId, mapRef, false);
                      const mappedRepRef = repMap?.MappedRepresentation?.value ?? repMap?.MappedRepresentation;
                      if (typeof mappedRepRef === 'number') {
                        const mappedRep = ifcApi.GetLine(modelId, mappedRepRef, false);
                        if (mappedRep?.Items) {
                          const innerItems = Array.isArray(mappedRep.Items) ? mappedRep.Items : [mappedRep.Items];
                          for (const innerRef of innerItems) {
                            const innerId = innerRef?.value ?? innerRef;
                            if (typeof innerId === 'number' && faceSetToColor.has(innerId)) {
                              foundColor = faceSetToColor.get(innerId)!;
                              break;
                            }
                          }
                        }
                      }
                    }
                  }
                } catch {
                  // Skip
                }
              }
            }

            if (foundColor) {
              expressIdToColor.set(prodExpressId, foundColor);
            }
          } catch (e) {
            // Skip problematic products
          }
        }

        console.log(`[Viewer] Mapped ${expressIdToColor.size} products to IfcIndexedColourMap colors`);
      }

      // ──────────────────────────────────────────────────────────────────────
      // Strategy B: Traditional IfcMaterialDefinitionRepresentation path
      // Used by IFC2x3 and some IFC4 exports with IfcStyledItem
      // Only runs if Strategy A found no colors
      // ──────────────────────────────────────────────────────────────────────
      if (expressIdToColor.size === 0) {
        const IFCRELASSOCIATESMATERIAL = 2655215786;
        const IFCMATERIAL = 1838606355;
        const IFCMATERIALLAYERSETUSAGE = 1303795690;
        const IFCMATERIALLAYERSET = 3303938423;
        const IFCMATERIALLAYER = 248100487;
        const IFCMATERIALCONSTITUENTSET = 3079605661;
        const IFCMATERIALPROFILESET = 164193824;
        const IFCMATERIALDEFINITIONREPRESENTATION = 2022407955;

        const materialColorCache = new Map<number, THREE.Color>();

        const extractColorFromSurfaceStyle = (styleExpressId: number): THREE.Color | null => {
          try {
            const style = ifcApi.GetLine(modelId, styleExpressId, true);
            if (!style) return null;

            const styles = style.Styles;
            if (styles) {
              const styleEntries = Array.isArray(styles) ? styles : [styles];
              for (const entry of styleEntries) {
                const entryObj = entry?.value ? ifcApi.GetLine(modelId, entry.value, true) : entry;
                if (!entryObj) continue;
                const surfaceColour = entryObj.SurfaceColour;
                if (surfaceColour) {
                  const colourObj = surfaceColour?.value
                    ? ifcApi.GetLine(modelId, surfaceColour.value, true) : surfaceColour;
                  if (colourObj) {
                    const r = colourObj.Red?.value ?? colourObj.Red ?? 0.5;
                    const g = colourObj.Green?.value ?? colourObj.Green ?? 0.5;
                    const b = colourObj.Blue?.value ?? colourObj.Blue ?? 0.5;
                    return new THREE.Color(
                      typeof r === 'number' ? r : 0.5, typeof g === 'number' ? g : 0.5, typeof b === 'number' ? b : 0.5);
                  }
                }
              }
            }
          } catch (e) { /* skip */ }
          return null;
        };

        // Query IfcMaterialDefinitionRepresentation
        const matDefRepIds = ifcApi.GetLineIDsWithType(modelId, IFCMATERIALDEFINITIONREPRESENTATION);
        console.log(`[Viewer] Found ${matDefRepIds.size()} IfcMaterialDefinitionRepresentation entities`);

        for (let i = 0; i < matDefRepIds.size(); i++) {
          try {
            const defRep = ifcApi.GetLine(modelId, matDefRepIds.get(i), false);
            if (!defRep) continue;
            const matId = defRep.RepresentedMaterial?.value ?? defRep.RepresentedMaterial;
            if (typeof matId !== 'number') continue;
            const representations = defRep.Representations;
            if (!representations) continue;
            const repList = Array.isArray(representations) ? representations : [representations];
            for (const repRef of repList) {
              const repId = repRef?.value ?? repRef;
              if (typeof repId !== 'number') continue;
              const styledRep = ifcApi.GetLine(modelId, repId, false);
              if (!styledRep?.Items) continue;
              const items = Array.isArray(styledRep.Items) ? styledRep.Items : [styledRep.Items];
              for (const itemRef of items) {
                const itemId = itemRef?.value ?? itemRef;
                if (typeof itemId !== 'number') continue;
                const styledItem = ifcApi.GetLine(modelId, itemId, false);
                if (!styledItem) continue;
                const stylesProp = styledItem.Styles ?? styledItem.styles;
                if (!stylesProp) continue;
                const stylesList = Array.isArray(stylesProp) ? stylesProp : [stylesProp];
                for (const styleRef of stylesList) {
                  const styleId = styleRef?.value ?? styleRef;
                  if (typeof styleId !== 'number') continue;
                  const psaOrStyle = ifcApi.GetLine(modelId, styleId, false);
                  if (!psaOrStyle) continue;
                  if (psaOrStyle.Styles) {
                    const innerStyles = Array.isArray(psaOrStyle.Styles) ? psaOrStyle.Styles : [psaOrStyle.Styles];
                    for (const innerRef of innerStyles) {
                      const innerId = innerRef?.value ?? innerRef;
                      if (typeof innerId !== 'number') continue;
                      const color = extractColorFromSurfaceStyle(innerId);
                      if (color) { materialColorCache.set(matId, color); break; }
                    }
                  } else {
                    const color = extractColorFromSurfaceStyle(styleId);
                    if (color) materialColorCache.set(matId, color);
                  }
                  if (materialColorCache.has(matId)) break;
                }
                if (materialColorCache.has(matId)) break;
              }
              if (materialColorCache.has(matId)) break;
            }
          } catch (e) { /* skip */ }
        }

        console.log(`[Viewer] Built material color cache with ${materialColorCache.size} entries`);

        const lookupColor = (matId: number): THREE.Color | null => {
          if (materialColorCache.has(matId)) return materialColorCache.get(matId)!;
          try {
            const mat = ifcApi.GetLine(modelId, matId, false);
            if (!mat) return null;
            const typeId = mat.type;
            if (typeId === IFCMATERIALLAYERSETUSAGE) {
              const layerSetRef = mat.ForLayerSet?.value ?? mat.ForLayerSet;
              if (typeof layerSetRef === 'number') {
                const ls = ifcApi.GetLine(modelId, layerSetRef, false);
                if (ls?.MaterialLayers) {
                  for (const lr of (Array.isArray(ls.MaterialLayers) ? ls.MaterialLayers : [ls.MaterialLayers])) {
                    const lid = lr?.value ?? lr; if (typeof lid !== 'number') continue;
                    const l = ifcApi.GetLine(modelId, lid, false);
                    const mid = l?.Material?.value ?? l?.Material;
                    if (typeof mid === 'number' && materialColorCache.has(mid)) return materialColorCache.get(mid)!;
                  }
                }
              }
            }
            if (typeId === IFCMATERIALLAYERSET && mat.MaterialLayers) {
              for (const lr of (Array.isArray(mat.MaterialLayers) ? mat.MaterialLayers : [mat.MaterialLayers])) {
                const lid = lr?.value ?? lr; if (typeof lid !== 'number') continue;
                const l = ifcApi.GetLine(modelId, lid, false);
                const mid = l?.Material?.value ?? l?.Material;
                if (typeof mid === 'number' && materialColorCache.has(mid)) return materialColorCache.get(mid)!;
              }
            }
            if (typeId === IFCMATERIALCONSTITUENTSET && mat.MaterialConstituents) {
              for (const cr of (Array.isArray(mat.MaterialConstituents) ? mat.MaterialConstituents : [mat.MaterialConstituents])) {
                const cid = cr?.value ?? cr; if (typeof cid !== 'number') continue;
                const c = ifcApi.GetLine(modelId, cid, false);
                const mid = c?.Material?.value ?? c?.Material;
                if (typeof mid === 'number' && materialColorCache.has(mid)) return materialColorCache.get(mid)!;
              }
            }
            if (typeId === IFCMATERIALPROFILESET && mat.MaterialProfiles) {
              for (const pr of (Array.isArray(mat.MaterialProfiles) ? mat.MaterialProfiles : [mat.MaterialProfiles])) {
                const pid = pr?.value ?? pr; if (typeof pid !== 'number') continue;
                const p = ifcApi.GetLine(modelId, pid, false);
                const mid = p?.Material?.value ?? p?.Material;
                if (typeof mid === 'number' && materialColorCache.has(mid)) return materialColorCache.get(mid)!;
              }
            }
            if (typeId === IFCMATERIALLAYER) {
              const mid = mat.Material?.value ?? mat.Material;
              if (typeof mid === 'number' && materialColorCache.has(mid)) return materialColorCache.get(mid)!;
            }
            if (typeId === IFCMATERIAL) return materialColorCache.get(matId) ?? null;
          } catch { /* skip */ }
          return null;
        };

        const relIds = ifcApi.GetLineIDsWithType(modelId, IFCRELASSOCIATESMATERIAL);
        for (let i = 0; i < relIds.size(); i++) {
          try {
            const rel = ifcApi.GetLine(modelId, relIds.get(i), false);
            if (!rel?.RelatedObjects) continue;
            const matRef = rel.RelatingMaterial?.value ?? rel.RelatingMaterial;
            if (typeof matRef !== 'number') continue;
            const color = lookupColor(matRef);
            if (!color) continue;
            for (const obj of (Array.isArray(rel.RelatedObjects) ? rel.RelatedObjects : [rel.RelatedObjects])) {
              const eid = obj?.value ?? obj;
              if (typeof eid === 'number') expressIdToColor.set(eid, color);
            }
          } catch { continue; }
        }
      }

      console.log(`[Viewer] Total extracted ${expressIdToColor.size} element material colors from IFC`);

      // Apply colors to model
      if (expressIdToColor.size > 0) {
        const colorGroups = new Map<string, number[]>();
        for (const [expressId, color] of expressIdToColor) {
          const key = color.getHexString();
          if (!colorGroups.has(key)) colorGroups.set(key, []);
          colorGroups.get(key)!.push(expressId);
        }

        for (const [hex, ids] of colorGroups) {
          try {
            const color = new THREE.Color('#' + hex);
            await model.setColor(ids, color);
          } catch (e) {
            // Some IDs may not have geometry, skip silently
          }
        }

        console.log(`[Viewer] Applied ${colorGroups.size} distinct material colors to model`);
      }
    } finally {
      ifcApi.CloseModel(modelId);
      (ifcApi as any).wasmModule = undefined;
    }
  }

  private async ensureModelRegistered(model: any): Promise<void> {
    if (!model) return;
    const modelId = this.getModelInternalId(model, '');
    const modelObject = model.object as THREE.Object3D | undefined;
    const isObjectRegistered = (): boolean => (
      !!modelObject && [...this.federatedModels.values()].some((record) => record.object === modelObject)
    );
    if (modelId && (this.federatedModels.has(modelId) || this.registeringModelIds.has(modelId))) return;
    if (isObjectRegistered()) return;

    const existingEntry = [...this.fragments.list.entries()].find(([, entry]) => entry === model);
    if (existingEntry) {
      const [entryId, entryModel] = existingEntry;
      await this.onModelAdded(String(entryId), entryModel);
      return;
    }

    if (modelId) {
      const directEntry = this.findFragmentsEntryByInternalId(modelId);
      if (directEntry) {
        const [entryId, entryModel] = directEntry;
        await this.onModelAdded(entryId, entryModel);
        return;
      }
    }

    const timeoutMs = 8000;
    const pollMs = 40;
    const startedAt = performance.now();
    while (performance.now() - startedAt < timeoutMs) {
      await new Promise((resolve) => window.setTimeout(resolve, pollMs));

      const match = [...this.fragments.list.entries()].find(([, entry]) => entry === model);
      if (match) {
        const [entryId, entryModel] = match;
        await this.onModelAdded(String(entryId), entryModel);
        return;
      }

      if (modelId) {
        const byId = this.findFragmentsEntryByInternalId(modelId);
        if (byId) {
          const [entryId, entryModel] = byId;
          await this.onModelAdded(entryId, entryModel);
          return;
        }
        if (this.federatedModels.has(modelId)) return;
      }
      if (isObjectRegistered()) return;
    }

    throw new Error(`Model registration timed out: ${modelId || 'unknown model id'}`);
  }

  private getModelBoundingBox(): THREE.Box3 | null {
    if (this.modelObjects.length === 0) return null;
    const bbox = new THREE.Box3();
    for (const object of this.modelObjects) bbox.expandByObject(object);
    return bbox;
  }

  private fitToModel(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x + maxDim, center.y + maxDim * 0.75, center.z + maxDim, center.x, center.y, center.z, true);
  }

  private setFrontView(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x, center.y, center.z + maxDim * 2, center.x, center.y, center.z, true);
  }

  private setTopView(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x, center.y + maxDim * 2, center.z, center.x, center.y, center.z, true);
  }

  private setRightView(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x + maxDim * 2, center.y, center.z, center.x, center.y, center.z, true);
  }

  private setLeftView(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x - maxDim * 2, center.y, center.z, center.x, center.y, center.z, true);
  }

  private setBackView(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x, center.y, center.z - maxDim * 2, center.x, center.y, center.z, true);
  }

  private setBottomView(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) return;
    const center = bbox.getCenter(new THREE.Vector3());
    const size = bbox.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z);
    void this.world.camera.controls.setLookAt(center.x, center.y - maxDim * 2, center.z, center.x, center.y, center.z, true);
  }

  private updateViewCubeFromCamera(): void {
    if (!this.world?.camera) return;
    const camera = this.world.camera.three;
    const inverse = camera.quaternion.clone().invert();
    const euler = new THREE.Euler().setFromQuaternion(inverse, 'XYZ');
    // CSS 3D: Y-down, Three.js: Y-up → negate X and Y rotations
    const rx = -THREE.MathUtils.radToDeg(euler.x);
    const ry = -THREE.MathUtils.radToDeg(euler.y);
    const rz = THREE.MathUtils.radToDeg(euler.z);
    this.dom.viewCubeBody.style.transform = `rotateX(${rx.toFixed(2)}deg) rotateY(${ry.toFixed(2)}deg) rotateZ(${rz.toFixed(2)}deg)`;

    this.world.camera.controls.getTarget(this.cubeTarget);
    this.cubeCamDir.copy(camera.position).sub(this.cubeTarget).normalize();

    let active: 'top' | 'front' | 'right' | 'left' | 'back' | 'bottom' = 'front';
    const ax = Math.abs(this.cubeCamDir.x);
    const ay = Math.abs(this.cubeCamDir.y);
    const az = Math.abs(this.cubeCamDir.z);
    if (ay >= ax && ay >= az) {
      active = this.cubeCamDir.y >= 0 ? 'top' : 'bottom';
    } else if (ax >= ay && ax >= az) {
      active = this.cubeCamDir.x >= 0 ? 'right' : 'left';
    } else {
      active = this.cubeCamDir.z >= 0 ? 'front' : 'back';
    }

    this.dom.cubeTop.classList.toggle('active', active === 'top');
    this.dom.cubeFront.classList.toggle('active', active === 'front');
    this.dom.cubeRight.classList.toggle('active', active === 'right');
    this.dom.cubeLeft.classList.toggle('active', active === 'left');
    this.dom.cubeBack.classList.toggle('active', active === 'back');
    this.dom.cubeBottom.classList.toggle('active', active === 'bottom');
  }

  private applySelectionMode(mode: SelectionMode): void {
    this.selectionMode = mode;
    this.dom.btnSelectSingle.classList.toggle('active', mode === 'single');
    this.dom.btnSelectMulti.classList.toggle('active', mode === 'multi');
    this.setStatus(mode === 'single' ? 'Single-selection mode' : 'Multi-selection mode');
    this.persistLocalState();
  }

  private applyNavigationMode(mode: NavigationMode): void {
    if (!this.world?.camera) return;
    this.navigationMode = mode;
    this.world.camera.set(mode);
    this.dom.btnModeOrbit.classList.toggle('active', mode === 'Orbit');
    this.dom.btnModePlan.classList.toggle('active', mode === 'Plan');
    this.dom.btnModeFirstPerson.classList.toggle('active', mode === 'FirstPerson');
    this.dom.headerModeLabel.textContent = mode;
    this.persistLocalState();
  }

  private setMeasureMode(mode: MeasureMode): void {
    this.measureMode = mode;
    this.lengthMeasurement.enabled = mode === 'length';
    this.areaMeasurement.enabled = mode === 'area';
    this.dom.btnMeasureLength.classList.toggle('active', mode === 'length');
    this.dom.btnMeasureArea.classList.toggle('active', mode === 'area');
    this.setStatus(mode === 'none' ? 'Measure mode disabled' : `${mode === 'length' ? 'Length' : 'Area'} measurement enabled`);
  }

  private clearMeasurements(): void {
    this.lengthMeasurement.list.clear();
    this.areaMeasurement.list.clear();
    this.lengthMeasurement.cancelCreation();
    this.areaMeasurement.cancelCreation();
    this.setMeasureMode('none');
  }

  private addSectionPlane(normal: THREE.Vector3): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) {
      this.setStatus('No model to section');
      return;
    }
    const center = bbox.getCenter(new THREE.Vector3());
    this.clipper.enabled = true;
    this.clipper.createFromNormalAndCoplanarPoint(this.world, normal, center);
    this.setStatus('Section plane added');
  }

  private createSectionBox(): void {
    const bbox = this.getModelBoundingBox();
    if (!bbox || bbox.isEmpty()) {
      this.setStatus('No model to section');
      return;
    }
    this.clearSections(false);
    this.clipper.enabled = true;
    const { min, max } = bbox;
    this.clipper.createFromNormalAndCoplanarPoint(this.world, new THREE.Vector3(1, 0, 0), new THREE.Vector3(max.x, 0, 0));
    this.clipper.createFromNormalAndCoplanarPoint(this.world, new THREE.Vector3(-1, 0, 0), new THREE.Vector3(min.x, 0, 0));
    this.clipper.createFromNormalAndCoplanarPoint(this.world, new THREE.Vector3(0, 1, 0), new THREE.Vector3(0, max.y, 0));
    this.clipper.createFromNormalAndCoplanarPoint(this.world, new THREE.Vector3(0, -1, 0), new THREE.Vector3(0, min.y, 0));
    this.clipper.createFromNormalAndCoplanarPoint(this.world, new THREE.Vector3(0, 0, 1), new THREE.Vector3(0, 0, max.z));
    this.clipper.createFromNormalAndCoplanarPoint(this.world, new THREE.Vector3(0, 0, -1), new THREE.Vector3(0, 0, min.z));
    this.setStatus('Section box created');
  }

  private clearSections(updateStatus = true): void {
    this.clipper.deleteAll();
    this.clipper.enabled = false;
    this.dom.btnSectionX.classList.remove('active');
    this.dom.btnSectionY.classList.remove('active');
    this.dom.btnSectionZ.classList.remove('active');
    this.dom.btnSectionBox.classList.remove('active');
    if (updateStatus) this.setStatus('Sections cleared');
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

      const baseOpacity = this.clamp(record.opacity, 0, 1);
      const effectiveOpacity = this.xrayEnabled ? this.clamp(baseOpacity * 0.28, 0, 1) : baseOpacity;
      const targetOpacity = effectiveOpacity >= 0.999
        ? 1
        : this.clamp(Number(effectiveOpacity.toFixed(3)), 0.02, 0.999);
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
    for (const overlay of this.edgeOverlays) {
      this.world.scene.three.remove(overlay);
      overlay.geometry.dispose();
    }
    this.edgeOverlays.length = 0;
    if (!this.edgesEnabled) return;

    for (const object of this.modelObjects) {
      if (!object.visible) continue;
      object.traverse((child: THREE.Object3D) => {
        const mesh = child as THREE.Mesh;
        if (!mesh.isMesh || !mesh.geometry || !mesh.visible) return;
        try {
          const geometry = new THREE.EdgesGeometry(mesh.geometry, 35);
          const lines = new THREE.LineSegments(geometry, this.edgeMaterial);
          lines.matrixAutoUpdate = false;
          lines.matrix.copy(mesh.matrixWorld);
          this.world.scene.three.add(lines);
          this.edgeOverlays.push(lines);
        } catch {
          // ignore invalid geometry
        }
      });
    }
  }

  private async onViewerClick(_event: MouseEvent): Promise<void> {
    if (this.pointerDragged || this.gizmoDragging || !!this.activeGizmoModelId) return;

    if (this.measureMode !== 'none') {
      if (this.measureMode === 'length') await this.lengthMeasurement.create();
      if (this.measureMode === 'area') await this.areaMeasurement.create();
      return;
    }

    const result = await this.raycaster.castRay() as any;
    if (!result || !result.fragments || result.localId === undefined) {
      if (this.selectionMode === 'single' && !this.issuePinMode) await this.clearSelection();
      return;
    }

    const modelId = String(result.fragments.modelId);
    const resolvedModelId = this.resolveModelId(modelId);
    if (!resolvedModelId) return;
    const localId = result.localId as number;
    if (result.point) this.lastHitPoint = result.point.clone();

    if (this.issuePinMode) {
      if (result.point) this.pendingIssuePoint = result.point.clone();
      if (this.selectionMode === 'single') await this.selectSingleItem(resolvedModelId, localId, false);
      this.activateTab('issues');
      this.setStatus('Issue point captured. Fill issue form and create issue');
      return;
    }

    if (this.selectionMode === 'single') {
      await this.selectSingleItem(resolvedModelId, localId, false);
      return;
    }

    this.toggleSelectionItem(resolvedModelId, localId);
    await this.refreshSelectionVisuals();
  }

  private async selectSingleItem(modelId: string, localId: number, zoomToItem: boolean): Promise<void> {
    const resolvedModelId = this.resolveModelId(modelId);
    if (!resolvedModelId) return;
    clearMap(this.selectedItems);
    this.selectedItems[resolvedModelId] = new Set([localId]);
    await this.refreshSelectionVisuals();
    if (zoomToItem) await this.zoomToItems(this.selectedItems);
  }

  private toggleSelectionItem(modelId: string, localId: number): void {
    const resolvedModelId = this.resolveModelId(modelId);
    if (!resolvedModelId) return;
    if (!this.selectedItems[resolvedModelId]) this.selectedItems[resolvedModelId] = new Set<number>();
    const set = this.selectedItems[resolvedModelId];
    if (set.has(localId)) set.delete(localId);
    else set.add(localId);
    if (set.size === 0) delete this.selectedItems[resolvedModelId];
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
    await this.fragments.core.update(true);
    this.updateCounters();
    await this.updatePropertiesPanel();
  }

  private getFirstSelection(): { modelId: string; localId: number } | null {
    for (const [modelId, ids] of Object.entries(this.selectedItems)) {
      for (const localId of ids) return { modelId, localId };
    }
    return null;
  }

  private unwrapIfcValue(value: unknown): unknown {
    let current = value;
    for (let i = 0; i < 8; i += 1) {
      if (!current || typeof current !== 'object' || Array.isArray(current)) break;
      const record = current as Record<string, unknown>;
      if (!Object.prototype.hasOwnProperty.call(record, 'value')) break;
      const next = record.value;
      if (next === undefined || next === current) break;
      current = next;
    }
    return current;
  }

  private truncatePropertyValue(value: string, maxLength = MAX_PROPERTY_VALUE_LENGTH): string {
    if (value.length <= maxLength) return value;
    return `${value.slice(0, maxLength - 1)}…`;
  }

  private summarizeObject(record: Record<string, unknown>): string {
    const preferredKeys = [
      'Name',
      'LongName',
      'ObjectType',
      'PredefinedType',
      'Description',
      'type',
      '_category',
      '_localId',
      'localId',
      'GlobalId',
      '_guid',
      'Tag',
    ];
    const parts: string[] = [];

    for (const key of preferredKeys) {
      if (!Object.prototype.hasOwnProperty.call(record, key)) continue;
      const value = this.unwrapIfcValue(record[key]);
      if (value === null || value === undefined || typeof value === 'object') continue;
      const text = this.toPropertyString(value, '');
      if (!text) continue;
      parts.push(`${key}: ${text}`);
      if (parts.length >= 4) break;
    }

    if (parts.length > 0) return this.truncatePropertyValue(parts.join(' | '));

    const keys = Object.keys(record);
    if (keys.length === 0) return '{}';
    const preview = keys.slice(0, 5).join(', ');
    return keys.length > 5 ? `{${preview}, ...}` : `{${preview}}`;
  }

  private toPropertyString(value: unknown, fallback = '-'): string {
    const normalized = this.unwrapIfcValue(value);
    if (normalized === null || normalized === undefined) return fallback;
    if (typeof normalized === 'string') return normalized.trim().length > 0 ? normalized : fallback;
    if (typeof normalized === 'number' || typeof normalized === 'boolean' || typeof normalized === 'bigint') return String(normalized);
    if (normalized instanceof Date) return normalized.toISOString();
    if (Array.isArray(normalized)) {
      if (normalized.length === 0) return '[]';
      const values = normalized
        .map((entry) => this.toPropertyString(entry, ''))
        .filter((entry) => entry.length > 0);
      if (values.length === 0) return fallback;
      return this.truncatePropertyValue(values.join(', '));
    }
    if (typeof normalized === 'object') {
      return this.summarizeObject(normalized as Record<string, unknown>);
    }
    return String(normalized);
  }

  private flattenPropertyEntries(
    value: unknown,
    prefix: string,
    output: Array<[string, string]>,
    visited: WeakSet<object>,
    state: { truncated: boolean },
    depth = 0,
  ): void {
    if (output.length >= MAX_PROPERTY_ROWS) {
      state.truncated = true;
      return;
    }

    const normalized = this.unwrapIfcValue(value);
    if (normalized === null || normalized === undefined) return;
    if (depth > MAX_PROPERTY_DEPTH) {
      output.push([prefix, '...']);
      return;
    }

    if (
      typeof normalized === 'string'
      || typeof normalized === 'number'
      || typeof normalized === 'boolean'
      || typeof normalized === 'bigint'
    ) {
      output.push([prefix, String(normalized)]);
      return;
    }

    if (normalized instanceof Date) {
      output.push([prefix, normalized.toISOString()]);
      return;
    }

    if (Array.isArray(normalized)) {
      if (normalized.length === 0) {
        output.push([prefix, '[]']);
        return;
      }

      const scalarOnly = normalized.every((entry) => {
        const item = this.unwrapIfcValue(entry);
        return (
          item === null
          || item === undefined
          || typeof item === 'string'
          || typeof item === 'number'
          || typeof item === 'boolean'
          || typeof item === 'bigint'
        );
      });

      if (scalarOnly) {
        const joined = normalized
          .map((entry) => this.toPropertyString(entry, ''))
          .filter((entry) => entry.length > 0)
          .join(', ');
        output.push([prefix, this.truncatePropertyValue(joined || `${normalized.length} entries`)]);
        return;
      }

      const preview = normalized
        .slice(0, MAX_PROPERTY_ARRAY_PREVIEW)
        .map((entry) => this.toPropertyString(entry, ''))
        .filter((entry) => entry.length > 0)
        .join(' || ');
      const suffix = normalized.length > MAX_PROPERTY_ARRAY_PREVIEW
        ? ` (+${normalized.length - MAX_PROPERTY_ARRAY_PREVIEW} more)`
        : '';
      const summary = preview
        ? `${normalized.length} entries: ${preview}${suffix}`
        : `${normalized.length} entries${suffix}`;
      output.push([prefix, this.truncatePropertyValue(summary)]);
      return;
    }

    if (typeof normalized === 'object') {
      const record = normalized as Record<string, unknown>;
      if (visited.has(record)) {
        output.push([prefix, '[Circular]']);
        return;
      }
      visited.add(record);

      const entries = Object.entries(record);
      if (entries.length === 0) {
        output.push([prefix, '{}']);
        return;
      }

      if (depth >= 2) {
        output.push([prefix, this.summarizeObject(record)]);
        return;
      }

      for (const [key, entryValue] of entries) {
        const nextPrefix = prefix ? `${prefix}.${key}` : key;
        const unwrapped = this.unwrapIfcValue(entryValue);
        if (depth >= 1 && unwrapped && typeof unwrapped === 'object') {
          if (Array.isArray(unwrapped)) {
            const preview = unwrapped
              .slice(0, MAX_PROPERTY_ARRAY_PREVIEW)
              .map((entry) => this.toPropertyString(entry, ''))
              .filter((entry) => entry.length > 0)
              .join(' || ');
            const suffix = unwrapped.length > MAX_PROPERTY_ARRAY_PREVIEW
              ? ` (+${unwrapped.length - MAX_PROPERTY_ARRAY_PREVIEW} more)`
              : '';
            output.push([nextPrefix, this.truncatePropertyValue(`${unwrapped.length} entries${preview ? `: ${preview}` : ''}${suffix}`)]);
          } else {
            output.push([nextPrefix, this.summarizeObject(unwrapped as Record<string, unknown>)]);
          }
          if (output.length >= MAX_PROPERTY_ROWS) {
            state.truncated = true;
            return;
          }
          continue;
        }
        this.flattenPropertyEntries(entryValue, nextPrefix, output, visited, state, depth + 1);
        if (state.truncated) return;
      }
      return;
    }

    output.push([prefix, String(normalized)]);
  }

  private async updatePropertiesPanel(): Promise<void> {
    if (isMapEmpty(this.selectedItems)) {
      this.dom.propsEmpty.hidden = false;
      this.dom.propsContent.hidden = true;
      this.dom.propAttributes.innerHTML = '';
      return;
    }

    const firstSelection = this.getFirstSelection();
    if (!firstSelection) return;
    const model = this.getFragmentsModel(firstSelection.modelId);
    if (!model) return;

    const itemData = await model.getItemsData([firstSelection.localId], {
      attributesDefault: true,
      relationsDefault: { attributes: true, relations: true },
    });
    const data = (itemData[0] || {}) as Record<string, unknown>;

    const typeValue = [
      this.toPropertyString(data.ObjectType, ''),
      this.toPropertyString(data.PredefinedType, ''),
      this.toPropertyString(data.type, ''),
      this.toPropertyString(data._category, ''),
    ].find((entry) => entry.length > 0) || '-';

    const nameValue = this.toPropertyString(data.Name, '');
    this.dom.propType.textContent = typeValue;
    this.dom.propName.textContent = nameValue || `Element ${firstSelection.localId}`;
    this.dom.propGlobalId.textContent = this.toPropertyString(data.GlobalId, '-');
    this.dom.propDescription.textContent = this.toPropertyString(data.Description, '-');
    this.dom.propStory.textContent = this.modelIndices.get(firstSelection.modelId)?.itemToLevel.get(firstSelection.localId) || '-';

    const attributes: Array<[string, string]> = [];
    const skipKeys = new Set(['Name', 'GlobalId', 'Description', 'ObjectType', 'PredefinedType']);
    const visited = new WeakSet<object>();
    const flattenState = { truncated: false };
    for (const [key, value] of Object.entries(data)) {
      if (skipKeys.has(key)) continue;
      this.flattenPropertyEntries(value, key, attributes, visited, flattenState, 0);
      if (flattenState.truncated) break;
    }

    try {
      const volume = await model.getItemsVolume([firstSelection.localId]);
      attributes.unshift(['Volume', `${volume.toFixed(3)} m3`]);
    } catch {
      // optional
    }

    const deduped = new Map<string, string>();
    for (const [key, value] of attributes) {
      if (!key) continue;
      if (!deduped.has(key)) deduped.set(key, value);
    }
    if (flattenState.truncated && !deduped.has('Info')) {
      deduped.set('Info', `Properties truncated to ${MAX_PROPERTY_ROWS} rows for readability`);
    }
    const sortedAttributes = [...deduped.entries()].sort(([a], [b]) => a.localeCompare(b));
    if (sortedAttributes.length === 0) sortedAttributes.push(['Info', 'No additional attributes']);

    this.dom.propAttributes.innerHTML = sortedAttributes
      .map(([key, value]) => `
        <div class="prop-row">
          <span class="prop-key">${escapeHtml(key)}</span>
          <span class="prop-val">${escapeHtml(value)}</span>
        </div>
      `)
      .join('');

    this.dom.propsEmpty.hidden = true;
    this.dom.propsContent.hidden = false;
    this.setStatus(`${countMapItems(this.selectedItems)} element(s) selected`);
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
      this.setStatus('Enter a viewpoint name');
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

    const renderer = this.world.renderer;
    if (!renderer) {
      this.setStatus('Renderer unavailable for snapshot');
      return;
    }
    const snapshot = renderer.three.domElement.toDataURL('image/png');

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
    this.setStatus(`Saved viewpoint: ${name}`);
  }

  private async applySelectedViewpoint(): Promise<void> {
    if (!this.selectedViewpointId) {
      this.setStatus('Select a viewpoint first');
      return;
    }

    const viewpoint = this.viewpoints.find((entry) => entry.id === this.selectedViewpointId);
    if (!viewpoint) {
      this.setStatus('Selected viewpoint not found');
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
    this.clipper.enabled = viewpoint.clippingPlanes.length > 0;
    for (const plane of viewpoint.clippingPlanes) {
      this.clipper.createFromNormalAndCoplanarPoint(
        this.world,
        new THREE.Vector3(plane.normal.x, plane.normal.y, plane.normal.z),
        new THREE.Vector3(plane.origin.x, plane.origin.y, plane.origin.z),
      );
    }

    await this.hider.set(true);
    const hiddenMap = this.getValidModelIdMap(toSetMap(viewpoint.hiddenItems));
    if (!isMapEmpty(hiddenMap)) await this.hider.set(false, hiddenMap);

    const viewpointStyle = this.parseVisualStyle(viewpoint.visualStyle ?? 'color-pen-shadows');
    await this.setVisualStyle(viewpointStyle, false, false);
    this.xrayEnabled = viewpoint.xray;
    this.edgesEnabled = viewpoint.edges;
    this.dom.btnTransparency.classList.toggle('active', this.xrayEnabled);
    this.dom.btnWireframe.classList.toggle('active', this.edgesEnabled);
    this.applyXRay();
    this.applyEdges();

    await this.updateVisibilityCount();
    this.setStatus(`Applied viewpoint: ${viewpoint.name}`);
  }

  private deleteSelectedViewpoint(): void {
    if (!this.selectedViewpointId) {
      this.setStatus('Select a viewpoint first');
      return;
    }

    this.viewpoints = this.viewpoints.filter((entry) => entry.id !== this.selectedViewpointId);
    this.selectedViewpointId = this.viewpoints[0]?.id ?? null;
    this.updateViewpointList();
    this.persistLocalState();
    this.setStatus('Viewpoint deleted');
  }

  private updateViewpointList(): void {
    if (this.viewpoints.length === 0) {
      this.dom.viewpointList.innerHTML = '<div class="viewpoint-item">No viewpoints saved</div>';
      return;
    }

    this.dom.viewpointList.innerHTML = this.viewpoints
      .map((entry) => {
        const active = entry.id === this.selectedViewpointId ? 'active' : '';
        return `
          <div class="viewpoint-item ${active}" data-viewpoint-id="${entry.id}">
            <div><strong>${entry.name}</strong></div>
            <div>${new Date(entry.createdAt).toLocaleString()}</div>
          </div>
        `;
      })
      .join('');

    this.dom.viewpointList.querySelectorAll<HTMLElement>('[data-viewpoint-id]').forEach((element) => {
      element.addEventListener('click', () => {
        this.selectedViewpointId = element.dataset.viewpointId || null;
        this.updateViewpointList();
      });
      element.addEventListener('dblclick', () => {
        this.selectedViewpointId = element.dataset.viewpointId || null;
        this.fireAndForget(this.applySelectedViewpoint(), 'Apply viewpoint');
      });
    });
  }

  private async createIssueFromCurrentContext(): Promise<void> {
    const title = this.dom.issueTitle.value.trim();
    if (!title) {
      this.setStatus('Issue title is required');
      return;
    }

    const selectedCount = countMapItems(this.selectedItems);
    if (selectedCount === 0 && !this.pendingIssuePoint && !this.lastHitPoint) {
      this.setStatus('Select element(s) or use issue pin mode to capture a point');
      return;
    }

    const firstSelection = this.getFirstSelection();
    const point = this.pendingIssuePoint ?? this.lastHitPoint;
    const issuePoint = point ? { x: point.x, y: point.y, z: point.z } : null;

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
    this.setStatus('Issue created');
  }

  private createIssueMarker(issue: IssueRecord): void {
    if (!issue.point) return;

    if (issue.markerId) {
      this.markerManager.delete(issue.markerId);
      issue.markerId = undefined;
    }

    const markerElement = document.createElement('button');
    markerElement.type = 'button';
    markerElement.className = `issue-marker issue-${issue.status.toLowerCase().replace(/\s+/g, '-')}`;
    markerElement.textContent = issue.priority[0];
    markerElement.title = `${issue.title} (${issue.status})`;

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
  }

  private updateIssuesList(): void {
    if (this.issues.length === 0) {
      this.dom.issuesList.innerHTML = '<div class="issue-item">No issues</div>';
      return;
    }

    this.dom.issuesList.innerHTML = this.issues
      .map((issue) => {
        const active = issue.id === this.activeIssueId ? 'active' : '';
        const linked = issue.localIds.length > 0 ? `${issue.localIds.length} linked` : 'No element link';
        return `
          <div class="issue-item ${active}" data-issue-id="${issue.id}">
            <div><strong>${issue.title}</strong></div>
            <div>${issue.status} | ${issue.priority}</div>
            <div>${linked}</div>
          </div>
        `;
      })
      .join('');

    this.dom.issuesList.querySelectorAll<HTMLElement>('[data-issue-id]').forEach((element) => {
      element.addEventListener('click', () => {
        const id = element.dataset.issueId;
        if (!id) return;
        this.selectIssue(id, false);
      });
      element.addEventListener('dblclick', () => {
        const id = element.dataset.issueId;
        if (!id) return;
        this.selectIssue(id, true);
      });
    });
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

    if (issue.modelId && issue.localIds.length > 0 && this.isLoadedModelId(issue.modelId)) {
      const selection: OBC.ModelIdMap = { [issue.modelId]: new Set(issue.localIds) };
      clearMap(this.selectedItems);
      Object.assign(this.selectedItems, selection);
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

    this.activateTab('issues');
    this.setStatus(`Selected issue: ${issue.title}`);
  }

  private deleteSelectedIssue(): void {
    if (!this.activeIssueId) {
      this.setStatus('Select an issue first');
      return;
    }

    const issue = this.issues.find((entry) => entry.id === this.activeIssueId);
    if (issue?.markerId) this.markerManager.delete(issue.markerId);

    this.issues = this.issues.filter((entry) => entry.id !== this.activeIssueId);
    this.activeIssueId = this.issues[0]?.id ?? null;

    this.updateIssuesList();
    this.updateIssueComments();
    this.persistLocalState();
    this.setStatus('Issue deleted');
  }

  private addCommentToActiveIssue(): void {
    if (!this.activeIssueId) {
      this.setStatus('Select an issue first');
      return;
    }

    const text = this.dom.issueCommentInput.value.trim();
    if (!text) {
      this.setStatus('Comment cannot be empty');
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
    this.setStatus('Comment added');
  }

  private updateIssueComments(): void {
    const issue = this.issues.find((entry) => entry.id === this.activeIssueId);
    if (!issue) {
      this.dom.issueComments.innerHTML = '<div class="comment-item">Select an issue to view comments</div>';
      return;
    }

    if (issue.comments.length === 0) {
      this.dom.issueComments.innerHTML = '<div class="comment-item">No comments</div>';
      return;
    }

    this.dom.issueComments.innerHTML = issue.comments
      .map((comment) => `
        <div class="comment-item">
          <div><strong>${comment.author}</strong> - ${new Date(comment.createdAt).toLocaleString()}</div>
          <div>${comment.text}</div>
        </div>
      `)
      .join('');
  }

  private exportScreenshot(): void {
    if (!this.world?.renderer) return;
    const dataUrl = this.world.renderer.three.domElement.toDataURL('image/png');
    fetch(dataUrl)
      .then((response) => response.blob())
      .then((blob) => {
        const name = `bim-view-${new Date().toISOString().replace(/[:.]/g, '-')}.png`;
        downloadBlob(name, blob);
        this.setStatus('Screenshot exported');
      })
      .catch((error) => {
        this.setStatus(`Screenshot export failed: ${serializeError(error)}`);
      });
  }

  private exportViewerState(): void {
    const payload = this.getPersistedState();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const name = `viewer-state-${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
    downloadBlob(name, blob);
    this.setStatus('Viewer data exported');
  }

  private async importViewerState(file: File): Promise<void> {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as PersistedViewerState;
      if (parsed.version !== 1) throw new Error('Unsupported viewer state version');

      this.selectionMode = parsed.selectionMode;
      this.navigationMode = parsed.navigationMode;
      this.visualStyle = this.parseVisualStyle(parsed.visualStyle ?? (parsed.xray ? 'xray' : 'shaded'));
      const restoredXray = parsed.xray;
      const restoredEdges = parsed.edges;
      this.gridVisible = parsed.gridVisible ?? this.gridVisible;
      this.backgroundColor = this.normalizeHexColor(parsed.backgroundColor ?? this.backgroundColor);
      this.viewpoints = parsed.viewpoints;
      this.issues = parsed.issues.map((issue) => ({ ...issue }));

      this.applySelectionMode(this.selectionMode);
      this.applyNavigationMode(this.navigationMode);

      await this.setVisualStyle(this.visualStyle, false, false);
      this.xrayEnabled = restoredXray;
      this.edgesEnabled = restoredEdges;
      this.dom.btnTransparency.classList.toggle('active', this.xrayEnabled);
      this.dom.btnWireframe.classList.toggle('active', this.edgesEnabled);
      this.setGridVisible(this.gridVisible, false);
      this.setBackgroundColor(this.backgroundColor, false);
      this.syncVisualSettingsUi();
      this.applyXRay();
      this.applyEdges();

      this.updateViewpointList();
      this.updateIssuesList();
      this.updateIssueComments();
      this.refreshIssueMarkers();

      this.persistLocalState();
      this.setStatus('Viewer data imported');
    } catch (error) {
      this.setStatus(`Import failed: ${serializeError(error)}`);
    }
  }

  private getPersistedState(): PersistedViewerState {
    const issues = this.issues.map((issue) => ({
      id: issue.id,
      title: issue.title,
      description: issue.description,
      priority: issue.priority,
      status: issue.status,
      assignee: issue.assignee,
      createdAt: issue.createdAt,
      updatedAt: issue.updatedAt,
      modelId: issue.modelId,
      localIds: [...issue.localIds],
      point: issue.point ? { ...issue.point } : null,
      comments: issue.comments.map((comment) => ({ ...comment })),
    }));

    return {
      version: 1,
      selectionMode: this.selectionMode,
      navigationMode: this.navigationMode,
      visualStyle: this.visualStyle,
      xray: this.xrayEnabled,
      edges: this.edgesEnabled,
      gridVisible: this.gridVisible,
      backgroundColor: this.backgroundColor,
      viewpoints: this.viewpoints,
      issues,
    };
  }

  private persistLocalState(): void {
    try {
      const payload = this.getPersistedState();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (error) {
      this.setStatus(`Unable to persist local state: ${serializeError(error)}`);
    }
  }

  private async restoreLocalState(): Promise<void> {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as PersistedViewerState;
      if (parsed.version !== 1) return;

      this.selectionMode = parsed.selectionMode;
      this.navigationMode = parsed.navigationMode;
      this.visualStyle = this.parseVisualStyle(parsed.visualStyle ?? 'color-pen-shadows');
      const restoredXray = parsed.xray;
      const restoredEdges = parsed.edges;
      this.gridVisible = parsed.gridVisible ?? this.gridVisible;
      this.backgroundColor = this.normalizeHexColor(parsed.backgroundColor ?? this.backgroundColor);
      this.viewpoints = parsed.viewpoints ?? [];
      this.issues = (parsed.issues ?? []).map((issue) => ({ ...issue }));

      await this.setVisualStyle(this.visualStyle, false, false);
      this.xrayEnabled = restoredXray;
      this.edgesEnabled = restoredEdges;
      this.dom.btnTransparency.classList.toggle('active', this.xrayEnabled);
      this.dom.btnWireframe.classList.toggle('active', this.edgesEnabled);
      this.setGridVisible(this.gridVisible, false);
      this.setBackgroundColor(this.backgroundColor, false);
      this.syncVisualSettingsUi();
      this.updateViewpointList();
      this.updateIssuesList();
      this.updateIssueComments();
      this.applyXRay();
      this.applyEdges();
      this.refreshIssueMarkers();
    } catch (error) {
      this.setStatus(`Failed to restore local state: ${serializeError(error)}`);
    }
  }

  private activateTab(tab: string): void {
    this.dom.tabButtons.forEach((button) => {
      button.classList.toggle('active', button.dataset.tab === tab);
    });

    document.querySelectorAll<HTMLElement>('.tab-panel').forEach((panel) => {
      panel.classList.toggle('active', panel.id === `panel-${tab}`);
    });
  }

  private async updateVisibilityCount(): Promise<void> {
    const visibleMap = await this.hider.getVisibilityMap(true);
    let visibleCount = 0;
    for (const [modelId, ids] of Object.entries(visibleMap)) {
      const resolvedModelId = this.resolveModelId(modelId) ?? modelId;
      const model = this.federatedModels.get(resolvedModelId);
      if (model && !model.visible) continue;
      visibleCount += ids.length;
    }
    this.dom.visibleCount.textContent = `${visibleCount} visible`;
  }

  private updateCounters(): void {
    this.dom.selectionCount.textContent = `${countMapItems(this.selectedItems)} selected`;
  }

  private onKeyDown(event: KeyboardEvent): void {
    const target = event.target as HTMLElement | null;
    const isFormTarget = !!target && (
      target.tagName === 'INPUT'
      || target.tagName === 'TEXTAREA'
      || target.tagName === 'SELECT'
      || target.isContentEditable
    );
    if (isFormTarget) return;

    if (event.key === 'Escape') {
      this.setMeasureMode('none');
      this.issuePinMode = false;
      this.dom.viewerHint.hidden = true;
      this.dom.btnIssuePinMode.classList.remove('active');
      if (this.activeGizmoModelId) this.detachModelGizmo();
      this.setStatus('Active tool canceled');
      return;
    }

    switch (event.key.toLowerCase()) {
      case 'f':
        this.fitToModel();
        break;
      case '1':
        this.applyNavigationMode('Orbit');
        break;
      case '2':
        this.applyNavigationMode('Plan');
        break;
      case '3':
        this.applyNavigationMode('FirstPerson');
        break;
      case 'm':
        this.applySelectionMode(this.selectionMode === 'single' ? 'multi' : 'single');
        break;
      case 'l':
        this.setMeasureMode(this.measureMode === 'length' ? 'none' : 'length');
        break;
      case 'a':
        this.setMeasureMode(this.measureMode === 'area' ? 'none' : 'area');
        break;
      case 'g':
        this.dom.toggleGrid.click();
        break;
      case 'x':
        this.dom.btnTransparency.click();
        break;
      case 'e':
        if (this.activeGizmoModelId) {
          this.transformControls?.setMode('rotate');
          this.setStatus('Gizmo mode: rotate');
        } else {
          this.dom.btnWireframe.click();
        }
        break;
      case 'w':
        if (this.activeGizmoModelId) {
          this.transformControls?.setMode('translate');
          this.setStatus('Gizmo mode: translate');
        }
        break;
      case 'r':
        if (this.activeGizmoModelId) {
          this.resetModelOffsets(this.activeGizmoModelId);
          this.setStatus('Gizmo: model transform reset');
        }
        break;
      case 'i':
        this.dom.btnIssuePinMode.click();
        break;
      case 'delete':
        this.deleteSelectedIssue();
        break;
      case 'enter':
        if (this.measureMode === 'area') this.areaMeasurement.endCreation();
        break;
      default:
        break;
    }
  }

  private setStatus(message: string): void {
    this.dom.statusText.textContent = message;
  }

  private startFpsMonitor(): void {
    const tick = (): void => {
      this.frameCount += 1;
      const now = performance.now();
      if (now - this.fpsLastTs >= 1000) {
        const fps = this.frameCount;
        this.frameCount = 0;
        this.fpsLastTs = now;
        this.dom.perfInfo.textContent = `${fps} FPS`;
      }
      requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  }
}

void new ViewerApp().init();
