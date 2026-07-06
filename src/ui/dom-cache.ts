/**
 * DOM element cache. Resolves every element the viewer needs up front (the ~90
 * required ids + the query-selected node lists) so a renamed/missing id throws
 * at construction with a clear message rather than surfacing as a later null
 * dereference (AUDIT A4). Extracted from viewer.ts so the orchestrator holds a
 * typed `ViewerDom` handle and this module owns the id list.
 */

const required = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing required DOM element #${id}`);
  return element as T;
};

export type ViewerDom = ReturnType<typeof createDomCache>;

/** Builds the DOM cache. Call once at construction (throws on any missing id). */
export function createDomCache() {
  return {
    root: required<HTMLDivElement>('btc-viewer-root'),
    viewerContainer: required<HTMLDivElement>('viewer-container'),
    // Top bar
    topbarModel: required<HTMLDivElement>('topbarModel'),
    btnUpload: required<HTMLButtonElement>('btnUpload'),
    btnUploadEmpty: required<HTMLButtonElement>('btnUploadEmpty'),
    fileInput: required<HTMLInputElement>('fileInput'),
    btnExportScreenshot: required<HTMLButtonElement>('btnExportScreenshot'),
    btnExportState: required<HTMLButtonElement>('btnExportState'),
    btnImportState: required<HTMLButtonElement>('btnImportState'),
    importStateInput: required<HTMLInputElement>('importStateInput'),
    btnThemeToggle: required<HTMLButtonElement>('btnThemeToggle'),
    btnPanelToggle: required<HTMLButtonElement>('btnPanelToggle'),
    // Overlays
    emptyState: required<HTMLDivElement>('emptyState'),
    loadingOverlay: required<HTMLDivElement>('loadingOverlay'),
    loadingText: required<HTMLDivElement>('loadingText'),
    loadingPct: required<HTMLSpanElement>('loadingPct'),
    loadingProgress: required<HTMLDivElement>('loadingProgress'),
    loadingErrorActions: required<HTMLDivElement>('loadingErrorActions'),
    btnRetryLoad: required<HTMLButtonElement>('btnRetryLoad'),
    btnDismissLoadError: required<HTMLButtonElement>('btnDismissLoadError'),
    viewerHint: required<HTMLDivElement>('viewerHint'),
    navPill: required<HTMLDivElement>('navPill'),
    measureHint: required<HTMLDivElement>('measureHint'),
    measureHintText: required<HTMLSpanElement>('measureHintText'),
    btnCancelMeasure: required<HTMLButtonElement>('btnCancelMeasure'),
    sectionSlider: required<HTMLDivElement>('sectionSlider'),
    sectionLabel: required<HTMLSpanElement>('sectionLabel'),
    sectionPos: required<HTMLInputElement>('sectionPos'),
    sectionPosLabel: required<HTMLSpanElement>('sectionPosLabel'),
    btnClearSectionSlider: required<HTMLButtonElement>('btnClearSectionSlider'),
    selectionChip: required<HTMLDivElement>('selectionChip'),
    selChipName: required<HTMLDivElement>('selChipName'),
    selChipMeta: required<HTMLDivElement>('selChipMeta'),
    btnClearSelection: required<HTMLButtonElement>('btnClearSelection'),
    // Status bar
    statusText: required<HTMLSpanElement>('statusText'),
    selectionCount: required<HTMLSpanElement>('selectionCount'),
    elementCount: required<HTMLSpanElement>('elementCount'),
    visibleCount: required<HTMLSpanElement>('visibleCount'),
    loadInfo: required<HTMLSpanElement>('loadInfo'),
    perfInfo: required<HTMLSpanElement>('perfInfo'),
    viewLabel: required<HTMLSpanElement>('viewLabel'),
    // View controls (rail replaces dock; cube widget removed per design)
    btnModeOrbit: required<HTMLButtonElement>('btnModeOrbit'),
    btnModePlan: required<HTMLButtonElement>('btnModePlan'),
    btnModeFirstPerson: required<HTMLButtonElement>('btnModeFirstPerson'),
    btnFitAll: required<HTMLButtonElement>('btnFitAll'),
    btnFront: required<HTMLButtonElement>('btnFront'),
    btnTop: required<HTMLButtonElement>('btnTop'),
    cubeHome: required<HTMLButtonElement>('cubeHome'),
    // Tool rail
    btnSelectSingle: required<HTMLButtonElement>('btnSelectSingle'),
    btnSelectMulti: required<HTMLButtonElement>('btnSelectMulti'),
    btnIsolate: required<HTMLButtonElement>('btnIsolate'),
    btnHide: required<HTMLButtonElement>('btnHide'),
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
    btnToggleGrid: required<HTMLButtonElement>('btnToggleGrid'),
    btnIssuePinMode: required<HTMLButtonElement>('btnIssuePinMode'),
    // Splitter + panel
    panelSplitter: required<HTMLDivElement>('panelSplitter'),
    panelTitle: required<HTMLSpanElement>('panelTitle'),
    tabStripButtons: Array.from(document.querySelectorAll<HTMLButtonElement>('.tab-strip-btn')),
    tabPanels: Array.from(document.querySelectorAll<HTMLDivElement>('.tab-panel')),
    // Explorer
    searchInput: required<HTMLInputElement>('searchInput'),
    btnClearSearch: required<HTMLButtonElement>('btnClearSearch'),
    searchResultsGroup: required<HTMLDivElement>('searchResultsGroup'),
    elementResults: required<HTMLDivElement>('elementResults'),
    classFilterList: required<HTMLDivElement>('classFilterList'),
    levelFilterList: required<HTMLDivElement>('levelFilterList'),
    modelBrowserTree: required<HTMLDivElement>('modelBrowserTree'),
    // Models
    federationTree: required<HTMLDivElement>('federationTree'),
    // Properties
    propsEmpty: required<HTMLDivElement>('propsEmpty'),
    propsContent: required<HTMLDivElement>('propsContent'),
    propType: required<HTMLSpanElement>('propType'),
    propName: required<HTMLSpanElement>('propName'),
    propGlobalId: required<HTMLSpanElement>('propGlobalId'),
    propDescription: required<HTMLSpanElement>('propDescription'),
    propStory: required<HTMLSpanElement>('propStory'),
    propFilterInput: required<HTMLInputElement>('propFilterInput'),
    propSections: required<HTMLDivElement>('propSections'),
    btnPropsIsolate: required<HTMLButtonElement>('btnPropsIsolate'),
    btnPropsHide: required<HTMLButtonElement>('btnPropsHide'),
    // Viewpoints
    viewpointName: required<HTMLInputElement>('viewpointName'),
    btnSaveViewpoint: required<HTMLButtonElement>('btnSaveViewpoint'),
    viewpointList: required<HTMLDivElement>('viewpointList'),
    // Issues
    issueTitle: required<HTMLInputElement>('issueTitle'),
    issueDescription: required<HTMLTextAreaElement>('issueDescription'),
    issuePriority: required<HTMLSelectElement>('issuePriority'),
    issueStatus: required<HTMLSelectElement>('issueStatus'),
    issueAssignee: required<HTMLInputElement>('issueAssignee'),
    btnCreateIssue: required<HTMLButtonElement>('btnCreateIssue'),
    btnDeleteIssue: required<HTMLButtonElement>('btnDeleteIssue'),
    issuesList: required<HTMLDivElement>('issuesList'),
    issueCommentsGroup: required<HTMLDivElement>('issueCommentsGroup'),
    issueCommentInput: required<HTMLInputElement>('issueCommentInput'),
    btnAddIssueComment: required<HTMLButtonElement>('btnAddIssueComment'),
    issueComments: required<HTMLDivElement>('issueComments'),
    // Mobile + sheet + scrim + confirm + toasts
    mobileFab: required<HTMLButtonElement>('mobileFab'),
    mobileNavButtons: Array.from(document.querySelectorAll<HTMLButtonElement>('[data-mobile-nav]')),
    scrim: required<HTMLDivElement>('scrim'),
    mobileSheet: required<HTMLDivElement>('mobileSheet'),
    sheetTitle: required<HTMLSpanElement>('sheetTitle'),
    sheetBody: required<HTMLDivElement>('sheetBody'),
    btnCloseSheet: required<HTMLButtonElement>('btnCloseSheet'),
    confirmDialog: required<HTMLDialogElement>('confirmDialog'),
    confirmMessage: required<HTMLParagraphElement>('confirmMessage'),
    confirmOk: required<HTMLButtonElement>('confirmOk'),
    confirmCancel: required<HTMLButtonElement>('confirmCancel'),
    toastRegion: required<HTMLDivElement>('toastRegion'),
  };
}
