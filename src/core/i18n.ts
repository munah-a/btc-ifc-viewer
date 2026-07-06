/**
 * Internationalization core (constraint C7 — EN + DE, launch-blocking).
 *
 * A small, hand-rolled, typed message catalog — no i18next/next-intl (the app
 * is a vanilla Vite SPA and C1 forbids runtime CDN / heavy deps). Design:
 *
 * - `en` is the source of truth: its key set defines `MessageKey`, and `de`
 *   is typed `Record<MessageKey, string>` so a missing German key is a compile
 *   error, not a runtime surprise.
 * - `t(key, params?)` interpolates `{name}`-style params and returns the string
 *   for the current language. A `de` value is trusted at compile time; if one
 *   is somehow empty it falls back to `en` and warns in dev only.
 * - `setLanguage` / `getLanguage` persist to localStorage, set `<html lang>`,
 *   re-hydrate `[data-i18n]` DOM, and notify subscribers so JS-rendered panels
 *   re-render. Default = EN (existing e2e assert English status strings).
 * - `hydrateI18n(root)` walks `[data-i18n]` / `[data-i18n-attr]` elements (same
 *   pattern as `hydrateIcons`) and fills text / attributes from the catalog.
 * - `formatDate` / `formatDateTime` / `formatNumber` apply the brand voice
 *   (Swiss DD.MM.YYYY; CHF `1'234.50` apostrophe grouping) via Intl. Pure and
 *   unit-tested; safe to import under Node (no DOM access at module scope).
 *
 * Brand voice (design-system README §2): sentence case everywhere (no Title
 * Case buttons), `—` for empty values, no idiom/wordplay, DE = same register.
 *
 * Translated strings are static catalog constants (never user data), so they
 * are NOT an A1 injection surface. Any *value* interpolated into innerHTML is
 * still escaped by the caller as before.
 */

export type Language = 'en' | 'de';

export const SUPPORTED_LANGUAGES: readonly Language[] = ['en', 'de'];
export const DEFAULT_LANGUAGE: Language = 'en';

const LANGUAGE_STORAGE_KEY = 'bim_for_field_viewer_lang_v1';

/**
 * English catalog — the source of truth. Keys are dotted-namespace strings
 * grouped by area (status.*, toast.*, confirm.*, load.*, panel.*, empty.*,
 * label.*, tree.*, issue.*). `{param}` placeholders are filled by `t()`.
 *
 * Enum-like display maps (visual styles, issue priority/status, nav view) are
 * translated for DISPLAY only — the stored/persisted value stays the English
 * enum (see persistence.ts) so imports and BCF back-compat are unaffected.
 */
export const en = {
  // ---- status bar (setStatus) ----
  'status.initializing': 'Initializing BTC IFC Viewer…',
  'status.ready': 'Ready — load IFC model(s)',
  'status.initFailed': 'Initialization failed: {error}',
  'status.sectionsCleared': 'Sections cleared',
  'status.measurementsCleared': 'Measurements cleared',
  'status.searchCleared': 'Search cleared',
  'status.noSelectionToIsolate': 'No selection to isolate',
  'status.selectionIsolated': 'Selection isolated',
  'status.noSelectionToHide': 'No selection to hide',
  'status.selectionHidden': 'Selection hidden',
  'status.visibilityReset': 'Visibility reset',
  'status.xrayEnabled': 'X-ray enabled',
  'status.xrayDisabled': 'X-ray disabled',
  'status.edgesEnabled': 'Edge overlay enabled',
  'status.edgesDisabled': 'Edge overlay disabled',
  'status.issuePinEnabled': 'Issue pin mode active',
  'status.issuePinDisabled': 'Issue pin mode disabled',
  'status.gizmoUpdated': 'Gizmo updated: {name}',
  'status.modelUnloaded': 'Model unloaded: {name}',
  'status.modelLoaded': 'Model loaded: {name}',
  'status.gridEnabled': 'Grid enabled',
  'status.gridHidden': 'Grid hidden',
  'status.backgroundSet': 'Background colour set to {color}',
  'status.visualStyle': 'Visual style: {style}',
  'status.modelSyncUpdated': 'Model synchronization updated. Reselect the element if needed.',
  'status.contextFailed': '{context} failed: {message}',
  'status.modelIndexNotReady': 'Model index not ready yet',
  'status.selectedFullModel': 'Selected full model ({count} elements)',
  'status.modelShown': 'Shown: {name}',
  'status.modelHidden': 'Hidden: {name}',
  'status.showModelBeforeGizmo': 'Show the model before enabling the gizmo',
  'status.gizmoDetached': 'Model gizmo detached',
  'status.gizmoActive': 'Model gizmo active: {name} (W move, E rotate, R reset transform)',
  'status.transformUpdated': 'Updated transform: {name}',
  'status.transformReset': 'Reset transform: {name}',
  'status.noElementsForLevel': 'No elements found for level {level}',
  'status.levelUnavailable': 'Level {level} is not available for the loaded model IDs',
  'status.isolatedLevel': 'Isolated {level}',
  'status.noClassInLevel': 'No {class} elements found in {level}',
  'status.classInLevelUnavailable': '{class} in {level} is not available for the loaded model IDs',
  'status.isolatedClassInLevel': 'Isolated {class} in {level}',
  'status.noFiltersSelected': 'No filters selected. Showing all elements',
  'status.filtersNoCommon': 'Selected filters have no elements in common',
  'status.filtersApplied': 'Filters applied',
  'status.searchMatches': 'Search found {count} matches',
  'status.batchLoaded': 'Loaded {count} IFC models',
  'status.modelLoadedOk': 'Model loaded successfully',
  'status.modelAlreadyLoading': 'A model is already loading. Please wait…',
  'status.loadingModel': 'Loading model…',
  'status.loadFailed': 'Failed to load IFC: {message}',
  'status.selectionModeSingle': 'Single-selection mode',
  'status.selectionModeMulti': 'Multi-selection mode',
  'status.measureDisabled': 'Measure mode disabled',
  'status.measureEnabled': '{mode} measurement enabled',
  'status.noModelToSection': 'No model to section',
  'status.sectionPlaneAdded': 'Section plane added',
  'status.sectionBoxCreated': 'Section box created',
  'status.issuePointCaptured': 'Issue point captured. Fill in the issue form and create the issue',
  'status.elementsSelected': '{count} element(s) selected',
  'status.enterViewpointName': 'Enter a viewpoint name',
  'status.viewpointSaved': 'Saved viewpoint: {name}',
  'status.selectViewpointFirst': 'Select a viewpoint first',
  'status.viewpointNotFound': 'Selected viewpoint not found',
  'status.viewpointApplied': 'Applied viewpoint: {name}',
  'status.viewpointDeleted': 'Viewpoint deleted',
  'status.issueTitleRequired': 'Issue title is required',
  'status.issueNeedsContext': 'Select element(s) or use issue pin mode to capture a point',
  'status.issueCreated': 'Issue created',
  'status.selectedIssue': 'Selected issue: {title}',
  'status.selectIssueFirst': 'Select an issue first',
  'status.issueDeleted': 'Issue deleted',
  'status.commentEmpty': 'Comment cannot be empty',
  'status.commentAdded': 'Comment added',
  'status.screenshotExported': 'Screenshot exported',
  'status.screenshotFailed': 'Screenshot export failed: {error}',
  'status.stateExported': 'Viewer data exported',
  'status.stateImported': 'Viewer data imported',
  'status.importFailed': 'Import failed: {error}',
  'status.savedWithoutThumbnails': 'Local state saved without thumbnails (storage quota)',
  'status.persistFailed': 'Unable to persist local state: {error}',
  'status.restoreFailed': 'Failed to restore local state: {error}',
  'status.toolCanceled': 'Active tool canceled',
  'status.gizmoModeRotate': 'Gizmo mode: rotate',
  'status.gizmoModeTranslate': 'Gizmo mode: translate',
  'status.gizmoTransformReset': 'Gizmo: model transform reset',
  'status.languageChanged': 'Language: English',

  // ---- toasts ----
  'toast.onlyIfc': 'Only IFC files are supported',
  'toast.filtersNothingHidden': 'No elements match the selected class and level filters — nothing was hidden',

  // ---- confirm dialogs ----
  'confirm.unloadModel': 'Unload {name}? Saved issues and viewpoints stay, but its elements leave the viewer.',
  'confirm.deleteViewpoint': 'Delete this viewpoint? This cannot be undone.',
  'confirm.deleteIssue': 'Delete this issue? This cannot be undone.',
  'confirm.ok': 'Confirm',
  'confirm.cancel': 'Cancel',
  'confirm.unload': 'Unload',

  // ---- loading / progress ----
  'load.parsingBatch': 'Parsing IFC {index}/{total}…',
  'load.parsing': 'Parsing IFC…',
  'load.converting': 'Converting IFC to fragments…',
  'load.buildingIndex': 'Building spatial index…',

  // ---- panel titles (tab activation) ----
  'panel.explorer': 'Explorer',
  'panel.models': 'Federated models',
  'panel.properties': 'Properties',
  'panel.viewpoints': 'Viewpoints',
  'panel.issues': 'Issues',
  'panel.help': 'Help',

  // ---- empty states / placeholders (rendered from JS) ----
  'empty.noClasses': 'No classes detected',
  'empty.noLevels': 'No levels detected',
  'empty.noModelLoaded': 'No model loaded',
  'empty.noModelsYet': 'No models loaded yet',
  'empty.noMatches': 'No matches',
  'empty.noViewpoints': 'No viewpoints yet. Save one before creating issues for better traceability.',
  'empty.noIssues': 'No issues yet.',
  'empty.selectIssueForComments': 'Select an issue to view comments',
  'empty.noComments': 'No comments',

  // ---- dynamic labels / counters ----
  'label.elementFallback': 'Element {id}',
  'label.itemFallback': 'Item',
  'label.elements': '{count} elements',
  'label.models': '{count} models',
  'label.selected': '{count} selected',
  'label.visible': '{count} visible',
  'label.percent': '{value}%',
  'label.viewOrbit': 'Orbit · perspective',
  'label.viewFront': 'Front · orthographic',
  'label.viewTop': 'Top · orthographic',
  'label.measureAreaHint': 'Area — click points, double-click to close',
  'label.measureLengthHint': 'Length — click two points',
  'label.sectionAxis': 'Section {axis}',
  'label.issueMarker': '{title} ({status})',
  'label.loadInfo': 'Loaded in {seconds}s | {size}MB',

  // ---- model-browser tree (passed into the pure builder) ----
  'tree.hidden': 'Hidden',
  'tree.building': 'Building model tree…',
  'tree.moreNodes': '{count} more nodes',
  'tree.moreElements': '{count} more elements',
  'tree.moreLevels': '{count} more levels',
  'tree.noElements': 'No elements',
  'tree.noClasses': 'No classes',
  'tree.default': 'Default',
  'tree.levels': 'Levels',
  'tree.levelsShort': '{count} lvls',
  'tree.spatialStructure': 'Spatial structure',
  'tree.noLevelsDetected': 'No levels detected',
  'tree.noSpatialData': 'No spatial tree data',

  // ---- federation panel (passed into the pure builder) ----
  'fed.show': 'Show',
  'fed.hide': 'Hide',
  'fed.noStoreys': 'No storeys found',
  'fed.opacity': 'Opacity',
  'fed.offsetXyz': 'Offset XYZ',
  'fed.rotationXyz': 'Rotation XYZ (deg)',
  'fed.metaLine': '{id} | {count} elements | {size}',
  'fed.select': 'Select',
  'fed.gizmo': 'Gizmo',
  'fed.fit': 'Fit',
  'fed.reset': 'Reset',
  'fed.unload': 'Unload',
  'fed.selectFullModel': 'Select full model',
  'fed.fitCamera': 'Fit camera to model',
  'fed.unloadTitle': 'Unload model and free its memory',
  'fed.isolateLevel': 'Isolate level {level}',
  'fed.levelsCount': 'Levels ({count})',

  // ---- issue list (passed into the pure builder) ----
  'issue.linked': '{count} linked · {models} model(s)',
  'issue.noLink': 'No element link',
  'issue.comments': 'Comments',

  // ---- viewpoint list (passed into the pure builder) ----
  'vp.apply': 'Apply',
  'vp.deleteTitle': 'Delete viewpoint',

  // ---- enum display maps (display only; stored value stays English) ----
  'style.basic': 'Basic',
  'style.pen': 'Pen',
  'style.colorPen': 'Colour pen',
  'style.colorShadows': 'Colour shadows',
  'style.colorPenShadows': 'Colour pen shadows',
  'measure.length': 'Length',
  'measure.area': 'Area',

  // ---- static shell (index.html data-i18n / data-i18n-attr) ----
  'shell.skipNav': 'Skip to viewer',
  'shell.eyebrow': 'IFC viewer',
  'shell.loadIfc': 'Load IFC',
  'shell.exportScreenshot': 'Export screenshot',
  'shell.exportState': 'Export state',
  'shell.importState': 'Import state',
  'shell.toggleTheme': 'Toggle light / dark theme',
  'shell.toggleLanguage': 'Switch language (EN / DE)',
  'shell.collapsePanel': 'Collapse panel',
  'shell.expandPanel': 'Expand panel',
  // tool rail
  'shell.selectSingle': 'Single select',
  'shell.selectMulti': 'Multi select (M)',
  'shell.isolate': 'Isolate selection',
  'shell.hide': 'Hide selection',
  'shell.resetVisibility': 'Show all / reset visibility',
  'shell.measureLength': 'Measure length (L)',
  'shell.measureArea': 'Measure area (A)',
  'shell.clearMeasurements': 'Clear measurements',
  'shell.sectionX': 'Section X',
  'shell.sectionY': 'Section Y',
  'shell.sectionZ': 'Section Z',
  'shell.sectionBox': 'Section box',
  'shell.clearSections': 'Clear sections',
  'shell.xray': 'X-ray (X)',
  'shell.edges': 'Edges (E)',
  'shell.grid': 'Grid (G)',
  'shell.issuePinMode': 'Issue pin mode (I)',
  // empty state
  'shell.emptyTitle': 'Load IFC models',
  'shell.emptyBody': 'Drop one or more IFC files in the viewport, or load them from disk. Federated models stay aligned.',
  'shell.emptyAction': 'Load IFC files',
  // loading
  'shell.retry': 'Retry',
  'shell.dismiss': 'Dismiss',
  // view controls / nav
  'shell.fitAll': 'Fit all (F)',
  'shell.orbitHome': 'Orbit / home view',
  'shell.frontView': 'Front view',
  'shell.topView': 'Top view',
  'shell.navOrbit': 'Orbit',
  'shell.navPlan': 'Plan',
  'shell.navWalk': 'Walk',
  // measure hint / section slider / selection chip
  'shell.measureHintDefault': 'Length — click two points',
  'shell.cancel': 'Esc',
  'shell.sectionLabel': 'Section',
  'shell.sectionPosAria': 'Section plane position',
  'shell.clearSection': 'Clear section',
  'shell.clearSelection': 'Clear selection',
  'shell.issuePinHint': 'Issue pin mode active: click the model to place an issue anchor',
  // splitter
  'shell.resizePanel': 'Resize panel',
  // panel header (default title) + explorer
  'shell.panelTitle': 'Explorer',
  'shell.searchPlaceholder': 'Search elements, GlobalId…',
  'shell.searchAria': 'Search elements',
  'shell.clearSearch': 'Clear search',
  'shell.classFilters': 'Class filters',
  'shell.levelFilters': 'Level filters',
  'shell.searchResults': 'Search results',
  'shell.spatialTree': 'Spatial tree',
  // models tab
  'shell.federatedModels': 'Federated models',
  // properties tab
  'shell.propsEmpty': 'Select an element in the viewport or the spatial tree to inspect its properties.',
  'shell.attributes': 'Attributes',
  'shell.propType': 'Type',
  'shell.propName': 'Name',
  'shell.propGlobalId': 'GlobalId',
  'shell.propStorey': 'Storey',
  'shell.propDescription': 'Description',
  'shell.filterProps': 'Filter properties',
  'shell.propsIsolate': 'Isolate',
  'shell.propsHide': 'Hide',
  // viewpoints tab
  'shell.viewpointName': 'Viewpoint name',
  'shell.saveViewpoint': 'Save current viewpoint',
  // issues tab
  'shell.newIssue': 'New issue',
  'shell.issueTitle': 'Issue title…',
  'shell.issueTitleAria': 'Issue title',
  'shell.issueDescription': 'Description (optional)',
  'shell.issueDescriptionAria': 'Issue description',
  'shell.priority': 'Priority',
  'shell.status': 'Status',
  'shell.prioCritical': 'Critical',
  'shell.prioHigh': 'High',
  'shell.prioMedium': 'Medium',
  'shell.prioLow': 'Low',
  'shell.statusOpen': 'Open',
  'shell.statusInProgress': 'In progress',
  'shell.statusResolved': 'Resolved',
  'shell.statusClosed': 'Closed',
  'shell.assignee': 'Assignee (optional)',
  'shell.assigneeAria': 'Assignee',
  'shell.createIssue': 'Create issue',
  'shell.comments': 'Comments',
  'shell.addComment': 'Add comment',
  'shell.deleteIssue': 'Delete selected issue',
  // help tab
  'shell.keyboardShortcuts': 'Keyboard shortcuts',
  'shell.scFit': 'Fit model',
  'shell.scModes': 'Orbit / plan / walk mode',
  'shell.scMultiSelect': 'Toggle multi-select',
  'shell.scLength': 'Length measurement',
  'shell.scArea': 'Area measurement',
  'shell.scGrid': 'Toggle grid',
  'shell.scXray': 'Toggle X-ray',
  'shell.scEdges': 'Toggle edges',
  'shell.scIssuePin': 'Toggle issue pin mode',
  'shell.scEscape': 'Cancel active tool',
  'shell.tips': 'Tips',
  'shell.tipsBody': 'Use class and level filters in the Explorer to isolate scope. Save viewpoints before creating issues for better traceability. Export state to persist a standalone session.',
  // status bar
  'shell.statusReady': 'Ready',
  'shell.viewLabel': 'Orbit · perspective',
  // tab strip tooltips
  'shell.tabExplorer': 'Explorer',
  'shell.tabModels': 'Models',
  'shell.tabProperties': 'Properties',
  'shell.tabViewpoints': 'Viewpoints',
  'shell.tabIssues': 'Issues',
  'shell.tabHelp': 'Help',
  // mobile
  'shell.fitModel': 'Fit model',
  'shell.mobileTree': 'Tree',
  'shell.mobileOrbit': 'Orbit',
  'shell.mobileSection': 'Section',
  'shell.mobileMeasure': 'Measure',
  'shell.mobileMore': 'More',
  'shell.viewSettings': 'View settings',
  'shell.closeSheet': 'Close',
  'shell.close': 'Close',
} as const;

/** The canonical key set — derived from the English catalog. */
export type MessageKey = keyof typeof en;

/** Placeholder tokens supplied to interpolate into a message. */
export type MessageParams = Record<string, string | number>;

/**
 * German catalog. Typed `Record<MessageKey, string>` so every English key MUST
 * have a German value (a gap is a compile error). Brand voice: sentence case,
 * Swiss/German engineering register, no idiom. IFC class names (IfcWall…) stay
 * untranslated by the caller; only friendly labels are translated.
 */
export const de: Record<MessageKey, string> = {
  // ---- status bar ----
  'status.initializing': 'BTC IFC Viewer wird initialisiert…',
  'status.ready': 'Bereit — IFC-Modell(e) laden',
  'status.initFailed': 'Initialisierung fehlgeschlagen: {error}',
  'status.sectionsCleared': 'Schnitte entfernt',
  'status.measurementsCleared': 'Messungen entfernt',
  'status.searchCleared': 'Suche zurückgesetzt',
  'status.noSelectionToIsolate': 'Keine Auswahl zum Isolieren',
  'status.selectionIsolated': 'Auswahl isoliert',
  'status.noSelectionToHide': 'Keine Auswahl zum Ausblenden',
  'status.selectionHidden': 'Auswahl ausgeblendet',
  'status.visibilityReset': 'Sichtbarkeit zurückgesetzt',
  'status.xrayEnabled': 'Röntgenansicht aktiviert',
  'status.xrayDisabled': 'Röntgenansicht deaktiviert',
  'status.edgesEnabled': 'Kantenüberlagerung aktiviert',
  'status.edgesDisabled': 'Kantenüberlagerung deaktiviert',
  'status.issuePinEnabled': 'Aufgaben-Pin-Modus aktiv',
  'status.issuePinDisabled': 'Aufgaben-Pin-Modus deaktiviert',
  'status.gizmoUpdated': 'Gizmo aktualisiert: {name}',
  'status.modelUnloaded': 'Modell entladen: {name}',
  'status.modelLoaded': 'Modell geladen: {name}',
  'status.gridEnabled': 'Raster aktiviert',
  'status.gridHidden': 'Raster ausgeblendet',
  'status.backgroundSet': 'Hintergrundfarbe auf {color} gesetzt',
  'status.visualStyle': 'Darstellungsstil: {style}',
  'status.modelSyncUpdated': 'Modellsynchronisierung aktualisiert. Element bei Bedarf neu auswählen.',
  'status.contextFailed': '{context} fehlgeschlagen: {message}',
  'status.modelIndexNotReady': 'Modellindex noch nicht bereit',
  'status.selectedFullModel': 'Gesamtes Modell ausgewählt ({count} Elemente)',
  'status.modelShown': 'Eingeblendet: {name}',
  'status.modelHidden': 'Ausgeblendet: {name}',
  'status.showModelBeforeGizmo': 'Modell einblenden, bevor das Gizmo aktiviert wird',
  'status.gizmoDetached': 'Modell-Gizmo gelöst',
  'status.gizmoActive': 'Modell-Gizmo aktiv: {name} (W verschieben, E drehen, R Transformation zurücksetzen)',
  'status.transformUpdated': 'Transformation aktualisiert: {name}',
  'status.transformReset': 'Transformation zurückgesetzt: {name}',
  'status.noElementsForLevel': 'Keine Elemente für Geschoss {level} gefunden',
  'status.levelUnavailable': 'Geschoss {level} ist für die geladenen Modell-IDs nicht verfügbar',
  'status.isolatedLevel': '{level} isoliert',
  'status.noClassInLevel': 'Keine {class}-Elemente in {level} gefunden',
  'status.classInLevelUnavailable': '{class} in {level} ist für die geladenen Modell-IDs nicht verfügbar',
  'status.isolatedClassInLevel': '{class} in {level} isoliert',
  'status.noFiltersSelected': 'Keine Filter ausgewählt. Alle Elemente werden angezeigt',
  'status.filtersNoCommon': 'Ausgewählte Filter haben keine gemeinsamen Elemente',
  'status.filtersApplied': 'Filter angewendet',
  'status.searchMatches': 'Suche fand {count} Treffer',
  'status.batchLoaded': '{count} IFC-Modelle geladen',
  'status.modelLoadedOk': 'Modell erfolgreich geladen',
  'status.modelAlreadyLoading': 'Ein Modell wird bereits geladen. Bitte warten…',
  'status.loadingModel': 'Modell wird geladen…',
  'status.loadFailed': 'IFC konnte nicht geladen werden: {message}',
  'status.selectionModeSingle': 'Einzelauswahl-Modus',
  'status.selectionModeMulti': 'Mehrfachauswahl-Modus',
  'status.measureDisabled': 'Messmodus deaktiviert',
  'status.measureEnabled': '{mode}-Messung aktiviert',
  'status.noModelToSection': 'Kein Modell zum Schneiden',
  'status.sectionPlaneAdded': 'Schnittebene hinzugefügt',
  'status.sectionBoxCreated': 'Schnittbox erstellt',
  'status.issuePointCaptured': 'Aufgabenpunkt erfasst. Aufgabenformular ausfüllen und Aufgabe erstellen',
  'status.elementsSelected': '{count} Element(e) ausgewählt',
  'status.enterViewpointName': 'Namen für den Blickpunkt eingeben',
  'status.viewpointSaved': 'Blickpunkt gespeichert: {name}',
  'status.selectViewpointFirst': 'Zuerst einen Blickpunkt auswählen',
  'status.viewpointNotFound': 'Ausgewählter Blickpunkt nicht gefunden',
  'status.viewpointApplied': 'Blickpunkt angewendet: {name}',
  'status.viewpointDeleted': 'Blickpunkt gelöscht',
  'status.issueTitleRequired': 'Aufgabentitel ist erforderlich',
  'status.issueNeedsContext': 'Element(e) auswählen oder den Aufgaben-Pin-Modus nutzen, um einen Punkt zu erfassen',
  'status.issueCreated': 'Aufgabe erstellt',
  'status.selectedIssue': 'Ausgewählte Aufgabe: {title}',
  'status.selectIssueFirst': 'Zuerst eine Aufgabe auswählen',
  'status.issueDeleted': 'Aufgabe gelöscht',
  'status.commentEmpty': 'Kommentar darf nicht leer sein',
  'status.commentAdded': 'Kommentar hinzugefügt',
  'status.screenshotExported': 'Screenshot exportiert',
  'status.screenshotFailed': 'Screenshot-Export fehlgeschlagen: {error}',
  'status.stateExported': 'Viewer-Daten exportiert',
  'status.stateImported': 'Viewer-Daten importiert',
  'status.importFailed': 'Import fehlgeschlagen: {error}',
  'status.savedWithoutThumbnails': 'Lokaler Zustand ohne Miniaturbilder gespeichert (Speicherlimit)',
  'status.persistFailed': 'Lokaler Zustand konnte nicht gespeichert werden: {error}',
  'status.restoreFailed': 'Lokaler Zustand konnte nicht wiederhergestellt werden: {error}',
  'status.toolCanceled': 'Aktives Werkzeug abgebrochen',
  'status.gizmoModeRotate': 'Gizmo-Modus: drehen',
  'status.gizmoModeTranslate': 'Gizmo-Modus: verschieben',
  'status.gizmoTransformReset': 'Gizmo: Modelltransformation zurückgesetzt',
  'status.languageChanged': 'Sprache: Deutsch',

  // ---- toasts ----
  'toast.onlyIfc': 'Es werden nur IFC-Dateien unterstützt',
  'toast.filtersNothingHidden': 'Keine Elemente entsprechen den gewählten Klassen- und Geschossfiltern — nichts wurde ausgeblendet',

  // ---- confirm dialogs ----
  'confirm.unloadModel': '{name} entladen? Gespeicherte Aufgaben und Blickpunkte bleiben erhalten, aber die Elemente verlassen den Viewer.',
  'confirm.deleteViewpoint': 'Diesen Blickpunkt löschen? Dies kann nicht rückgängig gemacht werden.',
  'confirm.deleteIssue': 'Diese Aufgabe löschen? Dies kann nicht rückgängig gemacht werden.',
  'confirm.ok': 'Bestätigen',
  'confirm.cancel': 'Abbrechen',
  'confirm.unload': 'Entladen',

  // ---- loading / progress ----
  'load.parsingBatch': 'IFC {index}/{total} wird gelesen…',
  'load.parsing': 'IFC wird gelesen…',
  'load.converting': 'IFC wird in Fragmente konvertiert…',
  'load.buildingIndex': 'Räumlicher Index wird aufgebaut…',

  // ---- panel titles ----
  'panel.explorer': 'Explorer',
  'panel.models': 'Föderierte Modelle',
  'panel.properties': 'Eigenschaften',
  'panel.viewpoints': 'Blickpunkte',
  'panel.issues': 'Aufgaben',
  'panel.help': 'Hilfe',

  // ---- empty states / placeholders ----
  'empty.noClasses': 'Keine Klassen erkannt',
  'empty.noLevels': 'Keine Geschosse erkannt',
  'empty.noModelLoaded': 'Kein Modell geladen',
  'empty.noModelsYet': 'Noch keine Modelle geladen',
  'empty.noMatches': 'Keine Treffer',
  'empty.noViewpoints': 'Noch keine Blickpunkte. Speichern Sie einen, bevor Sie Aufgaben erstellen, für bessere Nachvollziehbarkeit.',
  'empty.noIssues': 'Noch keine Aufgaben.',
  'empty.selectIssueForComments': 'Aufgabe auswählen, um Kommentare anzuzeigen',
  'empty.noComments': 'Keine Kommentare',

  // ---- dynamic labels / counters ----
  'label.elementFallback': 'Element {id}',
  'label.itemFallback': 'Element',
  'label.elements': '{count} Elemente',
  'label.models': '{count} Modelle',
  'label.selected': '{count} ausgewählt',
  'label.visible': '{count} sichtbar',
  'label.percent': '{value}%',
  'label.viewOrbit': 'Orbit · perspektivisch',
  'label.viewFront': 'Vorne · orthografisch',
  'label.viewTop': 'Oben · orthografisch',
  'label.measureAreaHint': 'Fläche — Punkte anklicken, Doppelklick zum Schliessen',
  'label.measureLengthHint': 'Länge — zwei Punkte anklicken',
  'label.sectionAxis': 'Schnitt {axis}',
  'label.issueMarker': '{title} ({status})',
  'label.loadInfo': 'Geladen in {seconds}s | {size}MB',

  // ---- model-browser tree ----
  'tree.hidden': 'Ausgeblendet',
  'tree.building': 'Modellbaum wird aufgebaut…',
  'tree.moreNodes': '{count} weitere Knoten',
  'tree.moreElements': '{count} weitere Elemente',
  'tree.moreLevels': '{count} weitere Geschosse',
  'tree.noElements': 'Keine Elemente',
  'tree.noClasses': 'Keine Klassen',
  'tree.default': 'Standard',
  'tree.levels': 'Geschosse',
  'tree.levelsShort': '{count} Gesch.',
  'tree.spatialStructure': 'Räumliche Struktur',
  'tree.noLevelsDetected': 'Keine Geschosse erkannt',
  'tree.noSpatialData': 'Keine räumlichen Baumdaten',

  // ---- federation panel ----
  'fed.show': 'Einblenden',
  'fed.hide': 'Ausblenden',
  'fed.noStoreys': 'Keine Geschosse gefunden',
  'fed.opacity': 'Deckkraft',
  'fed.offsetXyz': 'Versatz XYZ',
  'fed.rotationXyz': 'Rotation XYZ (Grad)',
  'fed.metaLine': '{id} | {count} Elemente | {size}',
  'fed.select': 'Auswählen',
  'fed.gizmo': 'Gizmo',
  'fed.fit': 'Einpassen',
  'fed.reset': 'Zurücksetzen',
  'fed.unload': 'Entladen',
  'fed.selectFullModel': 'Gesamtes Modell auswählen',
  'fed.fitCamera': 'Kamera an Modell einpassen',
  'fed.unloadTitle': 'Modell entladen und Speicher freigeben',
  'fed.isolateLevel': 'Geschoss {level} isolieren',
  'fed.levelsCount': 'Geschosse ({count})',

  // ---- issue list ----
  'issue.linked': '{count} verknüpft · {models} Modell(e)',
  'issue.noLink': 'Keine Elementverknüpfung',
  'issue.comments': 'Kommentare',

  // ---- viewpoint list ----
  'vp.apply': 'Anwenden',
  'vp.deleteTitle': 'Blickpunkt löschen',

  // ---- enum display maps ----
  'style.basic': 'Einfach',
  'style.pen': 'Stift',
  'style.colorPen': 'Farbstift',
  'style.colorShadows': 'Farbschatten',
  'style.colorPenShadows': 'Farbstift mit Schatten',
  'measure.length': 'Länge',
  'measure.area': 'Fläche',

  // ---- static shell ----
  'shell.skipNav': 'Zum Viewer springen',
  'shell.eyebrow': 'IFC-Viewer',
  'shell.loadIfc': 'IFC laden',
  'shell.exportScreenshot': 'Screenshot exportieren',
  'shell.exportState': 'Zustand exportieren',
  'shell.importState': 'Zustand importieren',
  'shell.toggleTheme': 'Helles / dunkles Design umschalten',
  'shell.toggleLanguage': 'Sprache wechseln (EN / DE)',
  'shell.collapsePanel': 'Panel einklappen',
  'shell.expandPanel': 'Panel ausklappen',
  // tool rail
  'shell.selectSingle': 'Einzelauswahl',
  'shell.selectMulti': 'Mehrfachauswahl (M)',
  'shell.isolate': 'Auswahl isolieren',
  'shell.hide': 'Auswahl ausblenden',
  'shell.resetVisibility': 'Alle einblenden / Sichtbarkeit zurücksetzen',
  'shell.measureLength': 'Länge messen (L)',
  'shell.measureArea': 'Fläche messen (A)',
  'shell.clearMeasurements': 'Messungen entfernen',
  'shell.sectionX': 'Schnitt X',
  'shell.sectionY': 'Schnitt Y',
  'shell.sectionZ': 'Schnitt Z',
  'shell.sectionBox': 'Schnittbox',
  'shell.clearSections': 'Schnitte entfernen',
  'shell.xray': 'Röntgenansicht (X)',
  'shell.edges': 'Kanten (E)',
  'shell.grid': 'Raster (G)',
  'shell.issuePinMode': 'Aufgaben-Pin-Modus (I)',
  // empty state
  'shell.emptyTitle': 'IFC-Modelle laden',
  'shell.emptyBody': 'Eine oder mehrere IFC-Dateien in den Viewport ziehen oder von der Festplatte laden. Föderierte Modelle bleiben ausgerichtet.',
  'shell.emptyAction': 'IFC-Dateien laden',
  // loading
  'shell.retry': 'Erneut versuchen',
  'shell.dismiss': 'Verwerfen',
  // view controls / nav
  'shell.fitAll': 'Alles einpassen (F)',
  'shell.orbitHome': 'Orbit / Startansicht',
  'shell.frontView': 'Vorderansicht',
  'shell.topView': 'Draufsicht',
  'shell.navOrbit': 'Orbit',
  'shell.navPlan': 'Plan',
  'shell.navWalk': 'Gehen',
  // measure hint / section slider / selection chip
  'shell.measureHintDefault': 'Länge — zwei Punkte anklicken',
  'shell.cancel': 'Esc',
  'shell.sectionLabel': 'Schnitt',
  'shell.sectionPosAria': 'Position der Schnittebene',
  'shell.clearSection': 'Schnitt entfernen',
  'shell.clearSelection': 'Auswahl aufheben',
  'shell.issuePinHint': 'Aufgaben-Pin-Modus aktiv: Modell anklicken, um einen Aufgabenanker zu setzen',
  // splitter
  'shell.resizePanel': 'Panelgrösse ändern',
  // panel header + explorer
  'shell.panelTitle': 'Explorer',
  'shell.searchPlaceholder': 'Elemente, GlobalId suchen…',
  'shell.searchAria': 'Elemente suchen',
  'shell.clearSearch': 'Suche zurücksetzen',
  'shell.classFilters': 'Klassenfilter',
  'shell.levelFilters': 'Geschossfilter',
  'shell.searchResults': 'Suchergebnisse',
  'shell.spatialTree': 'Räumlicher Baum',
  // models tab
  'shell.federatedModels': 'Föderierte Modelle',
  // properties tab
  'shell.propsEmpty': 'Ein Element im Viewport oder im räumlichen Baum auswählen, um seine Eigenschaften zu prüfen.',
  'shell.attributes': 'Attribute',
  'shell.propType': 'Typ',
  'shell.propName': 'Name',
  'shell.propGlobalId': 'GlobalId',
  'shell.propStorey': 'Geschoss',
  'shell.propDescription': 'Beschreibung',
  'shell.filterProps': 'Eigenschaften filtern',
  'shell.propsIsolate': 'Isolieren',
  'shell.propsHide': 'Ausblenden',
  // viewpoints tab
  'shell.viewpointName': 'Name des Blickpunkts',
  'shell.saveViewpoint': 'Aktuellen Blickpunkt speichern',
  // issues tab
  'shell.newIssue': 'Neue Aufgabe',
  'shell.issueTitle': 'Aufgabentitel…',
  'shell.issueTitleAria': 'Aufgabentitel',
  'shell.issueDescription': 'Beschreibung (optional)',
  'shell.issueDescriptionAria': 'Aufgabenbeschreibung',
  'shell.priority': 'Priorität',
  'shell.status': 'Status',
  'shell.prioCritical': 'Kritisch',
  'shell.prioHigh': 'Hoch',
  'shell.prioMedium': 'Mittel',
  'shell.prioLow': 'Niedrig',
  'shell.statusOpen': 'Offen',
  'shell.statusInProgress': 'In Bearbeitung',
  'shell.statusResolved': 'Gelöst',
  'shell.statusClosed': 'Geschlossen',
  'shell.assignee': 'Zuständig (optional)',
  'shell.assigneeAria': 'Zuständig',
  'shell.createIssue': 'Aufgabe erstellen',
  'shell.comments': 'Kommentare',
  'shell.addComment': 'Kommentar hinzufügen',
  'shell.deleteIssue': 'Ausgewählte Aufgabe löschen',
  // help tab
  'shell.keyboardShortcuts': 'Tastenkürzel',
  'shell.scFit': 'Modell einpassen',
  'shell.scModes': 'Orbit- / Plan- / Gehen-Modus',
  'shell.scMultiSelect': 'Mehrfachauswahl umschalten',
  'shell.scLength': 'Längenmessung',
  'shell.scArea': 'Flächenmessung',
  'shell.scGrid': 'Raster umschalten',
  'shell.scXray': 'Röntgenansicht umschalten',
  'shell.scEdges': 'Kanten umschalten',
  'shell.scIssuePin': 'Aufgaben-Pin-Modus umschalten',
  'shell.scEscape': 'Aktives Werkzeug abbrechen',
  'shell.tips': 'Tipps',
  'shell.tipsBody': 'Klassen- und Geschossfilter im Explorer nutzen, um den Umfang zu isolieren. Blickpunkte vor dem Erstellen von Aufgaben speichern, für bessere Nachvollziehbarkeit. Zustand exportieren, um eine eigenständige Sitzung zu sichern.',
  // status bar
  'shell.statusReady': 'Bereit',
  'shell.viewLabel': 'Orbit · perspektivisch',
  // tab strip tooltips
  'shell.tabExplorer': 'Explorer',
  'shell.tabModels': 'Modelle',
  'shell.tabProperties': 'Eigenschaften',
  'shell.tabViewpoints': 'Blickpunkte',
  'shell.tabIssues': 'Aufgaben',
  'shell.tabHelp': 'Hilfe',
  // mobile
  'shell.fitModel': 'Modell einpassen',
  'shell.mobileTree': 'Baum',
  'shell.mobileOrbit': 'Orbit',
  'shell.mobileSection': 'Schnitt',
  'shell.mobileMeasure': 'Messen',
  'shell.mobileMore': 'Mehr',
  'shell.viewSettings': 'Ansichtseinstellungen',
  'shell.closeSheet': 'Schliessen',
  'shell.close': 'Schliessen',
};

const catalogs: Record<Language, Record<MessageKey, string>> = { en, de };

const isDev = typeof import.meta !== 'undefined' && Boolean((import.meta as { env?: { DEV?: boolean } }).env?.DEV);

let currentLanguage: Language = DEFAULT_LANGUAGE;

/** Subscribers notified after a language change so JS-rendered panels re-render. */
const subscribers = new Set<(language: Language) => void>();

/** Narrows an arbitrary value to a supported Language, else null. */
export function coerceLanguage(value: unknown): Language | null {
  return typeof value === 'string' && (SUPPORTED_LANGUAGES as readonly string[]).includes(value)
    ? (value as Language)
    : null;
}

function readStoredLanguage(): Language | null {
  if (typeof localStorage === 'undefined') return null;
  try {
    return coerceLanguage(localStorage.getItem(LANGUAGE_STORAGE_KEY));
  } catch {
    return null;
  }
}

/**
 * Resolves the initial language at boot: a persisted choice wins; otherwise
 * the default (EN). Does NOT auto-detect the browser locale — EN is the product
 * default (and the e2e assert English), and the user can switch explicitly.
 * Call once during bootstrap before hydration.
 */
export function initLanguage(): Language {
  const stored = readStoredLanguage();
  currentLanguage = stored ?? DEFAULT_LANGUAGE;
  applyHtmlLang(currentLanguage);
  return currentLanguage;
}

export function getLanguage(): Language {
  return currentLanguage;
}

function applyHtmlLang(language: Language): void {
  if (typeof document !== 'undefined' && document.documentElement) {
    document.documentElement.lang = language;
  }
}

/**
 * Switches the active language: persists it, sets `<html lang>`, re-hydrates the
 * static DOM, and notifies subscribers (panels re-render). No-op if unchanged.
 */
export function setLanguage(language: Language): void {
  const next = coerceLanguage(language) ?? DEFAULT_LANGUAGE;
  if (next === currentLanguage) {
    // Still ensure lang attr is correct (e.g. first explicit set to the default).
    applyHtmlLang(next);
    return;
  }
  currentLanguage = next;
  if (typeof localStorage !== 'undefined') {
    try {
      localStorage.setItem(LANGUAGE_STORAGE_KEY, next);
    } catch {
      // Quota / disabled storage — language still applies for this session.
    }
  }
  applyHtmlLang(next);
  if (typeof document !== 'undefined') hydrateI18n(document);
  for (const subscriber of subscribers) subscriber(next);
}

/** Subscribe to language changes; returns an unsubscribe function. */
export function onLanguageChange(subscriber: (language: Language) => void): () => void {
  subscribers.add(subscriber);
  return () => subscribers.delete(subscriber);
}

/** Fills `{param}` placeholders in a message template. */
function interpolate(template: string, params?: MessageParams): string {
  if (!params) return template;
  return template.replace(/\{(\w+)\}/g, (match, name: string) =>
    Object.prototype.hasOwnProperty.call(params, name) ? String(params[name]) : match,
  );
}

/**
 * Returns the translated, interpolated string for `key` in the current
 * language. Falls back to English if a German value is somehow empty (warns in
 * dev only). `key` is compile-checked against the catalog.
 */
export function t(key: MessageKey, params?: MessageParams): string {
  const table = catalogs[currentLanguage];
  let template = table[key];
  if (template === undefined || template === '') {
    if (isDev && currentLanguage !== 'en') {
      console.warn(`[i18n] missing "${currentLanguage}" translation for "${key}"; using English.`);
    }
    template = en[key];
  }
  return interpolate(template, params);
}

/** Translate in an explicit language (used by the language-name status line). */
export function tIn(language: Language, key: MessageKey, params?: MessageParams): string {
  return interpolate(catalogs[language][key] || en[key], params);
}

/**
 * Hydrates every `[data-i18n]` / `[data-i18n-attr]` element under `root` from
 * the catalog (mirrors `hydrateIcons`). Called at boot and again on every
 * language change so the static shell re-localizes without a reload.
 *
 * - `data-i18n="key"` sets the element's textContent.
 * - `data-i18n-attr="placeholder:key,title:key,aria-label:key"` sets each named
 *   attribute from its key (for inputs / tooltips / aria labels).
 *
 * Keys are validated against the catalog; an unknown key is skipped (and warned
 * in dev) rather than writing the raw key string into the UI.
 */
export function hydrateI18n(root: ParentNode): void {
  root.querySelectorAll<HTMLElement>('[data-i18n]').forEach((el) => {
    const key = el.dataset.i18n;
    if (!key) return;
    if (!(key in en)) {
      if (isDev) console.warn(`[i18n] unknown data-i18n key "${key}"`);
      return;
    }
    el.textContent = t(key as MessageKey);
  });

  root.querySelectorAll<HTMLElement>('[data-i18n-attr]').forEach((el) => {
    const spec = el.dataset.i18nAttr;
    if (!spec) return;
    for (const pair of spec.split(',')) {
      const [attr, key] = pair.split(':').map((s) => s.trim());
      if (!attr || !key) continue;
      if (!(key in en)) {
        if (isDev) console.warn(`[i18n] unknown data-i18n-attr key "${key}"`);
        continue;
      }
      el.setAttribute(attr, t(key as MessageKey));
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Intl formatting helpers (brand voice: Swiss DD.MM.YYYY, CHF 1'234.50).
// Pure — no DOM/storage access — so they are import-safe under Node/Vitest.
// The `-CH` locales give the apostrophe (') group separator the brand wants;
// explicit 2-digit day/month force DD.MM.YYYY regardless of the locale default.
// ─────────────────────────────────────────────────────────────────────────────

const intlLocale = (language: Language): string => (language === 'de' ? 'de-CH' : 'en-CH');

/** Coerces a Date | ISO string | epoch-ms into a valid Date, else null. */
function toDate(value: Date | string | number): Date | null {
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

/**
 * Formats a date as Swiss `DD.MM.YYYY` in the given language (default: current).
 * Returns the brand's em-dash placeholder for an invalid/empty date.
 */
export function formatDate(value: Date | string | number, language: Language = currentLanguage): string {
  const date = toDate(value);
  if (!date) return '—';
  return new Intl.DateTimeFormat(intlLocale(language), {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(date);
}

/**
 * Formats a date + 24-hour time as `DD.MM.YYYY, HH:MM` (used for viewpoint /
 * comment timestamps). Em-dash for an invalid date.
 */
export function formatDateTime(value: Date | string | number, language: Language = currentLanguage): string {
  const date = toDate(value);
  if (!date) return '—';
  return new Intl.DateTimeFormat(intlLocale(language), {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(date);
}

/**
 * Formats a number with the brand's apostrophe grouping and tabular defaults.
 * Non-finite input returns the em-dash placeholder. Pass `opts` (e.g. currency
 * CHF, or a fixed fraction-digit count) for money / precise surfaces.
 */
export function formatNumber(
  value: number,
  language: Language = currentLanguage,
  opts?: Intl.NumberFormatOptions,
): string {
  if (!Number.isFinite(value)) return '—';
  return new Intl.NumberFormat(intlLocale(language), opts).format(value);
}
