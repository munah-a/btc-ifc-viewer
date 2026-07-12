/**
 * Versioned persisted-state schema + validation (AUDIT A7, extracted per the
 * W2.1 plan shape). `normalizePersistedState` turns any untrusted JSON value
 * (localStorage or an imported file) into a fully-defaulted, type-safe state —
 * or null when it isn't a v1 state at all. Malformed entries are dropped
 * rather than crashing the import path.
 *
 * DOM-free and engine-free; unit-tested in tests/unit/persistence.spec.ts.
 */

export type SelectionMode = 'single' | 'multi';
export type NavigationMode = 'Orbit' | 'Plan' | 'FirstPerson';
/** UI language (C7). Mirrors i18n.Language; kept local so persistence stays dep-free. */
export type PersistedLanguage = 'en' | 'de';
export type VisualStyle = 'basic' | 'pen' | 'color-pen' | 'color-shadows' | 'color-pen-shadows';
/** Mirrors OBC.CameraProjection ("Perspective" | "Orthographic"). */
export type CameraProjection = 'Perspective' | 'Orthographic';

export interface Vector3Record {
  x: number;
  y: number;
  z: number;
}

export interface SavedViewpoint {
  id: string;
  name: string;
  createdAt: string;
  camera: {
    position: Vector3Record;
    target: Vector3Record;
    projection: CameraProjection;
    mode: NavigationMode;
  };
  clippingPlanes: Array<{ normal: Vector3Record; origin: Vector3Record }>;
  hiddenItems: Record<string, number[]>;
  visualStyle?: VisualStyle;
  xray: boolean;
  edges: boolean;
  snapshot?: string;
}

export interface IssueCommentRecord {
  id: string;
  text: string;
  author: string;
  createdAt: string;
}

export const ISSUE_PRIORITIES = ['Critical', 'High', 'Medium', 'Low'] as const;
export const ISSUE_STATUSES = ['Open', 'In Progress', 'Resolved', 'Closed'] as const;

export interface PersistedIssue {
  id: string;
  title: string;
  description: string;
  priority: (typeof ISSUE_PRIORITIES)[number];
  status: (typeof ISSUE_STATUSES)[number];
  assignee: string;
  createdAt: string;
  updatedAt: string;
  modelId: string | null;
  localIds: number[];
  /** F9: full multi-model selection; modelId/localIds keep the first model for back-compat. */
  elementsByModel?: Record<string, number[]>;
  point: Vector3Record | null;
  comments: IssueCommentRecord[];
}

/** Current persisted-state schema version (C8 full-session persistence, W5.2). */
export const PERSISTED_STATE_VERSION = 2;

/**
 * C8: a loaded model, restorable without re-conversion. The `.frag` bytes live
 * in IndexedDB (see core/frag-cache.ts) keyed by `fragKey`; this record carries
 * the identity + every per-model modification so restore re-adds the model and
 * re-applies the changes.
 */
export interface PersistedModelRecord {
  /** The in-session model id (the id passed to the fragments loader). */
  modelId: string;
  fileName: string;
  sizeBytes: number;
  /** Content-hash key of the cached `.frag` bytes in IndexedDB. */
  fragKey: string;
  /** Transform offset applied via the gizmo / numeric inputs (metres / degrees). */
  offsetPosition: Vector3Record;
  offsetRotation: Vector3Record;
  /** Per-model opacity (0–1) and visibility. */
  opacity: number;
  visible: boolean;
  /** Isolated/hidden element ids for this model (isolation + hide/show state). */
  hiddenIds: number[];
}

/** C8: camera pose to restore the exact view on reopen. */
export interface PersistedCamera {
  position: Vector3Record;
  target: Vector3Record;
  projection: CameraProjection;
}

export interface PersistedViewerState {
  version: 2;
  selectionMode: SelectionMode;
  navigationMode: NavigationMode;
  visualStyle?: VisualStyle;
  xray: boolean;
  edges: boolean;
  gridVisible?: boolean;
  backgroundColor?: string;
  /** F8: each theme remembers its own background. */
  backgroundByTheme?: { dark: string; light: string };
  theme?: 'dark' | 'light';
  /** C7: persisted UI language so a restored session reopens in the same locale. */
  language?: PersistedLanguage;
  viewpoints: SavedViewpoint[];
  issues: PersistedIssue[];
  /** C8 (W5.2): the full loaded-model set + per-model modifications. */
  models: PersistedModelRecord[];
  /** C8: restored camera pose. */
  camera?: PersistedCamera;
  /** C8: active section planes (normal + coplanar origin). */
  sectionPlanes: Array<{ normal: Vector3Record; origin: Vector3Record }>;
  /** C8: selected elements per model, so the selection restores. */
  selection: Record<string, number[]>;
  /** C8: the active right-panel tab, restored for continuity. */
  activeTab?: string;
  /** Search sets (2026-07-12): saved searches with color override + visibility. */
  searchSets?: PersistedSearchSet[];
}

/** A Navisworks-style saved search: captured elements + color + visibility. */
export interface PersistedSearchSet {
  id: string;
  name: string;
  query: string;
  color: string;
  colorActive: boolean;
  visible: boolean;
  elementsByModel: Record<string, number[]>;
}

/**
 * Input to the pure serializer `buildPersistedState` (the W3.5-deferred
 * persistence-serializer extraction, folded into W5.2). The orchestrator gathers
 * the raw current state; this module projects it into a validated
 * PersistedViewerState. Kept DOM-free and engine-free for unit testing.
 */
export interface PersistedStateInput {
  selectionMode: SelectionMode;
  navigationMode: NavigationMode;
  visualStyle: VisualStyle;
  xray: boolean;
  edges: boolean;
  gridVisible: boolean;
  backgroundColor: string;
  backgroundByTheme: { dark: string; light: string };
  theme: 'dark' | 'light';
  language: PersistedLanguage;
  viewpoints: SavedViewpoint[];
  issues: PersistedIssue[];
  models: PersistedModelRecord[];
  camera?: PersistedCamera;
  sectionPlanes: Array<{ normal: Vector3Record; origin: Vector3Record }>;
  selection: Record<string, number[]>;
  activeTab?: string;
  searchSets: PersistedSearchSet[];
  /** Cap for persisted viewpoint snapshots (F2 size guard); larger are dropped. */
  maxSnapshotChars: number;
}

/**
 * Projects the current viewer state into a validated PersistedViewerState. Deep-
 * copies mutable structures so the caller's live objects are not aliased into
 * the persisted payload, and applies the F2 snapshot size guard.
 */
export const buildPersistedState = (input: PersistedStateInput): PersistedViewerState => {
  const issues = input.issues.map((issue) => ({
    ...issue,
    localIds: [...issue.localIds],
    elementsByModel: issue.elementsByModel
      ? Object.fromEntries(Object.entries(issue.elementsByModel).map(([modelId, ids]) => [modelId, [...ids]]))
      : undefined,
    point: issue.point ? { ...issue.point } : null,
    comments: issue.comments.map((comment) => ({ ...comment })),
  }));

  const viewpoints = input.viewpoints.map((viewpoint) => ({
    ...viewpoint,
    snapshot:
      viewpoint.snapshot && viewpoint.snapshot.length <= input.maxSnapshotChars
        ? viewpoint.snapshot
        : undefined,
  }));

  const models = input.models.map((model) => ({
    ...model,
    offsetPosition: { ...model.offsetPosition },
    offsetRotation: { ...model.offsetRotation },
    hiddenIds: [...model.hiddenIds],
  }));

  const selection: Record<string, number[]> = {};
  for (const [modelId, ids] of Object.entries(input.selection)) {
    if (ids.length > 0) selection[modelId] = [...ids];
  }

  return {
    version: 2,
    selectionMode: input.selectionMode,
    navigationMode: input.navigationMode,
    visualStyle: input.visualStyle,
    xray: input.xray,
    edges: input.edges,
    gridVisible: input.gridVisible,
    backgroundColor: input.backgroundColor,
    backgroundByTheme: { ...input.backgroundByTheme },
    theme: input.theme,
    language: input.language,
    viewpoints,
    issues,
    models,
    camera: input.camera
      ? {
          position: { ...input.camera.position },
          target: { ...input.camera.target },
          projection: input.camera.projection,
        }
      : undefined,
    sectionPlanes: input.sectionPlanes.map((plane) => ({
      normal: { ...plane.normal },
      origin: { ...plane.origin },
    })),
    selection,
    activeTab: input.activeTab,
    searchSets: input.searchSets.map((set) => ({
      ...set,
      elementsByModel: Object.fromEntries(
        Object.entries(set.elementsByModel).map(([modelId, ids]) => [modelId, [...ids]]),
      ),
    })),
  };
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

const asString = (value: unknown, fallback: string): string =>
  typeof value === 'string' ? value : fallback;

const asBoolean = (value: unknown, fallback: boolean): boolean =>
  typeof value === 'boolean' ? value : fallback;

const asFiniteNumber = (value: unknown): number | null =>
  typeof value === 'number' && Number.isFinite(value) ? value : null;

const asVector3 = (value: unknown): Vector3Record | null => {
  if (!isRecord(value)) return null;
  const x = asFiniteNumber(value.x);
  const y = asFiniteNumber(value.y);
  const z = asFiniteNumber(value.z);
  if (x === null || y === null || z === null) return null;
  return { x, y, z };
};

const asEnum = <T extends string>(value: unknown, allowed: readonly T[], fallback: T): T =>
  typeof value === 'string' && (allowed as readonly string[]).includes(value) ? (value as T) : fallback;

const asLocalIds = (value: unknown): number[] =>
  Array.isArray(value) ? value.filter((id): id is number => typeof id === 'number' && Number.isFinite(id)) : [];

const asHiddenItems = (value: unknown): Record<string, number[]> => {
  if (!isRecord(value)) return {};
  const result: Record<string, number[]> = {};
  for (const [modelId, ids] of Object.entries(value)) {
    const clean = asLocalIds(ids);
    if (clean.length > 0) result[modelId] = clean;
  }
  return result;
};

const asSnapshot = (value: unknown): string | undefined =>
  typeof value === 'string' && value.startsWith('data:image/') ? value : undefined;

const normalizeViewpoint = (value: unknown): SavedViewpoint | null => {
  if (!isRecord(value)) return null;
  const id = asString(value.id, '');
  const name = asString(value.name, '');
  if (!id || !name) return null;

  const camera = isRecord(value.camera) ? value.camera : null;
  if (!camera) return null;
  const position = asVector3(camera.position);
  const target = asVector3(camera.target);
  if (!position || !target) return null;

  const clippingPlanes: SavedViewpoint['clippingPlanes'] = [];
  if (Array.isArray(value.clippingPlanes)) {
    for (const plane of value.clippingPlanes) {
      if (!isRecord(plane)) continue;
      const normal = asVector3(plane.normal);
      const origin = asVector3(plane.origin);
      if (normal && origin) clippingPlanes.push({ normal, origin });
    }
  }

  return {
    id,
    name,
    createdAt: asString(value.createdAt, new Date(0).toISOString()),
    camera: {
      position,
      target,
      projection: asEnum<CameraProjection>(camera.projection, ['Perspective', 'Orthographic'], 'Perspective'),
      mode: asEnum<NavigationMode>(camera.mode, ['Orbit', 'Plan', 'FirstPerson'], 'Orbit'),
    },
    clippingPlanes,
    hiddenItems: asHiddenItems(value.hiddenItems),
    visualStyle: typeof value.visualStyle === 'string'
      ? asEnum<VisualStyle>(value.visualStyle, ['basic', 'pen', 'color-pen', 'color-shadows', 'color-pen-shadows'], 'color-pen-shadows')
      : undefined,
    xray: asBoolean(value.xray, false),
    edges: asBoolean(value.edges, false),
    snapshot: asSnapshot(value.snapshot),
  };
};

const normalizeComment = (value: unknown): IssueCommentRecord | null => {
  if (!isRecord(value)) return null;
  const id = asString(value.id, '');
  const text = asString(value.text, '');
  if (!id || !text) return null;
  return {
    id,
    text,
    author: asString(value.author, 'Unknown'),
    createdAt: asString(value.createdAt, new Date(0).toISOString()),
  };
};

const asElementsByModel = (value: unknown): Record<string, number[]> | undefined => {
  if (!isRecord(value)) return undefined;
  const result: Record<string, number[]> = {};
  for (const [modelId, ids] of Object.entries(value)) {
    const clean = asLocalIds(ids);
    if (clean.length > 0) result[modelId] = clean;
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

const normalizeIssue = (value: unknown): PersistedIssue | null => {
  if (!isRecord(value)) return null;
  const id = asString(value.id, '');
  const title = asString(value.title, '');
  if (!id || !title) return null;
  return {
    id,
    title,
    description: asString(value.description, ''),
    priority: asEnum(value.priority, ISSUE_PRIORITIES, 'Medium'),
    status: asEnum(value.status, ISSUE_STATUSES, 'Open'),
    assignee: asString(value.assignee, ''),
    createdAt: asString(value.createdAt, new Date(0).toISOString()),
    updatedAt: asString(value.updatedAt, new Date(0).toISOString()),
    modelId: typeof value.modelId === 'string' ? value.modelId : null,
    localIds: asLocalIds(value.localIds),
    elementsByModel: asElementsByModel(value.elementsByModel),
    point: asVector3(value.point),
    comments: Array.isArray(value.comments)
      ? value.comments.map(normalizeComment).filter((comment): comment is IssueCommentRecord => comment !== null)
      : [],
  };
};

const HEX_COLOR = /^#[0-9a-f]{6}$/i;

const asHexColor = (value: unknown): string | undefined =>
  typeof value === 'string' && HEX_COLOR.test(value.trim()) ? value.trim().toLowerCase() : undefined;

const asVector3Or = (value: unknown, fallback: Vector3Record): Vector3Record =>
  asVector3(value) ?? { ...fallback };

const ZERO_VEC: Vector3Record = { x: 0, y: 0, z: 0 };

/** C8: validates one persisted model record; drops records missing an identity. */
const normalizeModelRecord = (value: unknown): PersistedModelRecord | null => {
  if (!isRecord(value)) return null;
  const modelId = asString(value.modelId, '');
  const fragKey = asString(value.fragKey, '');
  if (!modelId || !fragKey) return null;
  const opacity = asFiniteNumber(value.opacity);
  return {
    modelId,
    fileName: asString(value.fileName, modelId),
    sizeBytes: asFiniteNumber(value.sizeBytes) ?? 0,
    fragKey,
    offsetPosition: asVector3Or(value.offsetPosition, ZERO_VEC),
    offsetRotation: asVector3Or(value.offsetRotation, ZERO_VEC),
    opacity: opacity === null ? 1 : Math.min(1, Math.max(0, opacity)),
    visible: asBoolean(value.visible, true),
    hiddenIds: asLocalIds(value.hiddenIds),
  };
};

const asModelRecords = (value: unknown): PersistedModelRecord[] =>
  Array.isArray(value)
    ? value.map(normalizeModelRecord).filter((record): record is PersistedModelRecord => record !== null)
    : [];

const asCamera = (value: unknown): PersistedCamera | undefined => {
  if (!isRecord(value)) return undefined;
  const position = asVector3(value.position);
  const target = asVector3(value.target);
  if (!position || !target) return undefined;
  return {
    position,
    target,
    projection: asEnum<CameraProjection>(value.projection, ['Perspective', 'Orthographic'], 'Perspective'),
  };
};

const asSectionPlanes = (value: unknown): Array<{ normal: Vector3Record; origin: Vector3Record }> => {
  if (!Array.isArray(value)) return [];
  const planes: Array<{ normal: Vector3Record; origin: Vector3Record }> = [];
  for (const plane of value) {
    if (!isRecord(plane)) continue;
    const normal = asVector3(plane.normal);
    const origin = asVector3(plane.origin);
    if (normal && origin) planes.push({ normal, origin });
  }
  return planes;
};

const asSelection = (value: unknown): Record<string, number[]> => {
  if (!isRecord(value)) return {};
  const result: Record<string, number[]> = {};
  for (const [modelId, ids] of Object.entries(value)) {
    const clean = asLocalIds(ids);
    if (clean.length > 0) result[modelId] = clean;
  }
  return result;
};

/**
 * Validates and defaults an untrusted persisted-state value (A7 + C8 W5.2).
 * Returns null when the value is not a recognizable persisted state; otherwise
 * every field is present, typed and safe to apply. Accepts both the current v2
 * schema and the legacy v1 schema (migrated forward: v1 had no models/camera/
 * section/selection — those default to empty). `{"version":2}` and
 * `{"version":1}` are both valid minimal imports.
 */
export const normalizePersistedState = (raw: unknown): PersistedViewerState | null => {
  if (!isRecord(raw)) return null;
  // v1 → v2 migration is transparent: the shared fields are read the same way,
  // and the C8 fields simply default to empty for a v1 blob.
  if (raw.version !== 1 && raw.version !== 2) return null;

  const backgroundByThemeRaw = isRecord(raw.backgroundByTheme) ? raw.backgroundByTheme : null;
  const dark = backgroundByThemeRaw ? asHexColor(backgroundByThemeRaw.dark) : undefined;
  const light = backgroundByThemeRaw ? asHexColor(backgroundByThemeRaw.light) : undefined;

  return {
    version: 2,
    selectionMode: asEnum<SelectionMode>(raw.selectionMode, ['single', 'multi'], 'single'),
    navigationMode: asEnum<NavigationMode>(raw.navigationMode, ['Orbit', 'Plan', 'FirstPerson'], 'Orbit'),
    visualStyle: asEnum<VisualStyle>(
      raw.visualStyle,
      ['basic', 'pen', 'color-pen', 'color-shadows', 'color-pen-shadows'],
      'color-pen-shadows',
    ),
    xray: asBoolean(raw.xray, false),
    edges: asBoolean(raw.edges, false),
    gridVisible: asBoolean(raw.gridVisible, false),
    backgroundColor: asHexColor(raw.backgroundColor),
    backgroundByTheme: dark && light ? { dark, light } : undefined,
    theme: asEnum<'dark' | 'light'>(raw.theme, ['dark', 'light'], 'dark'),
    // C7: language is optional — undefined means "leave as-is" (the i18n module
    // has its own persisted key), so we only carry a value when it's a valid enum.
    language: typeof raw.language === 'string'
      ? asEnum<PersistedLanguage>(raw.language, ['en', 'de'], 'en')
      : undefined,
    viewpoints: Array.isArray(raw.viewpoints)
      ? raw.viewpoints.map(normalizeViewpoint).filter((viewpoint): viewpoint is SavedViewpoint => viewpoint !== null)
      : [],
    issues: Array.isArray(raw.issues)
      ? raw.issues.map(normalizeIssue).filter((issue): issue is PersistedIssue => issue !== null)
      : [],
    // C8 (W5.2): full-session fields — absent (v1) → empty defaults.
    models: asModelRecords(raw.models),
    camera: asCamera(raw.camera),
    sectionPlanes: asSectionPlanes(raw.sectionPlanes),
    selection: asSelection(raw.selection),
    activeTab: typeof raw.activeTab === 'string' ? raw.activeTab : undefined,
    searchSets: asSearchSets(raw.searchSets),
  };
};

/** Search sets: absent/malformed entries are dropped, never crash (A7). */
const asSearchSets = (value: unknown): PersistedSearchSet[] => {
  if (!Array.isArray(value)) return [];
  const sets: PersistedSearchSet[] = [];
  for (const entry of value) {
    if (!isRecord(entry)) continue;
    if (typeof entry.id !== 'string' || typeof entry.name !== 'string') continue;
    const elementsByModel: Record<string, number[]> = {};
    if (isRecord(entry.elementsByModel)) {
      for (const [modelId, ids] of Object.entries(entry.elementsByModel)) {
        if (!Array.isArray(ids)) continue;
        const clean = ids.filter((id): id is number => typeof id === 'number' && Number.isFinite(id));
        if (clean.length > 0) elementsByModel[modelId] = clean;
      }
    }
    sets.push({
      id: entry.id,
      name: entry.name,
      query: asString(entry.query, ''),
      color: asString(entry.color, '#e4572e'),
      colorActive: asBoolean(entry.colorActive, false),
      visible: asBoolean(entry.visible, true),
      elementsByModel,
    });
  }
  return sets;
};
