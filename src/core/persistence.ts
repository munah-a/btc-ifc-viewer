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

export interface PersistedViewerState {
  version: 1;
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
  viewpoints: SavedViewpoint[];
  issues: PersistedIssue[];
}

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

/**
 * Validates and defaults an untrusted persisted-state value (A7). Returns
 * null when the value is not a recognizable v1 state; otherwise every field
 * is present, typed and safe to apply. `{"version":1}` is a valid minimal
 * import.
 */
export const normalizePersistedState = (raw: unknown): PersistedViewerState | null => {
  if (!isRecord(raw) || raw.version !== 1) return null;

  const backgroundByThemeRaw = isRecord(raw.backgroundByTheme) ? raw.backgroundByTheme : null;
  const dark = backgroundByThemeRaw ? asHexColor(backgroundByThemeRaw.dark) : undefined;
  const light = backgroundByThemeRaw ? asHexColor(backgroundByThemeRaw.light) : undefined;

  return {
    version: 1,
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
    viewpoints: Array.isArray(raw.viewpoints)
      ? raw.viewpoints.map(normalizeViewpoint).filter((viewpoint): viewpoint is SavedViewpoint => viewpoint !== null)
      : [],
    issues: Array.isArray(raw.issues)
      ? raw.issues.map(normalizeIssue).filter((issue): issue is PersistedIssue => issue !== null)
      : [],
  };
};
