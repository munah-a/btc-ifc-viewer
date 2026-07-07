/**
 * Shared runtime types used by the viewer orchestrator (viewer.ts) and the
 * extracted ui/* panel controllers. These describe the in-memory model index,
 * spatial tree nodes, and per-model federation records — the data the panel
 * renderers consume. Kept DOM-free so ui/* modules and unit tests can import
 * them without pulling in the engine.
 */
import type * as THREE from 'three';
import type { PersistedIssue } from './persistence';

export type MeasureMode = 'none' | 'length' | 'area';

/**
 * Persisted shapes (SavedViewpoint, PersistedIssue, PersistedViewerState) live
 * in core/persistence.ts (A7); the runtime issue record adds the marker handle.
 */
export interface IssueRecord extends PersistedIssue {
  markerId?: string;
}

export interface SearchResult {
  modelId: string;
  localId: number;
  name: string;
  type: string;
  globalId: string;
}

export interface ModelIndex {
  modelId: string;
  allIds: Set<number>;
  classes: Map<string, Set<number>>;
  levels: Map<string, Set<number>>;
  itemToLevel: Map<number, string>;
  itemNames: Map<number, string>;
  spatialRoot: BrowserTreeNode | null;
}

export interface BrowserTreeNode {
  category: string;
  localId: number | null;
  label: string;
  geometryCount: number;
  children: BrowserTreeNode[];
}

export interface TransformVector3 {
  x: number;
  y: number;
  z: number;
}

export interface FederatedModelRecord {
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
  /**
   * C8 (W5.2): content-hash key of this model's cached `.frag` bytes in
   * IndexedDB. Set after the fragments are available; empty until then (a model
   * with no fragKey is skipped by session persistence — it can't be restored).
   */
  fragKey?: string;
}
