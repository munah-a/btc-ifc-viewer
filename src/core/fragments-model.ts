/**
 * Typed boundary for the ThatOpen fragments model (AUDIT A8, W2.2).
 *
 * viewer.ts previously treated every fragments model as `any` (~10 methods, no
 * autocomplete, no compile-time safety on a `^`-ranged dependency). This module
 * declares `FragmentsModelLike` — a *structural* interface covering exactly the
 * methods and properties the app calls — reusing the library's own stable data
 * types (Identifier/ItemData/… are plain shapes, not the concrete class) so the
 * app is decoupled from the concrete `FragmentsModel` implementation while
 * still being fully type-checked.
 *
 * It also isolates the one private-field reach into the clipper's plane
 * controls (`_controls`) behind a single, clearly-commented helper, so that
 * ThatOpen-internal access lives in exactly one place (A8).
 */

import type * as THREE from 'three';
import type {
  Identifier,
  ItemData,
  ItemsDataConfig,
  ItemsQueryConfig,
  ItemsQueryParams,
  MeshData,
  SpatialTreeItem,
} from '@thatopen/fragments';

/**
 * The subset of `@thatopen/fragments`' `FragmentsModel` that this app uses.
 * Structural, so any object with these members satisfies it — insulating the
 * app from the concrete class shape across `^`-ranged upgrades.
 */
export interface FragmentsModelLike {
  /** Model id — the name passed to `ifcLoader.load` (A6). */
  readonly modelId: string;
  /** The three.js object representing the model in the scene. */
  object: THREE.Object3D;
  /**
   * Data-driven model bounding box (local space) — available right after load,
   * independent of whether geometry has streamed/rendered yet. Prefer this over
   * `expandByObject(object)` for fit/section, which reads empty until the
   * fragments worker has streamed meshes (AUDIT A17).
   */
  readonly box: THREE.Box3;
  /** Graphics quality, 0 (lowest) … 1 (highest). */
  graphicsQuality: number;

  useCamera(camera: THREE.PerspectiveCamera | THREE.OrthographicCamera): void;

  getItemsIdsWithGeometry(): Promise<number[]>;
  getItemsWithGeometryCategories(): Promise<(string | null)[]>;
  getSpatialStructure(): Promise<SpatialTreeItem>;
  getItemsOfCategories(categories: RegExp[]): Promise<Record<string, number[]>>;
  getItemsData(ids: Identifier[], config?: Partial<ItemsDataConfig>): Promise<ItemData[]>;
  getItemsByQuery(params: ItemsQueryParams, config?: ItemsQueryConfig): Promise<number[]>;
  getItemsGeometry(localIds: number[], lod?: unknown): Promise<MeshData[][]>;
  getItemsVolume(localIds: number[]): Promise<number>;

  setOpacity(localIds: number[] | undefined, opacity: number): Promise<void>;
  resetOpacity(localIds: number[] | undefined): Promise<void>;
  resetColor(localIds: number[] | undefined): Promise<void>;
}

/**
 * Reaches the transform-gizmo helper of a clipper plane so the app can exempt
 * it from clipping/depth-test (the arrow must always render on top).
 *
 * WARNING (A8): `_controls` is a **private field** of ThatOpen's clipper plane
 * (`SimplePlane`). There is no public accessor for its `TransformControls`
 * helper, so this cast is the one sanctioned reach into library internals.
 * Keep every such access here; revisit at the W6 stack upgrade in case a public
 * API appears. Returns null if the field is absent (defensive across versions).
 */
export const getClipperPlaneGizmoHelper = (plane: unknown): THREE.Object3D | null => {
  const controls = (plane as { _controls?: { getHelper?: () => THREE.Object3D } })?._controls;
  if (!controls || typeof controls.getHelper !== 'function') return null;
  return controls.getHelper();
};
