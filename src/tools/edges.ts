/**
 * Edges tool — pure edge-overlay geometry builder with a per-geometry cache (P3).
 *
 * The scene add/remove + overlay-tracking lifecycle stays in the orchestrator
 * (it holds the world/scene and the overlay list disposed on destroy). This
 * builds the LineSegments overlays for the visible meshes under the given model
 * objects, using the shared edge material.
 *
 * P3 (AUDIT / W5.3): the original inline applyEdges rebuilt an `EdgesGeometry`
 * (O(triangles)) for EVERY mesh on EVERY call — including per opacity-slider
 * tick and per pointer-move during a gizmo drag. `EdgeGeometryCache` keys the
 * computed `EdgesGeometry` by the SOURCE geometry's uuid so it is built once and
 * only the overlay's world matrix is refreshed on subsequent calls. The cached
 * EdgesGeometry is shared across overlays and is NOT disposed with them — the
 * cache owns it (dispose the whole cache on teardown).
 */
import * as THREE from 'three';

/** Feature angle (degrees) above which a mesh edge is drawn. */
export const EDGE_THRESHOLD_DEGREES = 35;

/**
 * Caches `EdgesGeometry` per source-geometry uuid (P3). Reused across
 * `buildEdgeOverlays` calls so edges are not recomputed on every opacity/gizmo
 * tick. `dispose()` frees every cached geometry (call on viewer teardown).
 */
export class EdgeGeometryCache {
  private readonly cache = new Map<string, THREE.BufferGeometry>();

  /** Returns the cached EdgesGeometry for `source`, building it on first use. */
  get(source: THREE.BufferGeometry): THREE.BufferGeometry {
    let edges = this.cache.get(source.uuid);
    if (!edges) {
      edges = new THREE.EdgesGeometry(source, EDGE_THRESHOLD_DEGREES);
      this.cache.set(source.uuid, edges);
    }
    return edges;
  }

  dispose(): void {
    for (const geometry of this.cache.values()) geometry.dispose();
    this.cache.clear();
  }
}

/**
 * Builds edge-line overlays for every visible mesh under `modelObjects`.
 * Overlays use `matrixAutoUpdate = false` with the mesh world matrix copied in,
 * so they track the mesh without per-frame recomputation. Invalid geometry is
 * skipped. When a `cache` is given the underlying `EdgesGeometry` is reused per
 * source geometry (P3) — only the overlay's world matrix is set each call.
 *
 * The caller adds the returned overlays to the scene and tracks them for removal.
 * IMPORTANT: when a cache is used, the overlays share the cache's geometry, so
 * the caller must NOT dispose overlay.geometry (dispose the cache instead).
 */
export function buildEdgeOverlays(
  modelObjects: THREE.Object3D[],
  material: THREE.Material,
  cache?: EdgeGeometryCache,
): THREE.LineSegments[] {
  const overlays: THREE.LineSegments[] = [];
  for (const object of modelObjects) {
    if (!object.visible) continue;
    object.traverse((child: THREE.Object3D) => {
      const mesh = child as THREE.Mesh;
      if (!mesh.isMesh || !mesh.geometry || !mesh.visible) return;
      try {
        const geometry = cache
          ? cache.get(mesh.geometry)
          : new THREE.EdgesGeometry(mesh.geometry, EDGE_THRESHOLD_DEGREES);
        const lines = new THREE.LineSegments(geometry, material);
        lines.matrixAutoUpdate = false;
        lines.matrix.copy(mesh.matrixWorld);
        overlays.push(lines);
      } catch {
        // ignore invalid geometry
      }
    });
  }
  return overlays;
}
