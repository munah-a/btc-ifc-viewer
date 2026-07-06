/**
 * Edges tool — pure edge-overlay geometry builder.
 *
 * The scene add/remove + overlay-tracking lifecycle stays in the orchestrator
 * (it holds the world/scene and the overlay list disposed on destroy). This
 * builds the LineSegments overlays for the visible meshes under the given model
 * objects, using the shared edge material, with the same threshold and
 * matrix-copy behaviour as the original inline applyEdges.
 */
import * as THREE from 'three';

/** Feature angle (degrees) above which a mesh edge is drawn. */
export const EDGE_THRESHOLD_DEGREES = 35;

/**
 * Builds edge-line overlays for every visible mesh under `modelObjects`.
 * Overlays use `matrixAutoUpdate = false` with the mesh world matrix copied in,
 * so they track the mesh without per-frame recomputation. Invalid geometry is
 * skipped. The caller adds the returned overlays to the scene and tracks them
 * for disposal.
 */
export function buildEdgeOverlays(
  modelObjects: THREE.Object3D[],
  material: THREE.Material,
): THREE.LineSegments[] {
  const overlays: THREE.LineSegments[] = [];
  for (const object of modelObjects) {
    if (!object.visible) continue;
    object.traverse((child: THREE.Object3D) => {
      const mesh = child as THREE.Mesh;
      if (!mesh.isMesh || !mesh.geometry || !mesh.visible) return;
      try {
        const geometry = new THREE.EdgesGeometry(mesh.geometry, EDGE_THRESHOLD_DEGREES);
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
