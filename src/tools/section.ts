/**
 * Section tool — pure geometry helpers.
 *
 * The Clipper wiring stays in the orchestrator (it holds engine handles and the
 * active-section slider state); this module holds the DOM-free math so it is
 * unit-testable and shared by the axis-plane slider and the section box.
 */
import * as THREE from 'three';

/**
 * The coplanar point for a single axis section plane at `pct` (0–100) along the
 * model bounding box. Slides along the dominant axis of the plane normal
 * between the box extremes; other coordinates stay at the box center.
 */
export function sectionPlanePoint(box: THREE.Box3, normal: THREE.Vector3, pct: number): THREE.Vector3 {
  const t = Math.min(1, Math.max(0, pct / 100));
  const { min, max } = box;
  const point = box.getCenter(new THREE.Vector3());
  if (Math.abs(normal.x) > 0.5) point.x = min.x + (max.x - min.x) * t;
  else if (Math.abs(normal.y) > 0.5) point.y = min.y + (max.y - min.y) * t;
  else point.z = min.z + (max.z - min.z) * t;
  return point;
}

/**
 * The six {normal, point} plane definitions that clip to the model bounding
 * box (a "section box"). Point coordinates match the original inline creation
 * so behaviour is identical.
 */
export function sectionBoxPlanes(box: THREE.Box3): Array<{ normal: THREE.Vector3; point: THREE.Vector3 }> {
  const { min, max } = box;
  return [
    { normal: new THREE.Vector3(-1, 0, 0), point: new THREE.Vector3(max.x, 0, 0) },
    { normal: new THREE.Vector3(1, 0, 0), point: new THREE.Vector3(min.x, 0, 0) },
    { normal: new THREE.Vector3(0, -1, 0), point: new THREE.Vector3(0, max.y, 0) },
    { normal: new THREE.Vector3(0, 1, 0), point: new THREE.Vector3(0, min.y, 0) },
    { normal: new THREE.Vector3(0, 0, -1), point: new THREE.Vector3(0, 0, max.z) },
    { normal: new THREE.Vector3(0, 0, 1), point: new THREE.Vector3(0, 0, min.z) },
  ];
}
