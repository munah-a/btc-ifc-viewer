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
 * Inverse of {@link sectionPlanePoint} (C5, W5-fixups): the slider percentage
 * (0–100) for a plane whose coplanar `origin` sits along the box's dominant axis.
 * Used to re-populate the section slider after a session restore recreates a
 * single-axis plane geometrically. Clamped to 0–100; degenerate (zero-extent)
 * axes return 50 (the neutral mid position).
 */
export function sectionPlanePercent(box: THREE.Box3, normal: THREE.Vector3, origin: THREE.Vector3): number {
  const { min, max } = box;
  let value: number;
  let lo: number;
  let hi: number;
  if (Math.abs(normal.x) > 0.5) {
    value = origin.x;
    lo = min.x;
    hi = max.x;
  } else if (Math.abs(normal.y) > 0.5) {
    value = origin.y;
    lo = min.y;
    hi = max.y;
  } else {
    value = origin.z;
    lo = min.z;
    hi = max.z;
  }
  const span = hi - lo;
  if (span <= 0) return 50;
  return Math.min(100, Math.max(0, ((value - lo) / span) * 100));
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
