/**
 * Pure view-cube geometry helpers (W2, extracted from viewer.ts's view-cube
 * cluster). These are stateless THREE math — the camera basis axes, the
 * navigation distance for a direction, and the up-vector resolution — with no
 * engine/DOM/instance dependencies, so they are unit-testable. The stateful
 * parts of the view cube (camera moves, hotspot DOM projection) stay in
 * viewer.ts. `three` is imported for its vector/quaternion math only.
 */

import * as THREE from 'three';

export interface ViewCubeAxes {
  right: THREE.Vector3;
  up: THREE.Vector3;
  front: THREE.Vector3;
}

/** The world-space right/up/front axes for a camera basis quaternion. */
export const getViewCubeAxes = (basis: THREE.Quaternion): ViewCubeAxes => ({
  right: new THREE.Vector3(1, 0, 0).applyQuaternion(basis).normalize(),
  up: new THREE.Vector3(0, 1, 0).applyQuaternion(basis).normalize(),
  front: new THREE.Vector3(0, 0, 1).applyQuaternion(basis).normalize(),
});

/**
 * Camera distance for a view direction: corners pull back furthest, then edges,
 * then faces — so the whole model stays framed from any cube target.
 */
export const getViewCubeNavigationDistance = (maxDim: number, localDirection: THREE.Vector3): number => {
  const components =
    Number(Math.abs(localDirection.x) > 0.01)
    + Number(Math.abs(localDirection.y) > 0.01)
    + Number(Math.abs(localDirection.z) > 0.01);
  if (components >= 3) return maxDim * 2.45;
  if (components === 2) return maxDim * 2.25;
  return maxDim * 2;
};

/**
 * Resolves a stable camera up-vector for a view direction, orthogonalized
 * against the world look direction. Axis-aligned views get a deterministic up
 * (e.g. top view looks down its front axis); oblique views fall back through
 * up/front/right candidates. Returns +Y if none is usable.
 */
export const resolveViewCubeCameraUp = (
  localDirection: THREE.Vector3,
  worldDirection: THREE.Vector3,
  axes: ViewCubeAxes,
): THREE.Vector3 => {
  const candidates: THREE.Vector3[] = [];
  const axialX = Math.abs(localDirection.x) > 0.999 && Math.abs(localDirection.y) < 0.001 && Math.abs(localDirection.z) < 0.001;
  const axialY = Math.abs(localDirection.y) > 0.999 && Math.abs(localDirection.x) < 0.001 && Math.abs(localDirection.z) < 0.001;
  const axialZ = Math.abs(localDirection.z) > 0.999 && Math.abs(localDirection.x) < 0.001 && Math.abs(localDirection.y) < 0.001;

  if (axialY) {
    candidates.push(localDirection.y >= 0 ? axes.front.clone() : axes.front.clone().negate(), axes.right.clone());
  } else if (axialX) {
    candidates.push(axes.up.clone(), localDirection.x >= 0 ? axes.front.clone() : axes.front.clone().negate());
  } else if (axialZ) {
    candidates.push(axes.up.clone(), localDirection.z >= 0 ? axes.right.clone() : axes.right.clone().negate());
  } else {
    candidates.push(axes.up.clone(), axes.front.clone(), axes.right.clone());
  }

  for (const candidate of candidates) {
    candidate.addScaledVector(worldDirection, -candidate.dot(worldDirection));
    if (candidate.lengthSq() > 0.0001) return candidate.normalize();
  }
  return new THREE.Vector3(0, 1, 0);
};
