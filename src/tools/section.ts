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

// ---------------------------------------------------------------------------
// Autodesk-style interactive sectioning (plane + box gizmo math).
//
// Everything below is pure geometry consumed by tools/section-gizmos.ts (the
// three.js gizmo layer): pointer-ray → drag parameters, plane/box rebuilding,
// and the oriented-box reconstruction used when a persisted session or
// viewpoint restores six clip planes that form a section box.
// ---------------------------------------------------------------------------

/** A pointer ray in world space (direction normalized by the caller). */
export interface SectionRay {
  origin: THREE.Vector3;
  direction: THREE.Vector3;
}

/**
 * An oriented section box: center + half-sizes along three orthonormal axes.
 * Axis-aligned boxes use the world basis; rotating the box gizmo rotates the
 * axes. Face indexing convention (used by every consumer): face `f` covers
 * axis `i = Math.floor(f / 2)` with outward sign `s = f % 2 === 0 ? +1 : -1`
 * — i.e. faces are ordered [+a0, −a0, +a1, −a1, +a2, −a2].
 */
export interface OrientedSectionBox {
  center: THREE.Vector3;
  halfSizes: THREE.Vector3;
  axes: [THREE.Vector3, THREE.Vector3, THREE.Vector3];
}

/** Outward normal of face `f` of an oriented box (see face convention above). */
export function sectionBoxFaceNormal(box: OrientedSectionBox, face: number): THREE.Vector3 {
  const axis = box.axes[Math.floor(face / 2)];
  return axis.clone().multiplyScalar(face % 2 === 0 ? 1 : -1);
}

/** World-space center point of face `f` of an oriented box. */
export function sectionBoxFaceCenter(box: OrientedSectionBox, face: number): THREE.Vector3 {
  const i = Math.floor(face / 2);
  const s = face % 2 === 0 ? 1 : -1;
  const half = i === 0 ? box.halfSizes.x : i === 1 ? box.halfSizes.y : box.halfSizes.z;
  return box.center.clone().addScaledVector(box.axes[i], s * half);
}

/**
 * The six clipping planes of an oriented section box, ordered by face index.
 * Clip normals point INWARD (each plane keeps the box interior), so the clip
 * normal of face `f` is the negated outward face normal and the coplanar point
 * is the face center.
 */
export function orientedSectionBoxPlanes(
  box: OrientedSectionBox,
): Array<{ normal: THREE.Vector3; point: THREE.Vector3 }> {
  const planes: Array<{ normal: THREE.Vector3; point: THREE.Vector3 }> = [];
  for (let face = 0; face < 6; face += 1) {
    planes.push({
      normal: sectionBoxFaceNormal(box, face).negate(),
      point: sectionBoxFaceCenter(box, face),
    });
  }
  return planes;
}

/** An axis-aligned oriented box covering the given Box3 (initial section box). */
export function orientedBoxFromBox3(box: THREE.Box3): OrientedSectionBox {
  const center = box.getCenter(new THREE.Vector3());
  const halfSizes = box.getSize(new THREE.Vector3()).multiplyScalar(0.5);
  return {
    center,
    halfSizes,
    axes: [new THREE.Vector3(1, 0, 0), new THREE.Vector3(0, 1, 0), new THREE.Vector3(0, 0, 1)],
  };
}

const ANTIPARALLEL_TOL = 1e-3;
const ORTHOGONAL_TOL = 1e-3;

/**
 * Reconstructs an oriented section box from six persisted clip planes
 * ({normal, origin} records, any order), or null when the planes do not form
 * a box (wrong count, unpaired normals, non-orthogonal pairs, or an empty /
 * inverted extent). Used by the C8 session restore and viewpoint apply so a
 * restored section box comes back as an interactive box gizmo instead of six
 * anonymous planes.
 *
 * `facePlaneIndices[f]` maps face `f` (see {@link OrientedSectionBox}) to the
 * index of its source plane in the input array, so the caller can bind each
 * gizmo face to the clipper plane it must drive.
 */
export function sectionBoxFromPlanes(
  planes: Array<{ normal: { x: number; y: number; z: number }; origin: { x: number; y: number; z: number } }>,
): { box: OrientedSectionBox; facePlaneIndices: number[] } | null {
  if (planes.length !== 6) return null;
  const normals: THREE.Vector3[] = [];
  const origins: THREE.Vector3[] = [];
  for (const plane of planes) {
    const n = new THREE.Vector3(plane.normal.x, plane.normal.y, plane.normal.z);
    if (n.lengthSq() < 1e-12) return null;
    normals.push(n.normalize());
    origins.push(new THREE.Vector3(plane.origin.x, plane.origin.y, plane.origin.z));
  }

  // Group the six planes into three antiparallel pairs.
  const used = new Array<boolean>(6).fill(false);
  const pairs: Array<{ a: number; b: number }> = [];
  for (let i = 0; i < 6; i += 1) {
    if (used[i]) continue;
    used[i] = true;
    let match = -1;
    for (let j = i + 1; j < 6; j += 1) {
      if (!used[j] && normals[i].dot(normals[j]) < -(1 - ANTIPARALLEL_TOL)) {
        match = j;
        break;
      }
    }
    if (match === -1) return null;
    used[match] = true;
    pairs.push({ a: i, b: match });
  }
  if (pairs.length !== 3) return null;

  // Axis k points along the clip normal of pair-member `a` (which keeps
  // u·x ≥ u·origin_a); member `b` keeps u·x ≤ u·origin_b. lo/hi follow.
  const axes = pairs.map(({ a }) => normals[a].clone()) as [THREE.Vector3, THREE.Vector3, THREE.Vector3];
  const lows = pairs.map(({ a }, k) => axes[k].dot(origins[a]));
  const highs = pairs.map(({ b }, k) => axes[k].dot(origins[b]));
  for (let k = 0; k < 3; k += 1) {
    if (!(highs[k] - lows[k] > 0)) return null;
  }
  if (
    Math.abs(axes[0].dot(axes[1])) > ORTHOGONAL_TOL ||
    Math.abs(axes[0].dot(axes[2])) > ORTHOGONAL_TOL ||
    Math.abs(axes[1].dot(axes[2])) > ORTHOGONAL_TOL
  ) {
    return null;
  }

  // Keep the basis right-handed: flip the third axis (and its extent + pair
  // roles) when the reconstructed basis is left-handed.
  const flipped = new THREE.Vector3().crossVectors(axes[0], axes[1]).dot(axes[2]) < 0;
  if (flipped) {
    axes[2].negate();
    const lo = lows[2];
    lows[2] = -highs[2];
    highs[2] = -lo;
    const pair = pairs[2];
    pairs[2] = { a: pair.b, b: pair.a };
  }

  const center = new THREE.Vector3();
  const halfSizes = new THREE.Vector3();
  const halves = [0, 0, 0];
  for (let k = 0; k < 3; k += 1) {
    center.addScaledVector(axes[k], (lows[k] + highs[k]) / 2);
    halves[k] = (highs[k] - lows[k]) / 2;
  }
  halfSizes.set(halves[0], halves[1], halves[2]);

  // Face (k, +1) clips with normal −axis_k → pair member `b`; face (k, −1)
  // clips with +axis_k → pair member `a`.
  const facePlaneIndices: number[] = [];
  for (let k = 0; k < 3; k += 1) {
    facePlaneIndices.push(pairs[k].b, pairs[k].a);
  }

  return { box: { center, halfSizes, axes }, facePlaneIndices };
}

/**
 * Drag parameter along an axis line: the position `t` (in world units, along
 * `axisDir` from `axisOrigin`) of the point on the axis closest to the pointer
 * ray. Null when the ray is (nearly) parallel to the axis — the drag has no
 * well-defined solution there and callers keep the previous value.
 */
export function axisDragOffset(axisOrigin: THREE.Vector3, axisDir: THREE.Vector3, ray: SectionRay): number | null {
  const w0 = new THREE.Vector3().subVectors(axisOrigin, ray.origin);
  const b = axisDir.dot(ray.direction);
  const denom = 1 - b * b;
  if (Math.abs(denom) < 1e-6) return null;
  const d = axisDir.dot(w0);
  const e = ray.direction.dot(w0);
  return (b * e - d) / denom;
}

/** Ray/plane intersection point, or null when the ray is parallel or points away. */
export function intersectRayPlane(
  ray: SectionRay,
  planePoint: THREE.Vector3,
  planeNormal: THREE.Vector3,
): THREE.Vector3 | null {
  const denom = planeNormal.dot(ray.direction);
  if (Math.abs(denom) < 1e-9) return null;
  const t = planeNormal.dot(new THREE.Vector3().subVectors(planePoint, ray.origin)) / denom;
  if (t < 0) return null;
  return ray.origin.clone().addScaledVector(ray.direction, t);
}

/**
 * Signed rotation angle (radians) from `from` to `to` around `axis`, with both
 * vectors projected onto the plane perpendicular to the axis. Zero when either
 * projection degenerates.
 */
export function signedAngleAroundAxis(axis: THREE.Vector3, from: THREE.Vector3, to: THREE.Vector3): number {
  const f = from.clone().addScaledVector(axis, -axis.dot(from));
  const t = to.clone().addScaledVector(axis, -axis.dot(to));
  if (f.lengthSq() < 1e-12 || t.lengthSq() < 1e-12) return 0;
  return Math.atan2(axis.dot(new THREE.Vector3().crossVectors(f, t)), f.dot(t));
}

/**
 * A stable in-plane orthonormal basis for a plane normal: `u` and `v` span the
 * plane, `(u, v, normal)` is right-handed. Deterministic so the plane quad
 * does not spin as the normal changes slightly.
 */
export function planeBasis(normal: THREE.Vector3): { u: THREE.Vector3; v: THREE.Vector3 } {
  const ref = Math.abs(normal.y) < 0.9 ? new THREE.Vector3(0, 1, 0) : new THREE.Vector3(1, 0, 0);
  const u = new THREE.Vector3().crossVectors(ref, normal).normalize();
  const v = new THREE.Vector3().crossVectors(normal, u);
  return { u, v };
}

/** The eight corner points of a Box3. */
export function box3Corners(box: THREE.Box3): THREE.Vector3[] {
  const { min, max } = box;
  const corners: THREE.Vector3[] = [];
  for (const x of [min.x, max.x])
    for (const y of [min.y, max.y]) for (const z of [min.z, max.z]) corners.push(new THREE.Vector3(x, y, z));
  return corners;
}

/** The dot-product range of a Box3's corners along an axis (plane travel limits). */
export function axisRangeOfBox(box: THREE.Box3, axis: THREE.Vector3): { lo: number; hi: number } {
  let lo = Infinity;
  let hi = -Infinity;
  for (const corner of box3Corners(box)) {
    const d = axis.dot(corner);
    if (d < lo) lo = d;
    if (d > hi) hi = d;
  }
  return { lo, hi };
}

/**
 * The rectangle the section-plane visual should cover: the model bounding box
 * projected onto the plane through `origin` with in-plane basis `(u, v)`,
 * expanded by `marginPct` per side (Autodesk sizes the plane visual to the
 * model with a small margin). Null for an empty box.
 */
export function planeQuadRect(
  bounds: THREE.Box3,
  origin: THREE.Vector3,
  u: THREE.Vector3,
  v: THREE.Vector3,
  marginPct = 0.05,
): { center: THREE.Vector3; halfU: number; halfV: number } | null {
  if (bounds.isEmpty()) return null;
  let minU = Infinity;
  let maxU = -Infinity;
  let minV = Infinity;
  let maxV = -Infinity;
  const rel = new THREE.Vector3();
  for (const corner of box3Corners(bounds)) {
    rel.subVectors(corner, origin);
    const pu = u.dot(rel);
    const pv = v.dot(rel);
    if (pu < minU) minU = pu;
    if (pu > maxU) maxU = pu;
    if (pv < minV) minV = pv;
    if (pv > maxV) maxV = pv;
  }
  const halfU = ((maxU - minU) / 2) * (1 + marginPct);
  const halfV = ((maxV - minV) / 2) * (1 + marginPct);
  const center = origin
    .clone()
    .addScaledVector(u, (minU + maxU) / 2)
    .addScaledVector(v, (minV + maxV) / 2);
  return { center, halfU: Math.max(halfU, 1e-6), halfV: Math.max(halfV, 1e-6) };
}

/**
 * Moves one face of a box along its axis. `centerA`/`half` are the box's
 * center coordinate and half-size along that axis, `sign` the face (+1 = high
 * face, −1 = low face), `targetA` the desired face coordinate. The face is
 * clamped so the box keeps at least `minThickness`. Returns the new
 * center coordinate and half-size along the axis.
 */
export function moveBoxFace(
  centerA: number,
  half: number,
  sign: 1 | -1,
  targetA: number,
  minThickness: number,
): { centerA: number; half: number } {
  const lo = centerA - half;
  const hi = centerA + half;
  let newLo = lo;
  let newHi = hi;
  if (sign === 1) newHi = Math.max(targetA, lo + minThickness);
  else newLo = Math.min(targetA, hi - minThickness);
  return { centerA: (newLo + newHi) / 2, half: (newHi - newLo) / 2 };
}

/**
 * 'x' | 'y' | 'z' when the normal is (within tolerance) aligned with a world
 * axis (either direction), else null. Gates the glass slider sync — a rotated
 * (oblique) plane has no axis slider.
 */
export function dominantWorldAxis(normal: THREE.Vector3, tol = 1e-3): 'x' | 'y' | 'z' | null {
  const ax = Math.abs(normal.x);
  const ay = Math.abs(normal.y);
  const az = Math.abs(normal.z);
  if (ax > 1 - tol && ay < tol && az < tol) return 'x';
  if (ay > 1 - tol && ax < tol && az < tol) return 'y';
  if (az > 1 - tol && ax < tol && ay < tol) return 'z';
  return null;
}
