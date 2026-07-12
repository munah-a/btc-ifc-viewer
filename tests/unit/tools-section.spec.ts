import { describe, expect, it } from 'vitest';
import * as THREE from 'three';
import {
  axisDragOffset,
  axisRangeOfBox,
  dominantWorldAxis,
  intersectRayPlane,
  moveBoxFace,
  orientedBoxFromBox3,
  orientedSectionBoxPlanes,
  planeBasis,
  planeQuadRect,
  sectionBoxFaceCenter,
  sectionBoxFaceNormal,
  sectionBoxFromPlanes,
  sectionBoxPlanes,
  sectionPlanePercent,
  sectionPlanePoint,
  signedAngleAroundAxis,
  type OrientedSectionBox,
} from '../../src/tools/section';

const box = () => new THREE.Box3(new THREE.Vector3(-2, 0, 10), new THREE.Vector3(4, 6, 20));

describe('sectionPlanePoint', () => {
  it('slides along X for an X-normal at 0/50/100%', () => {
    const b = box();
    const n = new THREE.Vector3(-1, 0, 0);
    expect(sectionPlanePoint(b, n, 0).x).toBeCloseTo(-2, 5);
    expect(sectionPlanePoint(b, n, 50).x).toBeCloseTo(1, 5);
    expect(sectionPlanePoint(b, n, 100).x).toBeCloseTo(4, 5);
  });

  it('non-dominant axes stay at the box center', () => {
    const b = box();
    const point = sectionPlanePoint(b, new THREE.Vector3(-1, 0, 0), 25);
    const center = b.getCenter(new THREE.Vector3());
    expect(point.y).toBeCloseTo(center.y, 5);
    expect(point.z).toBeCloseTo(center.z, 5);
  });

  it('slides along Y for a Y-normal and Z for a Z-normal', () => {
    const b = box();
    expect(sectionPlanePoint(b, new THREE.Vector3(0, -1, 0), 100).y).toBeCloseTo(6, 5);
    expect(sectionPlanePoint(b, new THREE.Vector3(0, 0, -1), 0).z).toBeCloseTo(10, 5);
  });

  it('clamps pct outside 0–100', () => {
    const b = box();
    const n = new THREE.Vector3(-1, 0, 0);
    expect(sectionPlanePoint(b, n, -50).x).toBeCloseTo(-2, 5);
    expect(sectionPlanePoint(b, n, 150).x).toBeCloseTo(4, 5);
  });
});

describe('sectionPlanePercent (C5 — W5-fixups: inverse of sectionPlanePoint)', () => {
  it('round-trips sectionPlanePoint at 0/25/50/75/100 for an X-normal', () => {
    const b = box();
    const n = new THREE.Vector3(-1, 0, 0);
    for (const pct of [0, 25, 50, 75, 100]) {
      const origin = sectionPlanePoint(b, n, pct);
      expect(sectionPlanePercent(b, n, origin)).toBeCloseTo(pct, 4);
    }
  });

  it('inverts along Y and Z for their respective normals', () => {
    const b = box();
    const yOrigin = sectionPlanePoint(b, new THREE.Vector3(0, -1, 0), 30);
    expect(sectionPlanePercent(b, new THREE.Vector3(0, -1, 0), yOrigin)).toBeCloseTo(30, 4);
    const zOrigin = sectionPlanePoint(b, new THREE.Vector3(0, 0, -1), 80);
    expect(sectionPlanePercent(b, new THREE.Vector3(0, 0, -1), zOrigin)).toBeCloseTo(80, 4);
  });

  it('clamps origins outside the box to 0–100', () => {
    const b = box();
    const n = new THREE.Vector3(-1, 0, 0);
    expect(sectionPlanePercent(b, n, new THREE.Vector3(-100, 3, 15))).toBe(0);
    expect(sectionPlanePercent(b, n, new THREE.Vector3(100, 3, 15))).toBe(100);
  });

  it('returns the neutral mid (50) for a zero-extent axis', () => {
    const flat = new THREE.Box3(new THREE.Vector3(5, 0, 0), new THREE.Vector3(5, 6, 20));
    expect(sectionPlanePercent(flat, new THREE.Vector3(-1, 0, 0), new THREE.Vector3(5, 3, 10))).toBe(50);
  });
});

describe('sectionBoxPlanes', () => {
  it('returns six planes hugging the box extremes', () => {
    const planes = sectionBoxPlanes(box());
    expect(planes).toHaveLength(6);
    // -X plane sits at max.x; +X plane at min.x (each plane clips the opposite side).
    expect(planes[0].normal.toArray()).toEqual([-1, 0, 0]);
    expect(planes[0].point.x).toBeCloseTo(4, 5);
    expect(planes[1].normal.toArray()).toEqual([1, 0, 0]);
    expect(planes[1].point.x).toBeCloseTo(-2, 5);
    expect(planes[4].point.z).toBeCloseTo(20, 5);
    expect(planes[5].point.z).toBeCloseTo(10, 5);
  });
});

// ---------------------------------------------------------------------------
// Autodesk-style gizmo math
// ---------------------------------------------------------------------------

/** Canonical key for a plane (unit normal + signed offset), for set comparison. */
const planeKey = (normal: THREE.Vector3, point: THREE.Vector3): string => {
  const n = normal.clone().normalize();
  const d = n.dot(point);
  const r = (value: number) => (Math.abs(value) < 1e-6 ? 0 : Number(value.toFixed(6)));
  return `${r(n.x)},${r(n.y)},${r(n.z)}:${r(d)}`;
};

const planeKeySet = (planes: Array<{ normal: THREE.Vector3; point: THREE.Vector3 }>): Set<string> =>
  new Set(planes.map((p) => planeKey(p.normal, p.point)));

const toRecords = (planes: Array<{ normal: THREE.Vector3; point: THREE.Vector3 }>) =>
  planes.map((p) => ({
    normal: { x: p.normal.x, y: p.normal.y, z: p.normal.z },
    origin: { x: p.point.x, y: p.point.y, z: p.point.z },
  }));

describe('orientedBoxFromBox3 / orientedSectionBoxPlanes', () => {
  it('covers the Box3 with world axes and inward clip normals', () => {
    const oriented = orientedBoxFromBox3(box());
    expect(oriented.center.toArray()).toEqual([1, 3, 15]);
    expect(oriented.halfSizes.toArray()).toEqual([3, 3, 5]);
    const planes = orientedSectionBoxPlanes(oriented);
    expect(planes).toHaveLength(6);
    // Same clipping volume as the legacy sectionBoxPlanes definition.
    expect(planeKeySet(planes)).toEqual(planeKeySet(sectionBoxPlanes(box())));
  });

  it('face normals point outward and face centers sit on the surface', () => {
    const oriented = orientedBoxFromBox3(box());
    expect(sectionBoxFaceNormal(oriented, 0).toArray()).toEqual([1, 0, 0]);
    expect(sectionBoxFaceNormal(oriented, 1).distanceTo(new THREE.Vector3(-1, 0, 0))).toBeLessThan(1e-9);
    expect(sectionBoxFaceCenter(oriented, 0).x).toBeCloseTo(4, 6);
    expect(sectionBoxFaceCenter(oriented, 3).y).toBeCloseTo(0, 6);
    expect(sectionBoxFaceCenter(oriented, 4).z).toBeCloseTo(20, 6);
  });
});

describe('sectionBoxFromPlanes (C8 box reconstruction)', () => {
  it('round-trips an axis-aligned box (any plane order)', () => {
    const oriented = orientedBoxFromBox3(box());
    const planes = orientedSectionBoxPlanes(oriented);
    const shuffled = [planes[3], planes[0], planes[5], planes[2], planes[1], planes[4]];
    const result = sectionBoxFromPlanes(toRecords(shuffled));
    expect(result).not.toBeNull();
    expect(planeKeySet(orientedSectionBoxPlanes(result!.box))).toEqual(planeKeySet(planes));
    expect(result!.box.center.distanceTo(oriented.center)).toBeLessThan(1e-6);
  });

  it('round-trips a rotated (oblique) box', () => {
    const q = new THREE.Quaternion().setFromAxisAngle(new THREE.Vector3(1, 2, 3).normalize(), 0.7);
    const oriented: OrientedSectionBox = {
      center: new THREE.Vector3(5, -2, 8),
      halfSizes: new THREE.Vector3(2, 4, 1.5),
      axes: [
        new THREE.Vector3(1, 0, 0).applyQuaternion(q),
        new THREE.Vector3(0, 1, 0).applyQuaternion(q),
        new THREE.Vector3(0, 0, 1).applyQuaternion(q),
      ],
    };
    const planes = orientedSectionBoxPlanes(oriented);
    const result = sectionBoxFromPlanes(toRecords(planes));
    expect(result).not.toBeNull();
    expect(planeKeySet(orientedSectionBoxPlanes(result!.box))).toEqual(planeKeySet(planes));
    expect(result!.box.center.distanceTo(oriented.center)).toBeLessThan(1e-6);
    // The reconstructed basis stays right-handed.
    const [a0, a1, a2] = result!.box.axes;
    expect(new THREE.Vector3().crossVectors(a0, a1).dot(a2)).toBeCloseTo(1, 5);
  });

  it('maps each face to its source plane via facePlaneIndices', () => {
    const oriented = orientedBoxFromBox3(box());
    const planes = orientedSectionBoxPlanes(oriented);
    const shuffled = [planes[4], planes[1], planes[0], planes[5], planes[3], planes[2]];
    const result = sectionBoxFromPlanes(toRecords(shuffled));
    expect(result).not.toBeNull();
    const rebuilt = orientedSectionBoxPlanes(result!.box);
    for (let face = 0; face < 6; face += 1) {
      const source = shuffled[result!.facePlaneIndices[face]];
      expect(planeKey(rebuilt[face].normal, rebuilt[face].point)).toBe(planeKey(source.normal, source.point));
    }
  });

  it('rejects wrong counts, unpaired, non-orthogonal and inverted plane sets', () => {
    const oriented = orientedBoxFromBox3(box());
    const planes = orientedSectionBoxPlanes(oriented);
    // Five planes.
    expect(sectionBoxFromPlanes(toRecords(planes.slice(0, 5)))).toBeNull();
    // Unpaired: replace one plane's normal so its axis has no antiparallel partner.
    const unpaired = planes.map((p, i) =>
      i === 1 ? { normal: new THREE.Vector3(0, 1, 0), point: p.point.clone() } : p,
    );
    expect(sectionBoxFromPlanes(toRecords(unpaired))).toBeNull();
    // Non-orthogonal: rotate only the X pair by 30°.
    const rot = new THREE.Quaternion().setFromAxisAngle(new THREE.Vector3(0, 0, 1), Math.PI / 6);
    const skewed = planes.map((p, i) =>
      i < 2 ? { normal: p.normal.clone().applyQuaternion(rot), point: p.point.clone() } : p,
    );
    expect(sectionBoxFromPlanes(toRecords(skewed))).toBeNull();
    // Inverted (outward normals keep nothing): every extent reads hi < lo.
    const inverted = planes.map((p) => ({ normal: p.normal.clone().negate(), point: p.point.clone() }));
    expect(sectionBoxFromPlanes(toRecords(inverted))).toBeNull();
  });
});

describe('axisDragOffset', () => {
  it('returns the axis parameter closest to the pointer ray', () => {
    // Axis along X through origin; ray pointing straight down onto (3, 0, 0).
    const t = axisDragOffset(new THREE.Vector3(0, 0, 0), new THREE.Vector3(1, 0, 0), {
      origin: new THREE.Vector3(3, 10, 0),
      direction: new THREE.Vector3(0, -1, 0),
    });
    expect(t).not.toBeNull();
    expect(t!).toBeCloseTo(3, 6);
  });

  it('is null for a ray parallel to the axis', () => {
    const t = axisDragOffset(new THREE.Vector3(0, 0, 0), new THREE.Vector3(1, 0, 0), {
      origin: new THREE.Vector3(0, 5, 0),
      direction: new THREE.Vector3(1, 0, 0),
    });
    expect(t).toBeNull();
  });
});

describe('intersectRayPlane / signedAngleAroundAxis', () => {
  it('intersects a ray with a plane', () => {
    const hit = intersectRayPlane(
      { origin: new THREE.Vector3(0, 5, 0), direction: new THREE.Vector3(0, -1, 0) },
      new THREE.Vector3(0, 1, 0),
      new THREE.Vector3(0, 1, 0),
    );
    expect(hit).not.toBeNull();
    expect(hit!.y).toBeCloseTo(1, 6);
  });

  it('is null for parallel or receding rays', () => {
    const parallel = intersectRayPlane(
      { origin: new THREE.Vector3(0, 5, 0), direction: new THREE.Vector3(1, 0, 0) },
      new THREE.Vector3(0, 1, 0),
      new THREE.Vector3(0, 1, 0),
    );
    expect(parallel).toBeNull();
    const receding = intersectRayPlane(
      { origin: new THREE.Vector3(0, 5, 0), direction: new THREE.Vector3(0, 1, 0) },
      new THREE.Vector3(0, 1, 0),
      new THREE.Vector3(0, 1, 0),
    );
    expect(receding).toBeNull();
  });

  it('measures signed angles around an axis', () => {
    const axis = new THREE.Vector3(0, 0, 1);
    expect(signedAngleAroundAxis(axis, new THREE.Vector3(1, 0, 0), new THREE.Vector3(0, 1, 0))).toBeCloseTo(
      Math.PI / 2,
      6,
    );
    expect(signedAngleAroundAxis(axis, new THREE.Vector3(0, 1, 0), new THREE.Vector3(1, 0, 0))).toBeCloseTo(
      -Math.PI / 2,
      6,
    );
    // Components along the axis are ignored (projection).
    expect(
      signedAngleAroundAxis(axis, new THREE.Vector3(1, 0, 4), new THREE.Vector3(0, 1, -2)),
    ).toBeCloseTo(Math.PI / 2, 6);
  });
});

describe('planeBasis / planeQuadRect / axisRangeOfBox', () => {
  it('builds a right-handed orthonormal basis for any normal', () => {
    for (const normal of [
      new THREE.Vector3(1, 0, 0),
      new THREE.Vector3(0, 1, 0),
      new THREE.Vector3(0, 0, 1),
      new THREE.Vector3(1, 2, 3).normalize(),
    ]) {
      const { u, v } = planeBasis(normal);
      expect(u.length()).toBeCloseTo(1, 6);
      expect(v.length()).toBeCloseTo(1, 6);
      expect(u.dot(normal)).toBeCloseTo(0, 6);
      expect(v.dot(normal)).toBeCloseTo(0, 6);
      expect(new THREE.Vector3().crossVectors(u, v).dot(normal)).toBeCloseTo(1, 6);
    }
  });

  it('sizes the quad to the projected model bounds plus margin', () => {
    const b = box(); // spans 6 × 6 × 10
    const normal = new THREE.Vector3(1, 0, 0);
    const { u, v } = planeBasis(normal);
    const origin = b.getCenter(new THREE.Vector3());
    const rect = planeQuadRect(b, origin, u, v, 0.05);
    expect(rect).not.toBeNull();
    // In-plane extents for an X-normal are Y (6) and Z (10), regardless of
    // which of u/v carries which.
    const spans = [rect!.halfU * 2, rect!.halfV * 2].sort((a, c) => a - c);
    expect(spans[0]).toBeCloseTo(6 * 1.05, 5);
    expect(spans[1]).toBeCloseTo(10 * 1.05, 5);
    expect(rect!.center.distanceTo(origin)).toBeLessThan(1e-6);
    expect(planeQuadRect(new THREE.Box3(), origin, u, v)).toBeNull();
  });

  it('projects box corners onto an arbitrary axis', () => {
    const range = axisRangeOfBox(
      new THREE.Box3(new THREE.Vector3(0, 0, 0), new THREE.Vector3(1, 1, 1)),
      new THREE.Vector3(1, 1, 1).normalize(),
    );
    expect(range.lo).toBeCloseTo(0, 6);
    expect(range.hi).toBeCloseTo(Math.sqrt(3), 6);
  });
});

describe('moveBoxFace / dominantWorldAxis', () => {
  it('moves the high and low faces independently', () => {
    // Box along axis: center 0, half 5 → faces at −5 / +5.
    const hi = moveBoxFace(0, 5, 1, 8, 0.5);
    expect(hi.centerA + hi.half).toBeCloseTo(8, 6);
    expect(hi.centerA - hi.half).toBeCloseTo(-5, 6);
    const lo = moveBoxFace(0, 5, -1, -2, 0.5);
    expect(lo.centerA - lo.half).toBeCloseTo(-2, 6);
    expect(lo.centerA + lo.half).toBeCloseTo(5, 6);
  });

  it('clamps to the minimum thickness instead of inverting', () => {
    const clampedHi = moveBoxFace(0, 5, 1, -20, 1);
    expect(clampedHi.half * 2).toBeCloseTo(1, 6);
    expect(clampedHi.centerA - clampedHi.half).toBeCloseTo(-5, 6);
    const clampedLo = moveBoxFace(0, 5, -1, 20, 1);
    expect(clampedLo.half * 2).toBeCloseTo(1, 6);
    expect(clampedLo.centerA + clampedLo.half).toBeCloseTo(5, 6);
  });

  it('detects axis-aligned normals with tolerance', () => {
    expect(dominantWorldAxis(new THREE.Vector3(-1, 0, 0))).toBe('x');
    expect(dominantWorldAxis(new THREE.Vector3(0, 1, 0))).toBe('y');
    expect(dominantWorldAxis(new THREE.Vector3(0, 0.0005, -0.9999995))).toBe('z');
    expect(dominantWorldAxis(new THREE.Vector3(1, 1, 0).normalize())).toBeNull();
  });
});
