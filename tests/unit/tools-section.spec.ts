import { describe, expect, it } from 'vitest';
import * as THREE from 'three';
import { sectionBoxPlanes, sectionPlanePercent, sectionPlanePoint } from '../../src/tools/section';

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
