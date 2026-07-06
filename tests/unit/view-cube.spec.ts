import { describe, expect, it } from 'vitest';
import * as THREE from 'three';
import {
  getViewCubeAxes,
  getViewCubeNavigationDistance,
  resolveViewCubeCameraUp,
} from '../../src/core/view-cube';

const v = (x: number, y: number, z: number) => new THREE.Vector3(x, y, z);

describe('getViewCubeAxes', () => {
  it('returns world basis axes for the identity quaternion', () => {
    const axes = getViewCubeAxes(new THREE.Quaternion());
    expect(axes.right.toArray()).toEqual([1, 0, 0]);
    expect(axes.up.toArray()).toEqual([0, 1, 0]);
    expect(axes.front.toArray()).toEqual([0, 0, 1]);
  });

  it('rotates axes by the basis (90° about Y sends +X to -Z front to +X)', () => {
    const q = new THREE.Quaternion().setFromAxisAngle(v(0, 1, 0), Math.PI / 2);
    const axes = getViewCubeAxes(q);
    expect(axes.right.x).toBeCloseTo(0, 5);
    expect(axes.right.z).toBeCloseTo(-1, 5);
    expect(axes.front.x).toBeCloseTo(1, 5);
  });
});

describe('getViewCubeNavigationDistance', () => {
  it('pulls back furthest for corners, then edges, then faces', () => {
    const corner = getViewCubeNavigationDistance(10, v(1, 1, 1).normalize());
    const edge = getViewCubeNavigationDistance(10, v(1, 1, 0).normalize());
    const face = getViewCubeNavigationDistance(10, v(1, 0, 0));
    expect(corner).toBe(24.5);
    expect(edge).toBe(22.5);
    expect(face).toBe(20);
    expect(corner).toBeGreaterThan(edge);
    expect(edge).toBeGreaterThan(face);
  });
});

describe('resolveViewCubeCameraUp', () => {
  const axes = getViewCubeAxes(new THREE.Quaternion());

  it('top view (looking down) uses the front axis as up', () => {
    // localDirection +Y, world look direction points down (-Y from above).
    const up = resolveViewCubeCameraUp(v(0, 1, 0), v(0, 1, 0), axes);
    // orthogonalized against worldDirection (+Y) -> front (+Z) survives.
    expect(Math.abs(up.z)).toBeCloseTo(1, 5);
    expect(up.y).toBeCloseTo(0, 5);
  });

  it('front view keeps +Y up', () => {
    const up = resolveViewCubeCameraUp(v(0, 0, 1), v(0, 0, 1), axes);
    expect(up.y).toBeCloseTo(1, 5);
  });

  it('returns a unit vector', () => {
    const up = resolveViewCubeCameraUp(v(1, 1, 1).normalize(), v(1, 1, 1).normalize(), axes);
    expect(up.length()).toBeCloseTo(1, 5);
  });
});
