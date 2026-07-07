import { describe, expect, it, vi } from 'vitest';
import * as THREE from 'three';
import { buildEdgeOverlays, EdgeGeometryCache, EDGE_THRESHOLD_DEGREES } from '../../src/tools/edges';

const material = () => new THREE.LineBasicMaterial();

const meshObject = (visible = true): THREE.Object3D => {
  const mesh = new THREE.Mesh(new THREE.BoxGeometry(1, 1, 1), new THREE.MeshBasicMaterial());
  mesh.updateMatrixWorld(true);
  mesh.visible = visible;
  return mesh;
};

describe('buildEdgeOverlays', () => {
  it('builds one LineSegments overlay per visible mesh', () => {
    const overlays = buildEdgeOverlays([meshObject(), meshObject()], material());
    expect(overlays).toHaveLength(2);
    expect(overlays[0]).toBeInstanceOf(THREE.LineSegments);
  });

  it('skips objects that are not visible', () => {
    expect(buildEdgeOverlays([meshObject(false)], material())).toHaveLength(0);
  });

  it('overlays copy the mesh world matrix and disable auto-update', () => {
    const mesh = meshObject();
    mesh.position.set(3, 0, 0);
    mesh.updateMatrixWorld(true);
    const [overlay] = buildEdgeOverlays([mesh], material());
    expect(overlay.matrixAutoUpdate).toBe(false);
    expect(overlay.matrix.equals(mesh.matrixWorld)).toBe(true);
  });

  it('uses the documented feature-angle threshold', () => {
    expect(EDGE_THRESHOLD_DEGREES).toBe(35);
  });
});

describe('EdgeGeometryCache (P3 — W5.3)', () => {
  it('builds EdgesGeometry once per source geometry uuid and reuses it', () => {
    const cache = new EdgeGeometryCache();
    const source = new THREE.BoxGeometry(1, 1, 1);
    const a = cache.get(source);
    const b = cache.get(source);
    expect(a).toBe(b); // same cached instance — not rebuilt
    expect(a).toBeInstanceOf(THREE.BufferGeometry);
  });

  it('keys distinct geometries separately', () => {
    const cache = new EdgeGeometryCache();
    const a = cache.get(new THREE.BoxGeometry(1, 1, 1));
    const b = cache.get(new THREE.BoxGeometry(2, 2, 2));
    expect(a).not.toBe(b);
  });

  it('buildEdgeOverlays with a cache reuses geometry across calls (no rebuild)', () => {
    const cache = new EdgeGeometryCache();
    const mesh = meshObject();
    const first = buildEdgeOverlays([mesh], material(), cache);
    const second = buildEdgeOverlays([mesh], material(), cache);
    // The overlays across the two calls share the SAME cached EdgesGeometry —
    // the P3 win: the O(triangles) EdgesGeometry is not rebuilt per call.
    expect(first[0].geometry).toBe(second[0].geometry);
  });

  it('without a cache each call builds fresh geometry (baseline)', () => {
    const mesh = meshObject();
    const first = buildEdgeOverlays([mesh], material());
    const second = buildEdgeOverlays([mesh], material());
    expect(first[0].geometry).not.toBe(second[0].geometry);
  });

  it('dispose frees cached geometries', () => {
    const cache = new EdgeGeometryCache();
    const geo = cache.get(new THREE.BoxGeometry(1, 1, 1));
    const spy = vi.fn();
    geo.dispose = spy;
    cache.dispose();
    expect(spy).toHaveBeenCalled();
  });

  it('prune disposes + drops entries whose source uuid is not in the keep set (E1)', () => {
    const cache = new EdgeGeometryCache();
    const keep = new THREE.BoxGeometry(1, 1, 1);
    const drop = new THREE.BoxGeometry(2, 2, 2);
    const keptEdges = cache.get(keep);
    const droppedEdges = cache.get(drop);
    const keptSpy = vi.fn();
    const droppedSpy = vi.fn();
    keptEdges.dispose = keptSpy;
    droppedEdges.dispose = droppedSpy;

    cache.prune(new Set([keep.uuid]));

    // The dropped entry's GPU geometry is disposed and removed…
    expect(droppedSpy).toHaveBeenCalled();
    // …while the kept entry survives untouched (same cached instance on re-get).
    expect(keptSpy).not.toHaveBeenCalled();
    expect(cache.get(keep)).toBe(keptEdges);
    // A re-get of the dropped source rebuilds a fresh geometry (no visual change).
    expect(cache.get(drop)).not.toBe(droppedEdges);
  });

  it('prune with an empty keep set evicts everything (but leaves the cache usable)', () => {
    const cache = new EdgeGeometryCache();
    const source = new THREE.BoxGeometry(1, 1, 1);
    const first = cache.get(source);
    cache.prune(new Set());
    expect(cache.get(source)).not.toBe(first); // rebuilt after eviction
  });
});
