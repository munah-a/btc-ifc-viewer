import { describe, expect, it } from 'vitest';
import * as THREE from 'three';
import { buildEdgeOverlays, EDGE_THRESHOLD_DEGREES } from '../../src/tools/edges';

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
