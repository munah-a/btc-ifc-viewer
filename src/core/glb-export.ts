/**
 * GLB export (W4.5) — exports the currently visible/isolated model geometry to a
 * binary glTF (.glb) for the native PowerPoint Insert → 3D Models path (offline
 * decks, no add-in).
 *
 * `exportSceneToGlb` wraps three's GLTFExporter with `binary: true` +
 * `onlyVisible: true`, so hide/isolate state is honoured automatically (hidden
 * objects have `.visible = false`). Section-clipped geometry is a render-time
 * material effect, not a mesh change, so it cannot be excluded without CSG —
 * `hasActiveClipping` lets the caller warn the user that clipped-away geometry
 * will still appear in the export.
 *
 * The GLTFExporter needs a browser/DOM, so this module is browser-only; its GLB
 * header validation (`isValidGlb`) is pure and unit-tested (a real exporter
 * round-trip runs in the e2e/preview, and the unit test asserts the magic-byte
 * contract on a hand-built minimal GLB).
 */
import * as THREE from 'three';
import { GLTFExporter, type GLTFExporterOptions } from 'three/examples/jsm/exporters/GLTFExporter.js';

/** The 4-byte magic at the start of every binary glTF: ASCII "glTF". */
export const GLB_MAGIC = 0x46546c67; // 'glTF' little-endian
const GLB_MAGIC_BYTES = [0x67, 0x6c, 0x54, 0x46]; // 'g','l','T','F'

/**
 * Validates that a buffer is a well-formed binary glTF: the "glTF" magic,
 * version 2, and a declared total length matching the buffer. Pure — used by the
 * unit test and as a post-export sanity check.
 */
export function isValidGlb(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < 12) return false;
  const view = new DataView(buffer);
  const bytes = new Uint8Array(buffer, 0, 4);
  for (let i = 0; i < 4; i += 1) {
    if (bytes[i] !== GLB_MAGIC_BYTES[i]) return false;
  }
  const version = view.getUint32(4, true);
  if (version !== 2) return false;
  const declaredLength = view.getUint32(8, true);
  return declaredLength === buffer.byteLength;
}

/**
 * True when any clipping/section plane is active, so the caller can warn that
 * section-clipped geometry is NOT excluded from the export (render-time effect).
 */
export function hasActiveClipping(clipperEnabled: boolean, planeCount: number): boolean {
  return clipperEnabled && planeCount > 0;
}

/**
 * The subset of `@thatopen/fragments`' MeshData this exporter consumes. Fragment
 * geometry lives worker/GPU-side and is NOT exposed as CPU-readable three
 * BufferAttributes — but `model.getItemsGeometry()` returns this CPU-side data.
 * That is the only reliable geometry source for a client-side GLB export.
 */
export interface ExportMeshData {
  transform: THREE.Matrix4;
  positions?: Float32Array | Float64Array;
  indices?: Uint8Array | Uint16Array | Uint32Array;
  normals?: Int16Array;
}

/** A fragments model we can pull visible geometry out of for export. */
export interface ExportableModel {
  /** The model's world object (its matrix positions the whole model, incl. federation offsets). */
  object: THREE.Object3D;
  /** Fetches CPU-side geometry for the given local ids. */
  getItemsGeometry(localIds: number[]): Promise<ExportMeshData[][]>;
  /** All local ids that have geometry. */
  getItemsIdsWithGeometry(): Promise<number[]>;
}

/**
 * Builds a plain-three scene (THREE.Group of standard Meshes) from fragments
 * MeshData. Each mesh's own `transform` is applied, then the model's world
 * matrix, so federation offsets are honoured. Pure (no exporter/DOM) →
 * unit-testable with synthetic MeshData. Returns the group + a dispose().
 */
export function buildMeshDataScene(
  entries: { meshes: ExportMeshData[]; modelMatrix: THREE.Matrix4 }[],
): { group: THREE.Group; meshCount: number; dispose: () => void } {
  const group = new THREE.Group();
  const disposables: { dispose: () => void }[] = [];
  let meshCount = 0;

  for (const { meshes, modelMatrix } of entries) {
    for (const mesh of meshes) {
      if (!mesh.positions || mesh.positions.length === 0) continue;
      const geometry = new THREE.BufferGeometry();
      geometry.setAttribute('position', new THREE.Float32BufferAttribute(Float32Array.from(mesh.positions), 3));
      if (mesh.indices && mesh.indices.length > 0) {
        geometry.setIndex(new THREE.BufferAttribute(Uint32Array.from(mesh.indices), 1));
      }
      // Fragment normals are Int16 (normalized) — recompute for a clean export.
      geometry.computeVertexNormals();
      // Bake mesh transform, then the model's world transform (federation offset).
      geometry.applyMatrix4(mesh.transform);
      geometry.applyMatrix4(modelMatrix);
      const material = new THREE.MeshStandardMaterial({ color: 0xb0b4bd, metalness: 0, roughness: 0.9 });
      group.add(new THREE.Mesh(geometry, material));
      disposables.push(geometry, material);
      meshCount += 1;
    }
  }

  return { group, meshCount, dispose: () => disposables.forEach((d) => d.dispose()) };
}

/** Serializes a three Object3D/Group to a binary GLB via GLTFExporter. */
export async function serializeGroupToGlb(
  group: THREE.Object3D,
  options: GLTFExporterOptions = {},
): Promise<ArrayBuffer> {
  const exporter = new GLTFExporter();
  const result = await exporter.parseAsync(group, { binary: true, ...options });
  if (!(result instanceof ArrayBuffer)) throw new Error('GLB export did not return binary output.');
  if (!isValidGlb(result)) throw new Error('GLB export produced an invalid or empty file.');
  return result;
}

/**
 * Exports the VISIBLE geometry of the given fragments models to a binary GLB.
 * `visibleIds` per model already excludes hidden/isolated-away elements (the
 * caller computes it from the hider state), so hide/isolate carries through.
 * Rejects with a clear error when there is nothing visible to export.
 */
export async function exportModelsToGlb(
  models: { model: ExportableModel; visibleIds: number[] }[],
): Promise<ArrayBuffer> {
  if (models.length === 0) throw new Error('Nothing to export — no visible model geometry.');

  const entries: { meshes: ExportMeshData[]; modelMatrix: THREE.Matrix4 }[] = [];
  for (const { model, visibleIds } of models) {
    if (visibleIds.length === 0) continue;
    model.object.updateWorldMatrix(true, false);
    const perItem = await model.getItemsGeometry(visibleIds);
    const flat: ExportMeshData[] = perItem.flat();
    entries.push({ meshes: flat, modelMatrix: model.object.matrixWorld });
  }

  const scene = buildMeshDataScene(entries);
  if (scene.meshCount === 0) {
    scene.dispose();
    throw new Error('Nothing to export — no visible model geometry.');
  }
  try {
    return await serializeGroupToGlb(scene.group);
  } finally {
    scene.dispose();
  }
}
