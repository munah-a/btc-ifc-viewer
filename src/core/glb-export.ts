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
 * Exports the given model objects to a binary GLB. `onlyVisible` (default true)
 * makes hide/isolate state carry through. Rejects with a clear error if the
 * export produces an empty/invalid buffer (e.g. nothing visible to export).
 */
export async function exportObjectsToGlb(
  objects: THREE.Object3D[],
  options: GLTFExporterOptions = {},
): Promise<ArrayBuffer> {
  if (objects.length === 0) {
    throw new Error('Nothing to export — no visible model geometry.');
  }
  const exporter = new GLTFExporter();
  const result = await exporter.parseAsync(objects, {
    binary: true,
    onlyVisible: true,
    ...options,
  });
  if (!(result instanceof ArrayBuffer)) {
    throw new Error('GLB export did not return binary output.');
  }
  if (!isValidGlb(result)) {
    throw new Error('GLB export produced an invalid or empty file.');
  }
  return result;
}
