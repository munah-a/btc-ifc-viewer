/**
 * Dedicated IFC→fragments conversion worker (P4, W5.4).
 *
 * The web-ifc parse + fragments serialization (`IfcImporter.process`) is
 * CPU-heavy and previously ran on the MAIN thread inside `IfcLoader.load`,
 * freezing the UI on large models (the @thatopen fragments worker only handles
 * streaming/culling — it does NOT do the IFC parse, verified against
 * public/worker.mjs). This worker moves that work off the main thread so the UI
 * stays responsive; the produced `.frag` bytes are sent back and loaded via the
 * normal `fragments.core.load` streaming path.
 *
 * C1/A2: the web-ifc WASM path is passed in from the main thread as an ABSOLUTE
 * URL (derived from the self-hosted BASE_URL) so the worker fetches the vendored
 * wasm — never a CDN.
 *
 * Protocol (main → worker): { id, bytes, wasmPath, addAllAttributes,
 * addAllRelations }. Worker → main: { type:'progress', id, progress } … then
 * { type:'done', id, frag } (frag is a transferred ArrayBuffer) or
 * { type:'error', id, message }.
 */
import { IfcImporter } from '@thatopen/fragments';

interface ConvertRequest {
  id: string;
  bytes: ArrayBuffer;
  wasmPath: string;
  addAllAttributes: boolean;
  addAllRelations: boolean;
}

type WorkerResponse =
  | { type: 'progress'; id: string; progress: number }
  | { type: 'done'; id: string; frag: ArrayBuffer }
  | { type: 'error'; id: string; message: string };

const post = (message: WorkerResponse, transfer?: Transferable[]): void => {
  (self as unknown as Worker).postMessage(message, transfer ?? []);
};

self.onmessage = async (event: MessageEvent<ConvertRequest>): Promise<void> => {
  const { id, bytes, wasmPath, addAllAttributes, addAllRelations } = event.data;
  try {
    const importer = new IfcImporter();
    importer.wasm.path = wasmPath;
    importer.wasm.absolute = true;
    // Mirror the main-thread instanceCallback (viewer.ts loadIfcFile): pull in
    // all attributes + relations so the property/spatial index has full data.
    if (addAllAttributes) importer.addAllAttributes();
    if (addAllRelations) importer.addAllRelations();

    const frag = await importer.process({
      bytes: new Uint8Array(bytes),
      progressCallback: (progress: number) => {
        post({ type: 'progress', id, progress });
      },
    });

    // Transfer a fresh ArrayBuffer so the bytes aren't copied back.
    const buffer = new ArrayBuffer(frag.byteLength);
    new Uint8Array(buffer).set(frag);
    post({ type: 'done', id, frag: buffer }, [buffer]);
  } catch (error) {
    post({
      type: 'error',
      id,
      message: error instanceof Error ? error.message : String(error),
    });
  }
};
