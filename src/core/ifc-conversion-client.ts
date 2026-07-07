/**
 * Main-thread client for the IFC→fragments conversion worker (P4, W5.4).
 *
 * Lazily spawns the dedicated worker on first use and marshals one conversion at
 * a time (the app already serializes loads via `isModelLoading`). Converts IFC
 * bytes → `.frag` bytes off the main thread so the UI stays responsive on big
 * models. Progress is forwarded to a per-call callback.
 *
 * The worker is created via `new Worker(new URL(...), { type: 'module' })` so
 * Vite bundles it as a separate chunk (dev + prod). The wasm path is passed in
 * (self-hosted absolute URL, C1/A2). `terminate()` frees the worker on teardown.
 */

export interface IfcConversionOptions {
  /** Absolute URL to the self-hosted web-ifc wasm directory (C1). */
  wasmPath: string;
  addAllAttributes?: boolean;
  addAllRelations?: boolean;
  /** 0..1 conversion progress. */
  onProgress?: (progress: number) => void;
}

interface PendingJob {
  resolve: (frag: Uint8Array) => void;
  reject: (error: Error) => void;
  onProgress?: (progress: number) => void;
}

type WorkerResponse =
  | { type: 'progress'; id: string; progress: number }
  | { type: 'done'; id: string; frag: ArrayBuffer }
  | { type: 'error'; id: string; message: string };

/** Factory for the worker — injectable so unit tests can supply a fake. */
export type WorkerFactory = () => Worker;

const defaultWorkerFactory: WorkerFactory = () =>
  new Worker(new URL('../workers/ifc-conversion.worker.ts', import.meta.url), {
    type: 'module',
  });

export class IfcConversionClient {
  private worker: Worker | null = null;
  private readonly pending = new Map<string, PendingJob>();
  private seq = 0;

  constructor(private readonly workerFactory: WorkerFactory = defaultWorkerFactory) {}

  /** Whether a worker is currently spawned (used by tests / teardown checks). */
  get isActive(): boolean {
    return this.worker !== null;
  }

  private ensureWorker(): Worker {
    if (this.worker) return this.worker;
    const worker = this.workerFactory();
    worker.onmessage = (event: MessageEvent<WorkerResponse>) => this.onMessage(event.data);
    worker.onerror = (event: ErrorEvent) => this.onWorkerError(event.message || 'IFC conversion worker error');
    this.worker = worker;
    return worker;
  }

  private onMessage(data: WorkerResponse): void {
    const job = this.pending.get(data.id);
    if (!job) return;
    if (data.type === 'progress') {
      job.onProgress?.(data.progress);
      return;
    }
    this.pending.delete(data.id);
    if (data.type === 'done') {
      job.resolve(new Uint8Array(data.frag));
    } else {
      job.reject(new Error(data.message));
    }
  }

  /** A worker-level failure rejects every in-flight job (rare; e.g. worker crash). */
  private onWorkerError(message: string): void {
    for (const job of this.pending.values()) job.reject(new Error(message));
    this.pending.clear();
  }

  /**
   * Converts IFC bytes to `.frag` bytes on the worker. The input buffer is
   * transferred (not copied) to the worker.
   */
  convert(bytes: Uint8Array, options: IfcConversionOptions): Promise<Uint8Array> {
    const worker = this.ensureWorker();
    const id = `ifc-${++this.seq}`;
    return new Promise<Uint8Array>((resolve, reject) => {
      this.pending.set(id, { resolve, reject, onProgress: options.onProgress });
      // Copy into a fresh buffer we own, then transfer it (avoids detaching a
      // buffer the caller may still hold, e.g. the isProbablyIfc-checked array).
      const owned = bytes.slice();
      worker.postMessage(
        {
          id,
          bytes: owned.buffer,
          wasmPath: options.wasmPath,
          addAllAttributes: options.addAllAttributes ?? true,
          addAllRelations: options.addAllRelations ?? true,
        },
        [owned.buffer],
      );
    });
  }

  terminate(): void {
    this.worker?.terminate();
    this.worker = null;
    this.onWorkerError('IFC conversion client terminated');
  }
}
