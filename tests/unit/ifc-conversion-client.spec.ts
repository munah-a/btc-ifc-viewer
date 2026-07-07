import { describe, expect, it, vi } from 'vitest';

import { IfcConversionClient, type WorkerFactory } from '../../src/core/ifc-conversion-client';

/**
 * W1 + W2 (W5-fixups) worker-resilience coverage. Uses a hand-rolled fake Worker
 * (no jsdom / no external dep, C1) so we can assert:
 *  - a worker-level `onerror` terminates + nulls the worker so the NEXT convert()
 *    spawns a fresh one (dead workers are never reused);
 *  - cancel() (the timeout path) terminates the worker AND rejects the abandoned
 *    pending job so a retry spawns fresh instead of queueing behind a busy parse.
 */

class FakeWorker {
  onmessage: ((event: MessageEvent) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  posted: unknown[] = [];
  terminated = false;

  postMessage(data: unknown, _transfer?: Transferable[]): void {
    this.posted.push(data);
  }

  terminate(): void {
    this.terminated = true;
  }

  /** Simulate the worker's global error handler firing (a dead worker). */
  emitError(message: string): void {
    this.onerror?.({ message } as ErrorEvent);
  }

  /** Simulate a successful conversion result for the given job id. */
  emitDone(id: string, bytes: number[]): void {
    const buffer = new Uint8Array(bytes).buffer;
    this.onmessage?.({ data: { type: 'done', id, frag: buffer } } as MessageEvent);
  }
}

/** A factory that records every spawned fake worker, in order. */
const trackingFactory = (): { factory: WorkerFactory; spawned: FakeWorker[] } => {
  const spawned: FakeWorker[] = [];
  const factory: WorkerFactory = () => {
    const worker = new FakeWorker();
    spawned.push(worker);
    return worker as unknown as Worker;
  };
  return { factory, spawned };
};

const wasmPath = 'http://localhost/';
const bytes = () => new Uint8Array([1, 2, 3, 4]);

describe('IfcConversionClient — worker resilience (W1/W2)', () => {
  it('W1: a worker onerror rejects the in-flight job and terminates the dead worker', async () => {
    const { factory, spawned } = trackingFactory();
    const client = new IfcConversionClient(factory);

    const promise = client.convert(bytes(), { wasmPath });
    expect(spawned).toHaveLength(1);
    expect(client.isActive).toBe(true);

    spawned[0].emitError('worker crashed');

    await expect(promise).rejects.toThrow('worker crashed');
    // The dead worker was terminated + dropped so it can't be reused.
    expect(spawned[0].terminated).toBe(true);
    expect(client.isActive).toBe(false);
  });

  it('W1: the NEXT convert() after a dead worker spawns a fresh worker (no reuse)', async () => {
    const { factory, spawned } = trackingFactory();
    const client = new IfcConversionClient(factory);

    const first = client.convert(bytes(), { wasmPath });
    spawned[0].emitError('worker crashed');
    await expect(first).rejects.toThrow();

    // A retry must NOT hang on the dead worker — it spawns a new one.
    const second = client.convert(bytes(), { wasmPath });
    expect(spawned).toHaveLength(2);
    expect(spawned[1]).not.toBe(spawned[0]);

    spawned[1].emitDone('ifc-2', [9, 9]);
    await expect(second).resolves.toEqual(new Uint8Array([9, 9]));
  });

  it('W2: cancel() terminates the worker and rejects the abandoned pending job', async () => {
    const { factory, spawned } = trackingFactory();
    const client = new IfcConversionClient(factory);

    const promise = client.convert(bytes(), { wasmPath });
    expect(client.isActive).toBe(true);

    client.cancel('Model loading timed out');

    await expect(promise).rejects.toThrow('Model loading timed out');
    expect(spawned[0].terminated).toBe(true);
    expect(client.isActive).toBe(false);
  });

  it('W2: after cancel(), the next convert() spawns a fresh worker (retry is not queued)', async () => {
    const { factory, spawned } = trackingFactory();
    const client = new IfcConversionClient(factory);

    const first = client.convert(bytes(), { wasmPath });
    client.cancel();
    await expect(first).rejects.toThrow();

    const retry = client.convert(bytes(), { wasmPath });
    expect(spawned).toHaveLength(2);
    spawned[1].emitDone('ifc-2', [7]);
    await expect(retry).resolves.toEqual(new Uint8Array([7]));
  });

  it('W2: a late message from a cancelled worker is ignored (job already removed)', async () => {
    const { factory, spawned } = trackingFactory();
    const client = new IfcConversionClient(factory);

    const promise = client.convert(bytes(), { wasmPath });
    const rejection = expect(promise).rejects.toThrow();
    client.cancel();
    await rejection;

    // A stale 'done' for the cancelled job's id must not throw or resolve anything.
    expect(() => spawned[0].emitDone('ifc-1', [1, 2, 3])).not.toThrow();
  });

  it('terminate() delegates to cancel (frees the worker on teardown)', async () => {
    const { factory, spawned } = trackingFactory();
    const client = new IfcConversionClient(factory);
    const promise = client.convert(bytes(), { wasmPath });
    const rejection = expect(promise).rejects.toThrow();
    client.terminate();
    await rejection;
    expect(spawned[0].terminated).toBe(true);
    expect(client.isActive).toBe(false);
  });

  it('does not spawn a worker until the first convert()', () => {
    const spy = vi.fn(() => new FakeWorker() as unknown as Worker);
    const client = new IfcConversionClient(spy);
    expect(spy).not.toHaveBeenCalled();
    expect(client.isActive).toBe(false);
  });
});
