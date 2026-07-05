/**
 * Load-lifecycle bookkeeping for models (AUDIT A6/A10/F6, reshaped further in W2.2).
 *
 * Metadata is keyed by the model id passed to `ifcLoader.load` (the library uses
 * that name verbatim as `model.modelId` and as the `fragments.list` key), which
 * kills the FIFO metadata queue that could mis-attribute file names, and the
 * alias-resolution layer that papered over it.
 *
 * DOM-free and engine-free on purpose so it is unit-testable (tests/unit).
 */

export interface ModelLoadMeta {
  fileName: string;
  sizeBytes: number;
}

export class ModelRegistry {
  private readonly pending = new Map<string, ModelLoadMeta>();
  private readonly stale = new Set<string>();

  /**
   * Picks a unique model id for a file. The plain file name is used when free;
   * duplicate loads get a ` (n)` suffix so two files with the same name never
   * collide inside the engine's model list.
   *
   * @param fileName - the uploaded file's name.
   * @param activeIds - ids currently taken in the engine/UI (e.g. fragments.list keys).
   */
  allocateModelId(fileName: string, activeIds: Iterable<string>): string {
    const taken = new Set<string>(activeIds);
    const isTaken = (id: string): boolean =>
      taken.has(id) || this.pending.has(id) || this.stale.has(id);
    if (!isTaken(fileName)) return fileName;
    let suffix = 2;
    while (isTaken(`${fileName} (${suffix})`)) suffix += 1;
    return `${fileName} (${suffix})`;
  }

  /** Records the metadata for a load that is about to start. */
  beginLoad(modelId: string, meta: ModelLoadMeta): void {
    this.pending.set(modelId, meta);
  }

  /** Consumes and returns the metadata recorded for a completed load. */
  completeLoad(modelId: string): ModelLoadMeta | undefined {
    const meta = this.pending.get(modelId);
    this.pending.delete(modelId);
    return meta;
  }

  /** Drops the pending record for a load that failed before registration. */
  failLoad(modelId: string): void {
    this.pending.delete(modelId);
  }

  /**
   * Marks a timed-out load as stale (AUDIT A10): if the engine finishes the
   * load later, the late arrival must be disposed instead of registered.
   */
  markStale(modelId: string): void {
    this.pending.delete(modelId);
    this.stale.add(modelId);
  }

  isStale(modelId: string): boolean {
    return this.stale.has(modelId);
  }

  /**
   * Consumes a stale marker. Returns true when the id was stale — the caller
   * disposes the late-arriving model exactly once.
   */
  consumeStale(modelId: string): boolean {
    return this.stale.delete(modelId);
  }

  /** True while a load for this id is in flight. */
  isPending(modelId: string): boolean {
    return this.pending.has(modelId);
  }

  /** Resets all bookkeeping (viewer teardown). */
  clear(): void {
    this.pending.clear();
    this.stale.clear();
  }
}
