import { describe, expect, it } from 'vitest';

import { ModelRegistry } from '../../src/core/model-registry';

describe('ModelRegistry (AUDIT A6/A10/F6 — W1.4)', () => {
  it('keys metadata by model id, not FIFO order (A6)', () => {
    const registry = new ModelRegistry();
    registry.beginLoad('a.ifc', { fileName: 'a.ifc', sizeBytes: 100 });
    registry.beginLoad('b.ifc', { fileName: 'b.ifc', sizeBytes: 200 });

    // Loads can finish out of order — attribution must not shift.
    expect(registry.completeLoad('b.ifc')).toEqual({ fileName: 'b.ifc', sizeBytes: 200 });
    expect(registry.completeLoad('a.ifc')).toEqual({ fileName: 'a.ifc', sizeBytes: 100 });
    expect(registry.completeLoad('a.ifc')).toBeUndefined();
  });

  it('allocates unique ids for duplicate file names', () => {
    const registry = new ModelRegistry();
    const first = registry.allocateModelId('model.ifc', []);
    expect(first).toBe('model.ifc');
    registry.beginLoad(first, { fileName: 'model.ifc', sizeBytes: 1 });

    const second = registry.allocateModelId('model.ifc', []);
    expect(second).toBe('model.ifc (2)');
    registry.beginLoad(second, { fileName: 'model.ifc', sizeBytes: 1 });

    const third = registry.allocateModelId('model.ifc', ['model.ifc (3)']);
    expect(third).toBe('model.ifc (4)');
  });

  it('respects engine-active ids when allocating', () => {
    const registry = new ModelRegistry();
    expect(registry.allocateModelId('taken.ifc', ['taken.ifc'])).toBe('taken.ifc (2)');
  });

  it('marks timed-out loads stale and consumes the marker once (A10)', () => {
    const registry = new ModelRegistry();
    registry.beginLoad('slow.ifc', { fileName: 'slow.ifc', sizeBytes: 5 });
    registry.markStale('slow.ifc');

    expect(registry.isPending('slow.ifc')).toBe(false);
    expect(registry.isStale('slow.ifc')).toBe(true);
    // The stale id stays reserved so a retry can't collide with the ghost.
    expect(registry.allocateModelId('slow.ifc', [])).toBe('slow.ifc (2)');

    // Late arrival disposed exactly once.
    expect(registry.consumeStale('slow.ifc')).toBe(true);
    expect(registry.consumeStale('slow.ifc')).toBe(false);
    expect(registry.isStale('slow.ifc')).toBe(false);
  });

  it('frees ids after unload so the plain name is reusable (F6)', () => {
    const registry = new ModelRegistry();
    const id = registry.allocateModelId('model.ifc', []);
    registry.beginLoad(id, { fileName: 'model.ifc', sizeBytes: 1 });
    registry.completeLoad(id);

    // While loaded, the engine list owns the id.
    expect(registry.allocateModelId('model.ifc', [id])).toBe('model.ifc (2)');
    // After unload (id no longer active anywhere) it is reusable.
    expect(registry.allocateModelId('model.ifc', [])).toBe('model.ifc');
  });

  it('drops pending metadata for failed loads', () => {
    const registry = new ModelRegistry();
    registry.beginLoad('bad.ifc', { fileName: 'bad.ifc', sizeBytes: 9 });
    registry.failLoad('bad.ifc');
    expect(registry.isPending('bad.ifc')).toBe(false);
    expect(registry.completeLoad('bad.ifc')).toBeUndefined();
    expect(registry.allocateModelId('bad.ifc', [])).toBe('bad.ifc');
  });

  it('clear() resets all bookkeeping', () => {
    const registry = new ModelRegistry();
    registry.beginLoad('x.ifc', { fileName: 'x.ifc', sizeBytes: 1 });
    registry.markStale('x.ifc');
    registry.beginLoad('y.ifc', { fileName: 'y.ifc', sizeBytes: 2 });
    registry.clear();
    expect(registry.isStale('x.ifc')).toBe(false);
    expect(registry.isPending('y.ifc')).toBe(false);
  });
});
