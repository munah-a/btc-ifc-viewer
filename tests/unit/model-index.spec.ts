import { describe, expect, it } from 'vitest';

import { buildModelIndex } from '../../src/core/model-index';
import type { FragmentsModelLike } from '../../src/core/fragments-model';

/**
 * M1 (W5-fixups): the per-chunk getItemsData round-trips are now issued in
 * bounded batches (INDEX_CHUNK_CONCURRENCY) and folded IN CHUNK ORDER, instead
 * of one unbounded Promise.all. This must:
 *  - keep the built maps deterministic (identical to a sequential build);
 *  - cap the number of concurrent getItemsData calls in flight.
 *
 * Uses a minimal fake of the FragmentsModelLike boundary (only the methods
 * buildModelIndex touches), with enough items to span many CHUNK_SIZE (=360)
 * chunks and more than one concurrency batch.
 */

const CHUNK_SIZE = 360;
const ITEM_COUNT = CHUNK_SIZE * 30; // 30 chunks → several batches of 12

interface Tracker {
  inFlight: number;
  peak: number;
}

const makeFakeModel = (tracker: Tracker): FragmentsModelLike => {
  const ids = Array.from({ length: ITEM_COUNT }, (_, i) => i + 1);
  const spatial = {
    category: 'IFCPROJECT',
    localId: 0,
    children: [] as unknown[],
  };
  const model = {
    modelId: 'm1',
    getItemsIdsWithGeometry: () => Promise.resolve(ids),
    getItemsWithGeometryCategories: () => Promise.resolve(ids.map(() => 'IfcWall')),
    getSpatialStructure: () => Promise.resolve(spatial),
    getItemsData: async (chunk: number[]) => {
      tracker.inFlight += 1;
      tracker.peak = Math.max(tracker.peak, tracker.inFlight);
      // Yield so overlapping in-flight calls actually accumulate.
      await Promise.resolve();
      await Promise.resolve();
      tracker.inFlight -= 1;
      // A stable per-id name so we can assert deterministic ordering.
      return chunk.map((id) => ({
        _localId: { value: id },
        Name: { value: `Item ${id}` },
      }));
    },
  };
  return model as unknown as FragmentsModelLike;
};

describe('buildModelIndex — bounded chunk concurrency (M1 — W5-fixups)', () => {
  it('caps concurrent getItemsData calls (does not fan out unbounded)', async () => {
    const tracker: Tracker = { inFlight: 0, peak: 0 };
    await buildModelIndex('m1', makeFakeModel(tracker));
    // 30 chunks would all be in flight at once with the old unbounded Promise.all.
    // The bounded batch keeps peak at the concurrency cap (12).
    expect(tracker.peak).toBeGreaterThan(0);
    expect(tracker.peak).toBeLessThanOrEqual(12);
  });

  it('builds a deterministic index (all ids present, stable names)', async () => {
    const a = await buildModelIndex('m1', makeFakeModel({ inFlight: 0, peak: 0 }));
    const b = await buildModelIndex('m1', makeFakeModel({ inFlight: 0, peak: 0 }));

    expect(a.allIds.size).toBe(ITEM_COUNT);
    // Two independent builds agree on every item name (order-independent map).
    expect(a.itemNames.get(1)).toBe('Item 1');
    expect(a.itemNames.get(ITEM_COUNT)).toBe(`Item ${ITEM_COUNT}`);
    expect([...a.itemNames.entries()].sort()).toEqual([...b.itemNames.entries()].sort());
    // The IfcWall class bucket holds every geometry item.
    expect(a.classes.get('IfcWall')?.size).toBe(ITEM_COUNT);
  });
});
