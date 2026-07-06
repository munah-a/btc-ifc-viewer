/**
 * Pure set-algebra over the engine's selection/visibility map type (AUDIT A4/T7,
 * W2.1). A `ModelIdMap` is `Record<modelId, Set<localId>>` — structurally
 * identical to `@thatopen/components`' `ModelIdMap`, declared here without the
 * engine import so this module stays DOM-free and engine-free (unit-tested in
 * tests/unit/model-id-map.spec.ts). Moved verbatim from viewer.ts:224-268.
 */

export type ModelIdMap = Record<string, Set<number>>;

/**
 * Builds a `ModelIdMap` from a plain record of arrays or sets, dropping any
 * model whose id-collection is empty.
 */
export const toSetMap = (plain: Record<string, number[] | Set<number>>): ModelIdMap => {
  const result: ModelIdMap = {};
  for (const [modelId, ids] of Object.entries(plain)) {
    const set = ids instanceof Set ? ids : new Set(ids);
    if (set.size > 0) result[modelId] = set;
  }
  return result;
};

/** Deep-copies a map (each id-set is cloned). */
export const cloneMap = (map: ModelIdMap): ModelIdMap => {
  const copy: ModelIdMap = {};
  for (const [modelId, ids] of Object.entries(map)) copy[modelId] = new Set(ids);
  return copy;
};

/** Empties a map in place (preserves the object identity the engine holds). */
export const clearMap = (map: ModelIdMap): void => {
  for (const key of Object.keys(map)) delete map[key];
};

/** True when no model in the map has any id. */
export const isMapEmpty = (map: ModelIdMap): boolean => {
  for (const ids of Object.values(map)) {
    if (ids.size > 0) return false;
  }
  return true;
};

/** Total number of ids across all models. */
export const countMapItems = (map: ModelIdMap): number => {
  let count = 0;
  for (const ids of Object.values(map)) count += ids.size;
  return count;
};

/** Per-model intersection of two maps; models/ids present in only one are dropped. */
export const intersectMaps = (a: ModelIdMap, b: ModelIdMap): ModelIdMap => {
  const result: ModelIdMap = {};
  for (const [modelId, idsA] of Object.entries(a)) {
    const idsB = b[modelId];
    if (!idsB) continue;
    const intersection = new Set<number>();
    for (const id of idsA) {
      if (idsB.has(id)) intersection.add(id);
    }
    if (intersection.size > 0) result[modelId] = intersection;
  }
  return result;
};
