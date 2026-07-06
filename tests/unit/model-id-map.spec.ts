import { describe, expect, it } from 'vitest';
import {
  clearMap,
  cloneMap,
  countMapItems,
  intersectMaps,
  isMapEmpty,
  toSetMap,
  type ModelIdMap,
} from '../../src/core/model-id-map';

describe('toSetMap', () => {
  it('coerces arrays to Sets and keeps object identity per model', () => {
    const map = toSetMap({ a: [1, 2, 2, 3] });
    expect(map.a).toBeInstanceOf(Set);
    expect([...map.a]).toEqual([1, 2, 3]);
  });

  it('passes through existing Sets', () => {
    const set = new Set([7, 8]);
    const map = toSetMap({ a: set });
    expect(map.a).toBe(set);
  });

  it('drops models with an empty id-collection', () => {
    const map = toSetMap({ a: [], b: new Set<number>(), c: [1] });
    expect(Object.keys(map)).toEqual(['c']);
  });
});

describe('cloneMap', () => {
  it('deep-copies each id-set', () => {
    const original: ModelIdMap = { a: new Set([1, 2]) };
    const copy = cloneMap(original);
    expect(copy).not.toBe(original);
    expect(copy.a).not.toBe(original.a);
    expect([...copy.a]).toEqual([1, 2]);
    copy.a.add(99);
    expect(original.a.has(99)).toBe(false);
  });
});

describe('clearMap', () => {
  it('empties the map in place, preserving object identity', () => {
    const map: ModelIdMap = { a: new Set([1]), b: new Set([2]) };
    const ref = map;
    clearMap(map);
    expect(map).toBe(ref);
    expect(Object.keys(map)).toEqual([]);
  });
});

describe('isMapEmpty', () => {
  it('is true for {} and for maps whose sets are all empty', () => {
    expect(isMapEmpty({})).toBe(true);
    expect(isMapEmpty({ a: new Set<number>() })).toBe(true);
  });

  it('is false when any set has ids', () => {
    expect(isMapEmpty({ a: new Set<number>(), b: new Set([5]) })).toBe(false);
  });
});

describe('countMapItems', () => {
  it('sums ids across all models', () => {
    expect(countMapItems({ a: new Set([1, 2]), b: new Set([3]) })).toBe(3);
    expect(countMapItems({})).toBe(0);
  });
});

describe('intersectMaps', () => {
  it('keeps only ids present in the same model in both maps', () => {
    const a: ModelIdMap = { m1: new Set([1, 2, 3]), m2: new Set([9]) };
    const b: ModelIdMap = { m1: new Set([2, 3, 4]), m3: new Set([9]) };
    const result = intersectMaps(a, b);
    expect(Object.keys(result)).toEqual(['m1']);
    expect([...result.m1]).toEqual([2, 3]);
  });

  it('drops a model whose intersection is empty', () => {
    const result = intersectMaps({ m1: new Set([1]) }, { m1: new Set([2]) });
    expect(result).toEqual({});
  });

  it('returns {} when the maps share no model', () => {
    expect(intersectMaps({ a: new Set([1]) }, { b: new Set([1]) })).toEqual({});
  });
});
