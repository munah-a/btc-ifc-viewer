import { describe, expect, it } from 'vitest';
import {
  buildPropertySections,
  classifyPropertyFact,
  collectPropertyFacts,
  collectRawPropertyEntries,
  extractStoreyNameFromItemData,
  findNameLikeValue,
  formatNumber,
  getModelTreeItemLabel,
  MAX_PROPERTY_ROWS,
  prettifyPropertyLabel,
  readPrimitiveValue,
  summarizeRelationValue,
  toPropertyString,
  unwrapIfcValue,
  type ExtractedPropertyFact,
} from '../../src/core/property-engine';
import { DEFAULT_MODEL_UNITS, type ModelUnits } from '../../src/core/units';

const wrap = (value: unknown) => ({ value, type: 1 });

describe('unwrapIfcValue', () => {
  it('unwraps nested {value} wrappers', () => {
    expect(unwrapIfcValue(wrap(wrap('Wall')))).toBe('Wall');
  });

  it('leaves plain values and arrays untouched', () => {
    expect(unwrapIfcValue(42)).toBe(42);
    const arr = [1, 2];
    expect(unwrapIfcValue(arr)).toBe(arr);
  });

  it('stops on a self-referential wrapper (no infinite loop)', () => {
    const cyclic: Record<string, unknown> = {};
    cyclic.value = cyclic;
    expect(unwrapIfcValue(cyclic)).toBe(cyclic);
  });
});

describe('readPrimitiveValue', () => {
  it('reads scalars through wrappers, trims strings', () => {
    expect(readPrimitiveValue(wrap('  Level 1  '))).toBe('Level 1');
    expect(readPrimitiveValue(wrap(3))).toBe('3');
  });

  it('returns empty string for objects/arrays/nullish', () => {
    expect(readPrimitiveValue({ a: 1 })).toBe('');
    expect(readPrimitiveValue(null)).toBe('');
  });
});

describe('formatNumber', () => {
  it('rounds near-integers to integers', () => {
    expect(formatNumber(5.0000000001)).toBe('5');
    expect(formatNumber(7)).toBe('7');
  });

  it('applies context-sensitive precision', () => {
    expect(formatNumber(1.23456, 'angle')).toBe('1.2');
    expect(formatNumber(1234.5678)).toBe('1234.57');
    expect(formatNumber(1.23456)).toBe('1.235');
    expect(formatNumber(0.00123)).toBe('0.0012');
  });
});

describe('toPropertyString', () => {
  it('falls back for empty/nullish', () => {
    expect(toPropertyString(undefined, '—')).toBe('—');
    expect(toPropertyString(wrap('  '), '—')).toBe('—');
  });

  it('joins scalar arrays and summarizes objects', () => {
    expect(toPropertyString([wrap('a'), wrap('b')])).toBe('a, b');
    expect(toPropertyString({ Name: wrap('Beam'), _category: wrap('IFCBEAM') })).toContain('Name: Beam');
  });
});

describe('findNameLikeValue', () => {
  it('prefers Name over deeper values and recurses', () => {
    expect(findNameLikeValue({ Foo: 1, Name: wrap('Door 3') }, new WeakSet(), 0)).toBe('Door 3');
    expect(findNameLikeValue({ nested: { LongName: wrap('Storey A') } }, new WeakSet(), 0)).toBe('Storey A');
  });
});

describe('getModelTreeItemLabel', () => {
  it('uses the name when present, else category + id', () => {
    expect(getModelTreeItemLabel({ Name: wrap('Wall-1') }, 7, 'IFCWALL')).toBe('Wall-1');
    // With no name-like attribute at all, it composes the default category + id.
    expect(getModelTreeItemLabel({}, 9, 'Item')).toBe('Item 9');
  });
});

describe('extractStoreyNameFromItemData', () => {
  it('finds a storey through ContainedInStructure relations', () => {
    const data = {
      ContainedInStructure: [
        { _category: wrap('IFCRELCONTAINEDINSPATIALSTRUCTURE'), RelatingStructure: { _category: wrap('IFCBUILDINGSTOREY'), Name: wrap('Ground Floor') } },
      ],
    };
    expect(extractStoreyNameFromItemData(data)).toBe('Ground Floor');
  });

  it('returns null with no storey relation', () => {
    expect(extractStoreyNameFromItemData({ Name: wrap('x') })).toBeNull();
  });
});

describe('prettifyPropertyLabel', () => {
  it('splits camelCase and underscores, title-cases, keeps last path segment', () => {
    expect(prettifyPropertyLabel('Pset_WallCommon.LoadBearing')).toBe('Load Bearing');
    expect(prettifyPropertyLabel('_localId')).toBe('Local Id');
  });
});

describe('summarizeRelationValue', () => {
  it('summarizes one, and previews many with a +N more suffix', () => {
    expect(summarizeRelationValue({ Name: wrap('Type A') })).toBe('Type A');
    const many = Array.from({ length: 5 }, (_, i) => ({ Name: wrap(`R${i}`) }));
    expect(summarizeRelationValue(many)).toBe('R0, R1, R2 (+2 more)');
  });
});

describe('collectRawPropertyEntries', () => {
  it('flattens nested attributes into dotted paths', () => {
    const { rows, truncated } = collectRawPropertyEntries({
      Name: wrap('Wall'),
      Pset: { FireRating: wrap('2h') },
    });
    const map = new Map(rows);
    expect(map.get('Name')).toBe('Wall');
    expect(map.get('Pset.FireRating')).toBe('2h');
    expect(truncated).toBe(false);
  });

  it('caps at MAX_PROPERTY_ROWS and flags truncation', () => {
    const big: Record<string, unknown> = {};
    for (let i = 0; i < MAX_PROPERTY_ROWS + 50; i += 1) big[`k${i}`] = wrap(i);
    const { rows, truncated } = collectRawPropertyEntries(big);
    expect(truncated).toBe(true);
    expect(rows.length).toBeLessThanOrEqual(MAX_PROPERTY_ROWS);
  });
});

describe('collectPropertyFacts + classifyPropertyFact', () => {
  it('extracts a single-value property fact and applies injected units', () => {
    const imperial: ModelUnits = { ...DEFAULT_MODEL_UNITS, area: 'ft²' };
    const data = {
      IsDefinedBy: [
        {
          _category: wrap('IFCPROPERTYSET'),
          Name: wrap('Pset_WallCommon'),
          HasProperties: [
            { _category: wrap('IFCPROPERTYSINGLEVALUE'), Name: wrap('FireRating'), NominalValue: wrap('2h') },
          ],
        },
        {
          _category: wrap('IFCELEMENTQUANTITY'),
          Name: wrap('Qto_WallBaseQuantities'),
          Quantities: [
            { _category: wrap('IFCQUANTITYAREA'), Name: wrap('NetSideArea'), AreaValue: wrap(12) },
          ],
        },
      ],
    };
    const facts: ExtractedPropertyFact[] = [];
    collectPropertyFacts(data, facts, new WeakSet(), imperial);
    const fire = facts.find((f) => f.label === 'FireRating');
    const area = facts.find((f) => f.label === 'NetSideArea');
    expect(fire?.value).toBe('2h');
    expect(area?.value).toBe('12 ft²');
  });

  it('classifies facts into sections by keyword and category', () => {
    const mk = (over: Partial<ExtractedPropertyFact>): ExtractedPropertyFact =>
      ({ label: '', value: '', category: '', setName: '', path: '', ...over });
    expect(classifyPropertyFact(mk({ label: 'Material' }))).toBe('materials');
    expect(classifyPropertyFact(mk({ label: 'Thickness' }))).toBe('dimensions');
    expect(classifyPropertyFact(mk({ label: 'NetVolume' }))).toBe('quantities');
    expect(classifyPropertyFact(mk({ label: 'Storey' }))).toBe('levels');
    expect(classifyPropertyFact(mk({ label: 'RandomThing' }))).toBe('identity');
  });
});

describe('buildPropertySections', () => {
  const baseData = {
    Name: wrap('Wall 101'),
    GlobalId: wrap('1a2b3c'),
    _category: wrap('IFCWALL'),
    ObjectType: wrap('Basic Wall'),
  };

  it('produces identity/type sections and honors the index storey', () => {
    const sections = buildPropertySections(baseData, 101, DEFAULT_MODEL_UNITS, 'Level 2', '', null);
    const byId = new Map(sections.map((s) => [s.id, s]));
    expect(byId.get('identity')?.rows.find((r) => r.key === 'Name')?.value).toBe('Wall 101');
    expect(byId.get('levels')?.rows.find((r) => r.key === 'Storey')?.value).toBe('Level 2');
    // sections come back in definition order, only non-empty ones
    expect(sections.every((s) => s.rows.length > 0)).toBe(true);
  });

  it('adds geometry-derived rows when a probe is supplied', () => {
    const probe = { center: { x: 1, y: 2, z: 3 }, size: { x: 4, y: 5, z: 6 } };
    const sections = buildPropertySections(baseData, 101, DEFAULT_MODEL_UNITS, undefined, '', probe);
    const dims = sections.find((s) => s.id === 'dimensions');
    expect(dims?.rows.some((r) => r.key === 'Extent X' && r.value === '4.000 m')).toBe(true);
  });

  it('falls back to the storey found inside the item data', () => {
    const data = {
      ...baseData,
      ContainedInStructure: [
        { _category: wrap('IFCRELCONTAINEDINSPATIALSTRUCTURE'), RelatingStructure: { _category: wrap('IFCBUILDINGSTOREY'), Name: wrap('Roof') } },
      ],
    };
    const sections = buildPropertySections(data, 101, DEFAULT_MODEL_UNITS, undefined, '', null);
    const levels = sections.find((s) => s.id === 'levels');
    expect(levels?.rows.find((r) => r.key === 'Storey')?.value).toBe('Roof');
  });
});
