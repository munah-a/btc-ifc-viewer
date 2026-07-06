/**
 * IFC "fact" extraction + section classification (AUDIT A4/T7, W2.1): walks an
 * item's attribute tree pulling property-set / quantity / material values into
 * flat facts, and classifies each fact into a display section by keyword. Pure;
 * moved verbatim from viewer.ts. Units are injected (F5) rather than read from
 * instance state.
 */

import { DEFAULT_MODEL_UNITS, unitSuffixForLabel, type ModelUnits } from '../units';
import {
  DIMENSION_KEYWORDS,
  IFC_FACT_VALUE_KEYS,
  LEVEL_KEYWORDS,
  LOCATION_KEYWORDS,
  MATERIAL_KEYWORDS,
  MAX_PROPERTY_DEPTH,
  QUANTITY_KEYWORDS,
  RELATION_KEYWORDS,
  TYPE_KEYWORDS,
  type ExtractedPropertyFact,
  type PropertySectionId,
} from './types';
import {
  findNameLikeValue,
  getRecordValueCaseInsensitive,
  prettifyPropertyLabel,
  readPrimitiveValue,
  toPropertyString,
  unwrapIfcValue,
} from './values';

const extractFactsFromRecord = (
  record: Record<string, unknown>,
  category: string,
  setName: string,
  path: string,
  units: ModelUnits,
): ExtractedPropertyFact[] => {
  const facts: ExtractedPropertyFact[] = [];
  const categoryUpper = category.toUpperCase();
  const name = readPrimitiveValue(getRecordValueCaseInsensitive(record, 'Name'))
    || readPrimitiveValue(getRecordValueCaseInsensitive(record, 'LongName'));

  const pushFact = (label: string, value: string): void => {
    const trimmedLabel = label.trim();
    const trimmedValue = value.trim();
    if (!trimmedLabel || !trimmedValue || trimmedValue === '-' || trimmedValue === '{}' || trimmedValue === '[]') return;
    facts.push({ label: trimmedLabel, value: trimmedValue, category, setName, path });
  };

  if (categoryUpper.includes('IFCPROPERTYSINGLEVALUE')) {
    pushFact(name, toPropertyString(getRecordValueCaseInsensitive(record, 'NominalValue'), ''));
    return facts;
  }

  if (categoryUpper.includes('IFCPROPERTYLISTVALUE')) {
    pushFact(name, toPropertyString(getRecordValueCaseInsensitive(record, 'ListValues'), ''));
    return facts;
  }

  if (categoryUpper.includes('IFCPROPERTYENUMERATEDVALUE')) {
    pushFact(name, toPropertyString(getRecordValueCaseInsensitive(record, 'EnumerationValues'), ''));
    return facts;
  }

  if (categoryUpper.includes('IFCPROPERTYBOUNDEDVALUE')) {
    const lower = toPropertyString(getRecordValueCaseInsensitive(record, 'LowerBoundValue'), '');
    const upper = toPropertyString(getRecordValueCaseInsensitive(record, 'UpperBoundValue'), '');
    const boundedValue = [lower && `Lower: ${lower}`, upper && `Upper: ${upper}`].filter(Boolean).join(' | ');
    pushFact(name, boundedValue);
    return facts;
  }

  if (categoryUpper.includes('IFCQUANTITY')) {
    const valueKey = IFC_FACT_VALUE_KEYS.find((key) => toPropertyString(getRecordValueCaseInsensitive(record, key), '').length > 0);
    if (valueKey) {
      const rawValue = toPropertyString(getRecordValueCaseInsensitive(record, valueKey), '', valueKey);
      const unit = unitSuffixForLabel(valueKey, units);
      pushFact(name || prettifyPropertyLabel(valueKey), rawValue + unit);
    }
    return facts;
  }

  if (categoryUpper.includes('IFCMATERIALLAYER')) {
    const materialName = findNameLikeValue(getRecordValueCaseInsensitive(record, 'Material'), new WeakSet<object>(), 0)
      || name;
    if (materialName) pushFact('Material', materialName);
    pushFact('Layer Thickness', toPropertyString(getRecordValueCaseInsensitive(record, 'LayerThickness'), ''));
    return facts;
  }

  if (categoryUpper === 'IFCMATERIAL') {
    if (name) pushFact('Material', name);
    return facts;
  }

  if (categoryUpper.includes('IFCMATERIALPROFILE')) {
    if (name) pushFact('Material Profile', name);
    const materialName = findNameLikeValue(getRecordValueCaseInsensitive(record, 'Material'), new WeakSet<object>(), 0);
    if (materialName) pushFact('Material', materialName);
    return facts;
  }

  if (categoryUpper.includes('IFCMATERIALCONSTITUENT')) {
    if (name) pushFact('Material Constituent', name);
    const materialName = findNameLikeValue(getRecordValueCaseInsensitive(record, 'Material'), new WeakSet<object>(), 0);
    if (materialName) pushFact('Material', materialName);
    return facts;
  }

  return facts;
};

/** Recursively harvests property-set / quantity / material facts from an item. */
export const collectPropertyFacts = (
  value: unknown,
  output: ExtractedPropertyFact[],
  visited: WeakSet<object>,
  units: ModelUnits = DEFAULT_MODEL_UNITS,
  path = '',
  setName = '',
  depth = 0,
): void => {
  if (depth > MAX_PROPERTY_DEPTH + 6) return;

  const normalized = unwrapIfcValue(value);
  if (normalized === null || normalized === undefined) return;

  if (Array.isArray(normalized)) {
    normalized.forEach((entry, index) => {
      collectPropertyFacts(entry, output, visited, units, `${path}[${index}]`, setName, depth + 1);
    });
    return;
  }

  if (typeof normalized !== 'object') return;
  const record = normalized as Record<string, unknown>;
  if (visited.has(record)) return;
  visited.add(record);

  const category = toPropertyString(record._category, '');
  const categoryUpper = category.toUpperCase();
  let nextSetName = setName;
  if (
    categoryUpper.includes('IFCPROPERTYSET')
    || categoryUpper.includes('IFCELEMENTQUANTITY')
    || categoryUpper.includes('IFCMATERIALPROPERTIES')
    || categoryUpper.includes('IFCPROFILEPROPERTIES')
    || categoryUpper.includes('IFCMATERIALPROFILESET')
    || categoryUpper.includes('IFCMATERIALLAYERSET')
  ) {
    nextSetName = readPrimitiveValue(getRecordValueCaseInsensitive(record, 'Name')) || setName;
  }

  output.push(...extractFactsFromRecord(record, category, nextSetName, path, units));

  for (const [key, entry] of Object.entries(record)) {
    const nextPath = path ? `${path}.${key}` : key;
    collectPropertyFacts(entry, output, visited, units, nextPath, nextSetName, depth + 1);
  }
};

/** Assigns a fact to a display section by keyword / category heuristics. */
export const classifyPropertyFact = (fact: ExtractedPropertyFact): PropertySectionId => {
  const haystack = `${fact.setName} ${fact.label} ${fact.category} ${fact.path}`.toLowerCase();
  if (MATERIAL_KEYWORDS.some((keyword) => haystack.includes(keyword)) || fact.category.toUpperCase().includes('IFCMATERIAL')) {
    return 'materials';
  }
  if (DIMENSION_KEYWORDS.some((keyword) => haystack.includes(keyword))) return 'dimensions';
  if (QUANTITY_KEYWORDS.some((keyword) => haystack.includes(keyword)) || fact.category.toUpperCase().includes('QUANTITY')) {
    return 'quantities';
  }
  if (LEVEL_KEYWORDS.some((keyword) => haystack.includes(keyword))) return 'levels';
  if (LOCATION_KEYWORDS.some((keyword) => haystack.includes(keyword))) return 'location';
  if (TYPE_KEYWORDS.some((keyword) => haystack.includes(keyword))) return 'type';
  if (RELATION_KEYWORDS.some((keyword) => haystack.includes(keyword))) return 'relations';
  return 'identity';
};
