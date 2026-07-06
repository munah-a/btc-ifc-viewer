/**
 * Value primitives for the IFC property engine (AUDIT A4/T7, W2.1): unwrapping
 * `{value,type}` wrappers, number formatting, stringification, case-insensitive
 * lookup, name discovery and storey-name discovery. All pure free functions
 * moved verbatim from viewer.ts; DOM-free and engine-free.
 */

import { MAX_PROPERTY_ARRAY_PREVIEW, MAX_PROPERTY_VALUE_LENGTH } from './types';

/** Unwraps nested web-ifc/fragments `{value}` attribute wrappers (bounded depth). */
export const unwrapIfcValue = (value: unknown): unknown => {
  let current = value;
  for (let i = 0; i < 8; i += 1) {
    if (!current || typeof current !== 'object' || Array.isArray(current)) break;
    const record = current as Record<string, unknown>;
    if (!Object.prototype.hasOwnProperty.call(record, 'value')) break;
    const next = record.value;
    if (next === undefined || next === current) break;
    current = next;
  }
  return current;
};

/** Truncates a display value with an ellipsis at `maxLength`. */
export const truncatePropertyValue = (value: string, maxLength = MAX_PROPERTY_VALUE_LENGTH): string => {
  if (value.length <= maxLength) return value;
  return `${value.slice(0, maxLength - 1)}…`;
};

/** Formats a number for display, honoring a context hint (angle/count etc.). */
export const formatNumber = (value: number, contextHint = ''): string => {
  if (Number.isNaN(value)) return 'NaN';
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) || Math.abs(value - Math.round(value)) < 1e-9) {
    return String(Math.round(value));
  }
  const hint = contextHint.toLowerCase();
  if (hint.includes('angle') || hint.includes('slope') || hint.includes('rotation') || hint.includes('tilt')) {
    return value.toFixed(1);
  }
  if (hint.includes('count') || hint.includes('number')) {
    return String(Math.round(value));
  }
  if (Math.abs(value) >= 1000) return value.toFixed(2);
  if (Math.abs(value) >= 1) return value.toFixed(3);
  return value.toFixed(4);
};

/** Compact human summary of an object by preferred identity keys. */
export const summarizeObject = (record: Record<string, unknown>): string => {
  const preferredKeys = [
    'Name',
    'LongName',
    'ObjectType',
    'PredefinedType',
    'Description',
    'type',
    '_category',
    '_localId',
    'localId',
    'GlobalId',
    '_guid',
    'Tag',
  ];
  const parts: string[] = [];

  for (const key of preferredKeys) {
    if (!Object.prototype.hasOwnProperty.call(record, key)) continue;
    const value = unwrapIfcValue(record[key]);
    if (value === null || value === undefined || typeof value === 'object') continue;
    const text = toPropertyString(value, '');
    if (!text) continue;
    parts.push(`${key}: ${text}`);
    if (parts.length >= 4) break;
  }

  if (parts.length > 0) return truncatePropertyValue(parts.join(' | '));

  const keys = Object.keys(record);
  if (keys.length === 0) return '{}';
  return `[Object: ${keys.length} properties]`;
};

/** Stringifies any IFC value into a display string, with a fallback. */
export const toPropertyString = (value: unknown, fallback = '-', contextHint = ''): string => {
  const normalized = unwrapIfcValue(value);
  if (normalized === null || normalized === undefined) return fallback;
  if (typeof normalized === 'string') return normalized.trim().length > 0 ? normalized : fallback;
  if (typeof normalized === 'number') return formatNumber(normalized, contextHint);
  if (typeof normalized === 'boolean' || typeof normalized === 'bigint') return String(normalized);
  if (normalized instanceof Date) return normalized.toISOString();
  if (Array.isArray(normalized)) {
    if (normalized.length === 0) return '[]';
    const values = normalized
      .map((entry) => toPropertyString(entry, ''))
      .filter((entry) => entry.length > 0);
    if (values.length === 0) return fallback;
    return truncatePropertyValue(values.join(', '));
  }
  if (typeof normalized === 'object') {
    return summarizeObject(normalized as Record<string, unknown>);
  }
  // `unknown` cannot be subtract-narrowed: every object/array/date shape has
  // already returned above, so only stringifiable primitives remain here.
  const primitive = normalized as string | number | boolean | bigint | symbol;
  return String(primitive);
};

/** Reads a scalar value as a trimmed string; '' for objects/arrays/nullish. */
export const readPrimitiveValue = (value: unknown): string => {
  const normalized = unwrapIfcValue(value);
  if (normalized === null || normalized === undefined) return '';
  if (typeof normalized === 'string') return normalized.trim();
  if (typeof normalized === 'number' || typeof normalized === 'boolean' || typeof normalized === 'bigint') return String(normalized);
  return '';
};

/** Case-insensitive record lookup, preferring the exact key. */
export const getRecordValueCaseInsensitive = (record: Record<string, unknown>, preferredKey: string): unknown => {
  if (Object.prototype.hasOwnProperty.call(record, preferredKey)) return record[preferredKey];
  const lowerPreferred = preferredKey.toLowerCase();
  const matched = Object.keys(record).find((key) => key.toLowerCase() === lowerPreferred);
  return matched ? record[matched] : undefined;
};

/** Recursively discovers the most name-like scalar within a value. */
export const findNameLikeValue = (value: unknown, visited: WeakSet<object>, depth: number): string => {
  if (depth > 5) return '';
  const normalized = unwrapIfcValue(value);
  if (normalized === null || normalized === undefined) return '';
  if (typeof normalized === 'string') return normalized.trim();
  if (typeof normalized === 'number' || typeof normalized === 'boolean' || typeof normalized === 'bigint') {
    return String(normalized);
  }
  if (Array.isArray(normalized)) {
    for (const entry of normalized) {
      const found = findNameLikeValue(entry, visited, depth + 1);
      if (found) return found;
    }
    return '';
  }
  if (typeof normalized !== 'object') return '';

  const record = normalized as Record<string, unknown>;
  if (visited.has(record)) return '';
  visited.add(record);

  const preferredKeys = ['Name', 'LongName', 'ObjectType', 'PredefinedType', 'Tag'];
  for (const key of preferredKeys) {
    const primitive = readPrimitiveValue(getRecordValueCaseInsensitive(record, key));
    if (primitive) return primitive;
  }

  for (const [key, entry] of Object.entries(record)) {
    const lower = key.toLowerCase();
    if (lower === 'name' || lower.endsWith('name') || lower.endsWith('objecttype')) {
      const primitive = readPrimitiveValue(entry);
      if (primitive) return primitive;
    }
  }

  for (const nested of Object.values(record)) {
    const found = findNameLikeValue(nested, visited, depth + 1);
    if (found) return found;
  }
  return '';
};

/** A display label for a spatial-tree item (Name/LongName/… then category+id). */
export const getModelTreeItemLabel = (
  data: Record<string, unknown>,
  localId: number,
  defaultCategory: string,
): string => {
  const label = readPrimitiveValue(getRecordValueCaseInsensitive(data, 'Name'))
    || readPrimitiveValue(getRecordValueCaseInsensitive(data, 'LongName'))
    || readPrimitiveValue(getRecordValueCaseInsensitive(data, 'ObjectType'))
    || readPrimitiveValue(getRecordValueCaseInsensitive(data, 'PredefinedType'))
    || findNameLikeValue(data, new WeakSet<object>(), 0);
  if (label) return label;
  const category = toPropertyString(data._category, defaultCategory) || defaultCategory;
  return `${category} ${localId}`;
};

const findStoreyNameInValue = (value: unknown, visited: WeakSet<object>, depth: number): string | null => {
  if (depth > 8) return null;
  const unwrapped = unwrapIfcValue(value);
  if (unwrapped === null || unwrapped === undefined) return null;

  if (Array.isArray(unwrapped)) {
    for (const entry of unwrapped) {
      const name = findStoreyNameInValue(entry, visited, depth + 1);
      if (name) return name;
    }
    return null;
  }

  if (typeof unwrapped !== 'object') return null;
  const record = unwrapped as Record<string, unknown>;
  if (visited.has(record)) return null;
  visited.add(record);

  const category = toPropertyString(record._category, '').toUpperCase();
  if (category.includes('IFCBUILDINGSTOREY')) {
    let storeyName = readPrimitiveValue(getRecordValueCaseInsensitive(record, 'Name'))
      || readPrimitiveValue(getRecordValueCaseInsensitive(record, 'LongName'));
    if (!storeyName) {
      for (const nested of Object.values(record)) {
        const nestedName = findNameLikeValue(nested, visited, depth + 1);
        if (nestedName) {
          storeyName = nestedName;
          break;
        }
      }
    }
    if (storeyName) return storeyName;
    return null;
  }

  for (const nested of Object.values(record)) {
    const name = findStoreyNameInValue(nested, visited, depth + 1);
    if (name) return name;
  }
  return null;
};

/** Extracts a containing-storey name from an item's relations, if any. */
export const extractStoreyNameFromItemData = (data: Record<string, unknown>): string | null => {
  const visited = new WeakSet<object>();
  const fromContained = findStoreyNameInValue(data.ContainedInStructure, visited, 0);
  if (fromContained) return fromContained;
  const fromDecomposes = findStoreyNameInValue(data.Decomposes, visited, 0);
  if (fromDecomposes) return fromDecomposes;
  return null;
};

/** Human-readable label from a dotted/underscored property path. */
export const prettifyPropertyLabel = (label: string): string => {
  const compact = label
    .replace(/\[\d+\]/g, ' ')
    .split('.')
    .filter((part) => part.length > 0)
    .slice(-1)
    .join(' ');
  const spaced = compact
    .replace(/^_+/, '')
    .replace(/[_-]+/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim();
  if (!spaced) return label;
  return spaced.replace(/\b\w/g, (char) => char.toUpperCase());
};

/** Summarizes an IFC relation value to a short name / name list. */
export const summarizeRelationValue = (value: unknown): string => {
  const normalized = unwrapIfcValue(value);
  if (normalized === null || normalized === undefined) return '';

  const summarizeOne = (entry: unknown): string => {
    const objectValue = unwrapIfcValue(entry);
    if (objectValue && typeof objectValue === 'object' && !Array.isArray(objectValue)) {
      const record = objectValue as Record<string, unknown>;
      const name = findNameLikeValue(record, new WeakSet<object>(), 0);
      return name || '';
    }
    return toPropertyString(objectValue, '');
  };

  if (Array.isArray(normalized)) {
    const entries = normalized.map((entry) => summarizeOne(entry)).filter((entry) => entry.length > 0);
    if (entries.length === 0) return '';
    if (entries.length === 1) return entries[0];
    const preview = entries.slice(0, 3).join(', ');
    const suffix = entries.length > 3 ? ` (+${entries.length - 3} more)` : '';
    return truncatePropertyValue(`${preview}${suffix}`);
  }

  return summarizeOne(normalized);
};

export { MAX_PROPERTY_ARRAY_PREVIEW };
