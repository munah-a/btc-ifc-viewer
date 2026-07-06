/**
 * Bounded recursive flattening of an IFC item's attribute tree into `[path,
 * value]` rows (AUDIT A4/T7, W2.1). Pure; moved verbatim from viewer.ts.
 */

import {
  MAX_PROPERTY_ARRAY_PREVIEW,
  MAX_PROPERTY_DEPTH,
  MAX_PROPERTY_ROWS,
} from './types';
import { formatNumber, summarizeObject, toPropertyString, truncatePropertyValue, unwrapIfcValue } from './values';

const flattenPropertyEntries = (
  value: unknown,
  prefix: string,
  output: Array<[string, string]>,
  visited: WeakSet<object>,
  state: { truncated: boolean },
  depth = 0,
): void => {
  if (output.length >= MAX_PROPERTY_ROWS) {
    state.truncated = true;
    return;
  }

  const normalized = unwrapIfcValue(value);
  if (normalized === null || normalized === undefined) return;
  if (depth > MAX_PROPERTY_DEPTH) {
    output.push([prefix, '...']);
    return;
  }

  if (typeof normalized === 'number') {
    output.push([prefix, formatNumber(normalized, prefix)]);
    return;
  }
  if (
    typeof normalized === 'string'
    || typeof normalized === 'boolean'
    || typeof normalized === 'bigint'
  ) {
    output.push([prefix, String(normalized)]);
    return;
  }

  if (normalized instanceof Date) {
    output.push([prefix, normalized.toISOString()]);
    return;
  }

  if (Array.isArray(normalized)) {
    if (normalized.length === 0) {
      output.push([prefix, '[]']);
      return;
    }

    const scalarOnly = normalized.every((entry) => {
      const item = unwrapIfcValue(entry);
      return (
        item === null
        || item === undefined
        || typeof item === 'string'
        || typeof item === 'number'
        || typeof item === 'boolean'
        || typeof item === 'bigint'
      );
    });

    if (scalarOnly) {
      const joined = normalized
        .map((entry) => toPropertyString(entry, ''))
        .filter((entry) => entry.length > 0)
        .join(', ');
      output.push([prefix, truncatePropertyValue(joined || `${normalized.length} entries`)]);
      return;
    }

    const preview = normalized
      .slice(0, MAX_PROPERTY_ARRAY_PREVIEW)
      .map((entry) => toPropertyString(entry, ''))
      .filter((entry) => entry.length > 0)
      .join(' || ');
    const suffix = normalized.length > MAX_PROPERTY_ARRAY_PREVIEW
      ? ` (+${normalized.length - MAX_PROPERTY_ARRAY_PREVIEW} more)`
      : '';
    const summary = preview
      ? `${normalized.length} entries: ${preview}${suffix}`
      : `${normalized.length} entries${suffix}`;
    output.push([prefix, truncatePropertyValue(summary)]);
    return;
  }

  if (typeof normalized === 'object') {
    const record = normalized as Record<string, unknown>;
    if (visited.has(record)) {
      output.push([prefix, '[Circular]']);
      return;
    }
    visited.add(record);

    const entries = Object.entries(record);
    if (entries.length === 0) {
      output.push([prefix, '{}']);
      return;
    }

    if (depth >= 2) {
      output.push([prefix, summarizeObject(record)]);
      return;
    }

    for (const [key, entryValue] of entries) {
      const nextPrefix = prefix ? `${prefix}.${key}` : key;
      const unwrapped = unwrapIfcValue(entryValue);
      if (depth >= 1 && unwrapped && typeof unwrapped === 'object') {
        if (Array.isArray(unwrapped)) {
          const preview = unwrapped
            .slice(0, MAX_PROPERTY_ARRAY_PREVIEW)
            .map((entry) => toPropertyString(entry, ''))
            .filter((entry) => entry.length > 0)
            .join(' || ');
          const suffix = unwrapped.length > MAX_PROPERTY_ARRAY_PREVIEW
            ? ` (+${unwrapped.length - MAX_PROPERTY_ARRAY_PREVIEW} more)`
            : '';
          output.push([nextPrefix, truncatePropertyValue(`${unwrapped.length} entries${preview ? `: ${preview}` : ''}${suffix}`)]);
        } else {
          output.push([nextPrefix, summarizeObject(unwrapped as Record<string, unknown>)]);
        }
        if (output.length >= MAX_PROPERTY_ROWS) {
          state.truncated = true;
          return;
        }
        continue;
      }
      flattenPropertyEntries(entryValue, nextPrefix, output, visited, state, depth + 1);
      if (state.truncated) return;
    }
    return;
  }

  // `unknown` cannot be subtract-narrowed: every object/array/date shape has
  // already returned above, so only stringifiable primitives remain here.
  const primitive = normalized as string | number | boolean | bigint | symbol;
  output.push([prefix, String(primitive)]);
};

/** Flattens every top-level attribute into `[path, value]` rows (capped). */
export const collectRawPropertyEntries = (
  data: Record<string, unknown>,
): { rows: Array<[string, string]>; truncated: boolean } => {
  const rows: Array<[string, string]> = [];
  const visited = new WeakSet<object>();
  const flattenState = { truncated: false };

  for (const [key, value] of Object.entries(data)) {
    flattenPropertyEntries(value, key, rows, visited, flattenState, 0);
    if (flattenState.truncated) break;
  }

  return { rows, truncated: flattenState.truncated };
};
