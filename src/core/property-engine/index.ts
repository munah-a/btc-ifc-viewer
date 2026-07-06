/**
 * Pure IFC property engine (AUDIT A4/T7, W2.1). The biggest pure-logic win of
 * the W2 extraction: value unwrapping/formatting, attribute flattening, fact
 * extraction, section classification and section assembly — all DOM-free and
 * engine-free, unit-tested in tests/unit/property-engine.spec.ts. viewer.ts
 * keeps only the DOM rendering (renderPropertySections / applyPropertiesFilter)
 * and the engine calls that fetch the raw item data.
 */

export * from './types';
export {
  extractStoreyNameFromItemData,
  findNameLikeValue,
  formatNumber,
  getModelTreeItemLabel,
  getRecordValueCaseInsensitive,
  prettifyPropertyLabel,
  readPrimitiveValue,
  summarizeObject,
  summarizeRelationValue,
  toPropertyString,
  truncatePropertyValue,
  unwrapIfcValue,
} from './values';
export { collectRawPropertyEntries } from './flatten';
export { classifyPropertyFact, collectPropertyFacts } from './facts';
export { buildPropertySections } from './sections';
