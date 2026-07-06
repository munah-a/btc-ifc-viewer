/**
 * Assembles the grouped property sections shown in the Properties panel (AUDIT
 * A4/T7, W2.1). Pure orchestration over the value/flatten/facts helpers; moved
 * verbatim from viewer.ts. Two pieces of instance/engine state that used to be
 * read from `this` are now parameters: the per-model display `units` (F5) and
 * the `indexStorey` string resolved from the model index (falls back here to
 * the storey discovered in the item's own relations).
 */

import { unitSuffixForLabel, type ModelUnits, DEFAULT_MODEL_UNITS } from '../units';
import { collectRawPropertyEntries } from './flatten';
import { classifyPropertyFact, collectPropertyFacts } from './facts';
import {
  DIMENSION_KEYWORDS,
  LEVEL_KEYWORDS,
  LOCATION_KEYWORDS,
  MATERIAL_KEYWORDS,
  MAX_PROPERTY_ROWS,
  PROPERTY_SECTION_DEFINITIONS,
  RELATION_KEYWORDS,
  type ExtractedPropertyFact,
  type GeometryProbe,
  type PropertySectionData,
  type PropertySectionId,
} from './types';
import {
  extractStoreyNameFromItemData,
  prettifyPropertyLabel,
  summarizeRelationValue,
  toPropertyString,
} from './values';

const createPropertySections = (): Record<PropertySectionId, PropertySectionData> => {
  const sections = {} as Record<PropertySectionId, PropertySectionData>;
  for (const definition of PROPERTY_SECTION_DEFINITIONS) {
    sections[definition.id] = { ...definition, rows: [] };
  }
  return sections;
};

const addPropertySectionRow = (
  sections: Record<PropertySectionId, PropertySectionData>,
  sectionId: PropertySectionId,
  key: string,
  value: string,
  extraSearch = '',
): void => {
  const normalizedKey = prettifyPropertyLabel(key);
  const normalizedValue = value.trim();
  if (!normalizedKey || !normalizedValue || normalizedValue === '-' || normalizedValue === '{}' || normalizedValue === '[]') return;

  const section = sections[sectionId];
  const duplicate = section.rows.some((row) => (
    row.key.toLowerCase() === normalizedKey.toLowerCase()
    && row.value.toLowerCase() === normalizedValue.toLowerCase()
  ));
  if (duplicate) return;

  section.rows.push({
    key: normalizedKey,
    value: normalizedValue,
    searchText: `${section.title} ${normalizedKey} ${normalizedValue} ${extraSearch}`.toLowerCase(),
  });
};

const addKeywordMatchedRows = (
  sections: Record<PropertySectionId, PropertySectionData>,
  sectionId: PropertySectionId,
  rows: Array<[string, string]>,
  keywords: string[],
  units: ModelUnits,
  pathPrefix = '',
): void => {
  for (const [path, value] of rows) {
    const haystack = `${pathPrefix} ${path}`.toLowerCase();
    if (!keywords.some((keyword) => haystack.includes(keyword))) continue;
    const unit = unitSuffixForLabel(path, units);
    const displayValue = (unit && /^-?\d+(\.\d+)?$/.test(value.trim())) ? value + unit : value;
    addPropertySectionRow(sections, sectionId, path, displayValue, path);
  }
};

/**
 * Builds the ordered, non-empty property sections for a selected item.
 *
 * @param data - the item's attribute+relation data from `getItemsData`.
 * @param localId - the selected element's local id (for default labels).
 * @param units - the selected model's display units (F5).
 * @param indexStorey - storey from the model index, if known; falls back to the
 *   storey discovered inside `data`.
 * @param modelVolume - pre-formatted volume string (from `getItemsVolume`), or ''.
 * @param geometryProbe - bbox center/size from the engine, or null.
 */
export const buildPropertySections = (
  data: Record<string, unknown>,
  localId: number,
  units: ModelUnits,
  indexStorey: string | undefined,
  modelVolume: string,
  geometryProbe: GeometryProbe | null,
): PropertySectionData[] => {
  const sections = createPropertySections();
  const itemName = toPropertyString(data.Name, '') || `Element ${localId}`;
  const typeValue = [
    toPropertyString(data.ObjectType, ''),
    toPropertyString(data.PredefinedType, ''),
    toPropertyString(data.type, ''),
    toPropertyString(data._category, ''),
  ].find((entry) => entry.length > 0) || '-';

  addPropertySectionRow(sections, 'identity', 'Name', itemName);
  addPropertySectionRow(sections, 'identity', 'Global Id', toPropertyString(data.GlobalId, '-'));
  addPropertySectionRow(sections, 'identity', 'Description', toPropertyString(data.Description, '-'));
  addPropertySectionRow(sections, 'identity', 'Tag', toPropertyString(data.Tag, ''));
  addPropertySectionRow(sections, 'identity', 'Category', toPropertyString(data._category, ''));

  addPropertySectionRow(sections, 'type', 'Type', typeValue);
  addPropertySectionRow(sections, 'type', 'Object Type', toPropertyString(data.ObjectType, ''));
  addPropertySectionRow(sections, 'type', 'Predefined Type', toPropertyString(data.PredefinedType, ''));
  addPropertySectionRow(sections, 'type', 'Type Relation', summarizeRelationValue(data.IsTypedBy));

  const storey = indexStorey || extractStoreyNameFromItemData(data) || '-';
  addPropertySectionRow(sections, 'levels', 'Storey', storey);
  addPropertySectionRow(sections, 'levels', 'Contained In Structure', summarizeRelationValue(data.ContainedInStructure));
  addPropertySectionRow(sections, 'levels', 'Decomposes', summarizeRelationValue(data.Decomposes));

  addPropertySectionRow(sections, 'location', 'Object Placement', summarizeRelationValue(data.ObjectPlacement));
  if (geometryProbe) {
    addPropertySectionRow(sections, 'location', 'Center X', `${geometryProbe.center.x.toFixed(3)} m`);
    addPropertySectionRow(sections, 'location', 'Center Y', `${geometryProbe.center.y.toFixed(3)} m`);
    addPropertySectionRow(sections, 'location', 'Center Z', `${geometryProbe.center.z.toFixed(3)} m`);
    addPropertySectionRow(sections, 'dimensions', 'Extent X', `${geometryProbe.size.x.toFixed(3)} m`);
    addPropertySectionRow(sections, 'dimensions', 'Extent Y', `${geometryProbe.size.y.toFixed(3)} m`);
    addPropertySectionRow(sections, 'dimensions', 'Extent Z', `${geometryProbe.size.z.toFixed(3)} m`);

    const slabLikeText = `${itemName} ${typeValue} ${toPropertyString(data._category, '')}`.toLowerCase();
    if (/(slab|floor|deck|roof|ceiling|covering)/.test(slabLikeText)) {
      const inferredThickness = Math.min(geometryProbe.size.x, geometryProbe.size.y, geometryProbe.size.z);
      addPropertySectionRow(sections, 'dimensions', 'Thickness (Approx)', `${inferredThickness.toFixed(3)} m`, 'geometry thickness');
    }
  }

  addPropertySectionRow(sections, 'relations', 'Property Definitions', summarizeRelationValue(data.IsDefinedBy));
  addPropertySectionRow(sections, 'relations', 'Associations', summarizeRelationValue(data.HasAssociations));
  addPropertySectionRow(sections, 'relations', 'Openings', summarizeRelationValue(data.HasOpenings));
  addPropertySectionRow(sections, 'relations', 'Voids Elements', summarizeRelationValue(data.VoidsElements));
  addPropertySectionRow(sections, 'relations', 'Fills Voids', summarizeRelationValue(data.FillsVoids));

  if (modelVolume) addPropertySectionRow(sections, 'quantities', 'Volume', modelVolume);

  const factRows: ExtractedPropertyFact[] = [];
  collectPropertyFacts(data, factRows, new WeakSet<object>(), units);
  for (const fact of factRows) {
    const sectionId = classifyPropertyFact(fact);
    addPropertySectionRow(sections, sectionId, fact.label, fact.value, `${fact.setName} ${fact.category} ${fact.path}`);
  }

  const { rows: rawRows, truncated } = collectRawPropertyEntries(data);
  addKeywordMatchedRows(sections, 'dimensions', rawRows, DIMENSION_KEYWORDS, units);
  addKeywordMatchedRows(sections, 'location', rawRows, LOCATION_KEYWORDS, units);
  addKeywordMatchedRows(sections, 'levels', rawRows, LEVEL_KEYWORDS, units);
  addKeywordMatchedRows(sections, 'materials', rawRows, MATERIAL_KEYWORDS, units);
  addKeywordMatchedRows(sections, 'relations', rawRows, RELATION_KEYWORDS, units);

  for (const [key, value] of rawRows) {
    const unit = unitSuffixForLabel(key, units);
    const displayValue = (unit && /^-?\d+(\.\d+)?$/.test(value.trim())) ? value + unit : value;
    addPropertySectionRow(sections, 'raw', key, displayValue, key);
  }
  if (truncated) addPropertySectionRow(sections, 'raw', 'Info', `Properties truncated to ${MAX_PROPERTY_ROWS} rows for readability`);

  for (const section of Object.values(sections)) {
    section.rows.sort((a, b) => a.key.localeCompare(b.key, undefined, { numeric: true }));
  }

  return PROPERTY_SECTION_DEFINITIONS
    .map((definition) => sections[definition.id])
    .filter((section) => section.rows.length > 0);
};

export { DEFAULT_MODEL_UNITS };
