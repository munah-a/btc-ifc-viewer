/**
 * Shared types + tuning constants for the pure IFC property engine (AUDIT A4/T7,
 * W2.1). Extracted verbatim from viewer.ts; DOM-free and engine-free.
 */

export type PropertySectionId =
  | 'identity'
  | 'type'
  | 'dimensions'
  | 'location'
  | 'levels'
  | 'materials'
  | 'quantities'
  | 'relations'
  | 'raw';

export interface PropertyRowData {
  key: string;
  value: string;
  searchText: string;
}

export interface PropertySectionData {
  id: PropertySectionId;
  title: string;
  defaultOpen: boolean;
  rows: PropertyRowData[];
}

export interface ExtractedPropertyFact {
  label: string;
  value: string;
  category: string;
  setName: string;
  path: string;
}

export interface GeometryProbe {
  center: { x: number; y: number; z: number };
  size: { x: number; y: number; z: number };
}

export const PROPERTY_SECTION_DEFINITIONS: Array<Omit<PropertySectionData, 'rows'>> = [
  { id: 'identity', title: 'Identity', defaultOpen: true },
  { id: 'type', title: 'Type', defaultOpen: true },
  { id: 'dimensions', title: 'Dimensions', defaultOpen: true },
  { id: 'location', title: 'Location', defaultOpen: true },
  { id: 'levels', title: 'Levels', defaultOpen: true },
  { id: 'materials', title: 'Materials', defaultOpen: true },
  { id: 'quantities', title: 'Quantities', defaultOpen: true },
  { id: 'relations', title: 'Relations', defaultOpen: false },
  { id: 'raw', title: 'Raw IFC', defaultOpen: false },
];

export const IFC_FACT_VALUE_KEYS = [
  'NominalValue',
  'LengthValue',
  'AreaValue',
  'VolumeValue',
  'CountValue',
  'WeightValue',
  'TimeValue',
  'PositiveLengthValue',
  'MassValue',
  'Width',
  'Depth',
  'Height',
  'Thickness',
  'LayerThickness',
  'Elevation',
];

export const DIMENSION_KEYWORDS = ['thickness', 'width', 'height', 'length', 'depth', 'diameter', 'radius', 'slope', 'span', 'perimeter'];
export const QUANTITY_KEYWORDS = ['area', 'volume', 'count', 'mass', 'weight', 'gross', 'net', 'perimeter', 'quantity', 'length'];
export const LOCATION_KEYWORDS = ['location', 'placement', 'coordinate', 'elevation', 'axis', 'direction', 'reference level', 'offset', 'origin', 'x', 'y', 'z'];
export const LEVEL_KEYWORDS = ['storey', 'level', 'floor', 'roof', 'sub level', 'contained in structure'];
export const MATERIAL_KEYWORDS = ['material', 'layer', 'constituent', 'finish', 'grade', 'profile'];
export const TYPE_KEYWORDS = ['type', 'family', 'assembly', 'classification', 'reference'];
export const RELATION_KEYWORDS = ['association', 'opening', 'void', 'fills', 'defines', 'typed', 'connected', 'decomposes', 'group'];

export const MAX_PROPERTY_ROWS = 280;
export const MAX_PROPERTY_DEPTH = 4;
export const MAX_PROPERTY_VALUE_LENGTH = 220;
export const MAX_PROPERTY_ARRAY_PREVIEW = 6;
