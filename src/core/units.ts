/**
 * Property-unit resolution (AUDIT F5). Units come from the model's own unit
 * entities (IfcSIUnit / IfcConversionBasedUnit — the members of the project's
 * IfcUnitAssignment); the label-keyword mapping only decides WHICH quantity
 * kind a property is, and the previous hardcoded metric suffixes remain the
 * fallback when a model carries no unit data.
 *
 * DOM-free and engine-free; unit-tested in tests/unit/units.spec.ts.
 */

export interface ModelUnits {
  length: string;
  area: string;
  volume: string;
  mass: string;
  angle: string;
}

/** The pre-F5 hardcoded suffixes — now only the fallback. */
export const DEFAULT_MODEL_UNITS: ModelUnits = {
  length: 'm',
  area: 'm²',
  volume: 'm³',
  mass: 'kg',
  angle: '°',
};

/** Unwraps web-ifc/fragments `{value,type}` attribute wrappers. */
const unwrapValue = (value: unknown): unknown => {
  let current = value;
  for (let i = 0; i < 4; i += 1) {
    if (!current || typeof current !== 'object' || Array.isArray(current)) break;
    const record = current as Record<string, unknown>;
    if (!Object.prototype.hasOwnProperty.call(record, 'value')) break;
    current = record.value;
  }
  return current;
};

/** Normalized enum text: unwrapped, trimmed of IFC enum dots, uppercased. */
const readEnumText = (value: unknown): string => {
  const unwrapped = unwrapValue(value);
  if (typeof unwrapped !== 'string') return '';
  return unwrapped.replace(/^\.+|\.+$/g, '').replace(/[\s_]+/g, '_').trim().toUpperCase();
};

const SI_LENGTH_PREFIX: Record<string, string> = {
  MILLI: 'mm',
  CENTI: 'cm',
  DECI: 'dm',
  '': 'm',
  KILO: 'km',
};

const lengthSymbol = (prefix: string, name: string): string | null => {
  if (name === 'METRE' || name === 'METER') return SI_LENGTH_PREFIX[prefix] ?? 'm';
  if (name === 'INCH') return 'in';
  if (name === 'FOOT' || name === 'FEET') return 'ft';
  if (name === 'YARD') return 'yd';
  if (name === 'MILE') return 'mi';
  return null;
};

const areaSymbol = (prefix: string, name: string): string | null => {
  if (name === 'SQUARE_METRE' || name === 'SQUARE_METER') {
    const base = SI_LENGTH_PREFIX[prefix];
    return base ? `${base}²` : 'm²';
  }
  if (name === 'SQUARE_FOOT' || name === 'SQUARE_FEET') return 'ft²';
  if (name === 'SQUARE_INCH') return 'in²';
  return null;
};

const volumeSymbol = (prefix: string, name: string): string | null => {
  if (name === 'CUBIC_METRE' || name === 'CUBIC_METER') {
    const base = SI_LENGTH_PREFIX[prefix];
    return base ? `${base}³` : 'm³';
  }
  if (name === 'CUBIC_FOOT' || name === 'CUBIC_FEET') return 'ft³';
  if (name === 'CUBIC_INCH') return 'in³';
  return null;
};

const massSymbol = (prefix: string, name: string): string | null => {
  if (name === 'GRAM') {
    if (prefix === 'KILO') return 'kg';
    if (prefix === 'MILLI') return 'mg';
    if (prefix === '') return 'g';
    return 'g';
  }
  if (name === 'POUND') return 'lb';
  if (name === 'TONNE') return 't';
  return null;
};

const angleSymbol = (name: string): string | null => {
  if (name === 'RADIAN') return 'rad';
  if (name === 'DEGREE' || name === 'DEGREES') return '°';
  return null;
};

/**
 * Resolves a model's display units from its unit-entity rows (the result of
 * getItemsData over the IFCSIUNIT/IFCCONVERSIONBASEDUNIT categories). Fields
 * that cannot be resolved keep the metric defaults.
 */
export const resolveModelUnits = (rows: unknown[]): ModelUnits => {
  const units: ModelUnits = { ...DEFAULT_MODEL_UNITS };
  for (const row of rows) {
    if (!row || typeof row !== 'object' || Array.isArray(row)) continue;
    const record = row as Record<string, unknown>;
    const unitType = readEnumText(record.UnitType);
    const name = readEnumText(record.Name);
    const prefix = readEnumText(record.Prefix);
    if (!unitType || !name) continue;

    switch (unitType) {
      case 'LENGTHUNIT': {
        const symbol = lengthSymbol(prefix, name);
        if (symbol) units.length = symbol;
        break;
      }
      case 'AREAUNIT': {
        const symbol = areaSymbol(prefix, name);
        if (symbol) units.area = symbol;
        break;
      }
      case 'VOLUMEUNIT': {
        const symbol = volumeSymbol(prefix, name);
        if (symbol) units.volume = symbol;
        break;
      }
      case 'MASSUNIT': {
        const symbol = massSymbol(prefix, name);
        if (symbol) units.mass = symbol;
        break;
      }
      case 'PLANEANGLEUNIT': {
        const symbol = angleSymbol(name);
        if (symbol) units.angle = symbol;
        break;
      }
      default:
        break;
    }
  }
  return units;
};

/**
 * Suffix for a property label: keyword classification picks the quantity
 * kind; the actual unit text comes from the model's resolved units (F5).
 * Returns '' when the label names no known quantity kind.
 */
export const unitSuffixForLabel = (label: string, units: ModelUnits = DEFAULT_MODEL_UNITS): string => {
  const lower = label.toLowerCase();
  if (lower.includes('area')) return ` ${units.area}`;
  if (lower.includes('volume')) return ` ${units.volume}`;
  if (lower.includes('length') || lower.includes('width') || lower.includes('height')
    || lower.includes('thickness') || lower.includes('depth') || lower.includes('radius')
    || lower.includes('diameter') || lower.includes('perimeter') || lower.includes('span')) {
    return ` ${units.length}`;
  }
  if (lower.includes('mass') || lower.includes('weight')) return ` ${units.mass}`;
  if (lower.includes('angle') || lower.includes('slope') || lower.includes('tilt')) {
    // The degree sign reads glued to the number; word units get a space.
    return units.angle === '°' ? '°' : ` ${units.angle}`;
  }
  return '';
};
