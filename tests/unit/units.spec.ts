import { describe, expect, it } from 'vitest';

import { DEFAULT_MODEL_UNITS, resolveModelUnits, unitSuffixForLabel } from '../../src/core/units';

// Rows as fragments' getItemsData returns them: {value,type} ItemAttribute
// wrappers around the raw web-ifc enum strings.
const wrap = (value: string): { value: string; type: number } => ({ value, type: 3 });

describe('resolveModelUnits (AUDIT F5 — W1.9)', () => {
  it('resolves a millimetre SI model (the common Revit export)', () => {
    const units = resolveModelUnits([
      { UnitType: wrap('LENGTHUNIT'), Prefix: wrap('MILLI'), Name: wrap('METRE') },
      { UnitType: wrap('AREAUNIT'), Name: wrap('SQUARE_METRE') },
      { UnitType: wrap('VOLUMEUNIT'), Name: wrap('CUBIC_METRE') },
      { UnitType: wrap('MASSUNIT'), Prefix: wrap('KILO'), Name: wrap('GRAM') },
      { UnitType: wrap('PLANEANGLEUNIT'), Name: wrap('RADIAN') },
    ]);
    expect(units).toEqual({ length: 'mm', area: 'm²', volume: 'm³', mass: 'kg', angle: 'rad' });
  });

  it('resolves imperial conversion-based units', () => {
    const units = resolveModelUnits([
      { UnitType: wrap('LENGTHUNIT'), Name: wrap('FOOT') },
      { UnitType: wrap('AREAUNIT'), Name: wrap('SQUARE FOOT') },
      { UnitType: wrap('VOLUMEUNIT'), Name: wrap('CUBIC FOOT') },
      { UnitType: wrap('PLANEANGLEUNIT'), Name: wrap('DEGREE') },
    ]);
    expect(units.length).toBe('ft');
    expect(units.area).toBe('ft²');
    expect(units.volume).toBe('ft³');
    expect(units.angle).toBe('°');
    expect(units.mass).toBe(DEFAULT_MODEL_UNITS.mass); // untouched → fallback
  });

  it('handles raw strings, dotted IFC enums, and unknown entries', () => {
    const units = resolveModelUnits([
      { UnitType: '.LENGTHUNIT.', Prefix: '.CENTI.', Name: '.METRE.' },
      { UnitType: wrap('THERMODYNAMICTEMPERATUREUNIT'), Name: wrap('KELVIN') }, // ignored kind
      { UnitType: wrap('LENGTHUNIT') }, // missing name — ignored
      'garbage',
      null,
    ]);
    expect(units.length).toBe('cm');
  });

  it('falls back to the metric defaults for models without unit data', () => {
    expect(resolveModelUnits([])).toEqual(DEFAULT_MODEL_UNITS);
  });
});

describe('unitSuffixForLabel (F5 keyword inference, demoted to kind-picker)', () => {
  const mmUnits = { ...DEFAULT_MODEL_UNITS, length: 'mm', angle: 'rad' };

  it('applies the model units for the inferred quantity kind', () => {
    expect(unitSuffixForLabel('Layer Thickness', mmUnits)).toBe(' mm');
    expect(unitSuffixForLabel('Gross Area', mmUnits)).toBe(' m²');
    expect(unitSuffixForLabel('Net Volume', mmUnits)).toBe(' m³');
    expect(unitSuffixForLabel('Slope', mmUnits)).toBe(' rad');
  });

  it('keeps the legacy metric fallback when no model units are given', () => {
    expect(unitSuffixForLabel('Width')).toBe(' m');
    expect(unitSuffixForLabel('Weight')).toBe(' kg');
    expect(unitSuffixForLabel('Tilt Angle')).toBe('°');
  });

  it('returns no suffix for non-quantity labels', () => {
    expect(unitSuffixForLabel('Name', mmUnits)).toBe('');
    expect(unitSuffixForLabel('GlobalId')).toBe('');
  });
});
