import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
  en,
  de,
  t,
  tIn,
  setLanguage,
  getLanguage,
  initLanguage,
  coerceLanguage,
  onLanguageChange,
  formatDate,
  formatDateTime,
  formatNumber,
  SUPPORTED_LANGUAGES,
  DEFAULT_LANGUAGE,
  type MessageKey,
} from '../../src/core/i18n';

// Reset to the default language between tests (module state is a singleton).
beforeEach(() => setLanguage('en'));
afterEach(() => setLanguage('en'));

describe('catalog integrity', () => {
  it('de has exactly the same key set as en (no missing/extra keys)', () => {
    const enKeys = Object.keys(en).sort();
    const deKeys = Object.keys(de).sort();
    expect(deKeys).toEqual(enKeys);
  });

  it('no catalog value is empty', () => {
    for (const key of Object.keys(en) as MessageKey[]) {
      expect(en[key].length, `en ${key}`).toBeGreaterThan(0);
      expect(de[key].length, `de ${key}`).toBeGreaterThan(0);
    }
  });

  it('every {param} placeholder in en also appears in de (and vice versa)', () => {
    const tokensOf = (s: string): string[] => (s.match(/\{(\w+)\}/g) ?? []).sort();
    for (const key of Object.keys(en) as MessageKey[]) {
      expect(tokensOf(de[key]), `placeholders mismatch for "${key}"`).toEqual(tokensOf(en[key]));
    }
  });

  it('DEFAULT_LANGUAGE is en (existing e2e assert English strings)', () => {
    expect(DEFAULT_LANGUAGE).toBe('en');
  });
});

describe('coerceLanguage', () => {
  it('accepts supported languages', () => {
    expect(coerceLanguage('en')).toBe('en');
    expect(coerceLanguage('de')).toBe('de');
  });
  it('rejects anything else', () => {
    expect(coerceLanguage('fr')).toBeNull();
    expect(coerceLanguage('')).toBeNull();
    expect(coerceLanguage(null)).toBeNull();
    expect(coerceLanguage(42)).toBeNull();
  });
  it('SUPPORTED_LANGUAGES is en + de', () => {
    expect([...SUPPORTED_LANGUAGES]).toEqual(['en', 'de']);
  });
});

describe('t() translation + interpolation', () => {
  it('returns the English string by default', () => {
    expect(t('status.ready')).toBe('Ready — load IFC model(s)');
  });

  it('returns the German string after switching', () => {
    setLanguage('de');
    expect(getLanguage()).toBe('de');
    expect(t('status.ready')).toBe('Bereit — IFC-Modell(e) laden');
  });

  it('interpolates named params', () => {
    expect(t('status.modelLoaded', { name: 'tower.ifc' })).toBe('Model loaded: tower.ifc');
    setLanguage('de');
    expect(t('status.modelLoaded', { name: 'tower.ifc' })).toBe('Modell geladen: tower.ifc');
  });

  it('interpolates numeric params and leaves unknown placeholders intact', () => {
    expect(t('status.selectedFullModel', { count: 1526 })).toBe('Selected full model (1526 elements)');
    // A template token with no matching param is left as-is (defensive).
    expect(t('status.contextFailed', { context: 'Load' })).toBe('Load failed: {message}');
  });

  it('tIn() translates in an explicit language regardless of current', () => {
    expect(tIn('de', 'panel.help')).toBe('Hilfe');
    expect(tIn('en', 'panel.help')).toBe('Help');
    expect(getLanguage()).toBe('en'); // unchanged
  });
});

describe('setLanguage subscribers', () => {
  it('notifies subscribers on change and not on no-op', () => {
    const seen: string[] = [];
    const off = onLanguageChange((lang) => seen.push(lang));
    setLanguage('de'); // change → notify
    setLanguage('de'); // no-op → no notify
    setLanguage('en'); // change → notify
    off();
    setLanguage('de'); // after unsubscribe → not recorded
    expect(seen).toEqual(['de', 'en']);
  });
});

describe('initLanguage', () => {
  it('falls back to the default when no storage / no stored value', () => {
    // Under Node there is no localStorage; initLanguage must not throw.
    expect(() => initLanguage()).not.toThrow();
    expect(initLanguage()).toBe('en');
  });
});

describe('formatDate — Swiss DD.MM.YYYY', () => {
  const date = new Date('2026-03-09T13:05:00');

  it('formats DD.MM.YYYY with leading zeros in both languages', () => {
    expect(formatDate(date, 'en')).toBe('09.03.2026');
    expect(formatDate(date, 'de')).toBe('09.03.2026');
  });

  it('accepts ISO strings and epoch ms', () => {
    expect(formatDate('2026-03-09T13:05:00', 'de')).toBe('09.03.2026');
    expect(formatDate(date.getTime(), 'en')).toBe('09.03.2026');
  });

  it('uses the current language when none is passed', () => {
    setLanguage('de');
    expect(formatDate(date)).toBe('09.03.2026');
  });

  it('returns the em-dash placeholder for an invalid date', () => {
    expect(formatDate('not-a-date')).toBe('—');
    expect(formatDate(NaN)).toBe('—');
  });
});

describe('formatDateTime — DD.MM.YYYY, HH:MM (24h)', () => {
  it('appends 24-hour time', () => {
    const date = new Date('2026-03-09T13:05:00');
    expect(formatDateTime(date, 'en')).toBe('09.03.2026, 13:05');
    expect(formatDateTime(date, 'de')).toBe('09.03.2026, 13:05');
  });

  it('returns em-dash for an invalid date', () => {
    expect(formatDateTime('nope')).toBe('—');
  });
});

describe('formatNumber — apostrophe grouping, CHF', () => {
  it("groups thousands with an apostrophe (Swiss)", () => {
    expect(formatNumber(1234567, 'en')).toBe("1'234'567");
    expect(formatNumber(1234567, 'de')).toBe("1'234'567");
  });

  it("formats CHF currency as CHF 1'234.50 (NBSP after the code)", () => {
    const chf: Intl.NumberFormatOptions = { style: 'currency', currency: 'CHF' };
    // Intl places a non-breaking space (U+00A0) between the currency code and
    // the amount — that is the correct Swiss rendering the brand voice wants.
    expect(formatNumber(1234.5, 'en', chf)).toBe("CHF 1'234.50");
    expect(formatNumber(1234.5, 'de', chf)).toBe("CHF 1'234.50");
  });

  it('respects fraction-digit options', () => {
    expect(formatNumber(3.14159, 'en', { maximumFractionDigits: 2 })).toBe('3.14');
  });

  it('returns em-dash for non-finite input', () => {
    expect(formatNumber(NaN)).toBe('—');
    expect(formatNumber(Infinity)).toBe('—');
  });
});
