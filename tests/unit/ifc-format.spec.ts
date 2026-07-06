import { describe, expect, it } from 'vitest';

import { isProbablyIfc } from '../../src/core/ifc-format';

const bytes = (text: string): Uint8Array => new TextEncoder().encode(text);

describe('isProbablyIfc (AUDIT T13 — deterministic non-IFC rejection)', () => {
  it('accepts a STEP/IFC header', () => {
    expect(isProbablyIfc(bytes("ISO-10303-21;\nHEADER;\nFILE_DESCRIPTION(('ViewDefinition'),'2;1');"))).toBe(true);
  });

  it('accepts a header after a UTF-8 BOM and leading whitespace', () => {
    expect(isProbablyIfc(bytes('﻿  \nISO-10303-21;'))).toBe(true);
  });

  it('rejects arbitrary garbage (the U4 corrupt-file case)', () => {
    expect(isProbablyIfc(bytes('NOT-AN-IFC-FILE'))).toBe(false);
  });

  it('rejects an empty file', () => {
    expect(isProbablyIfc(new Uint8Array(0))).toBe(false);
  });

  it('rejects other structured text (e.g. JSON, HTML)', () => {
    expect(isProbablyIfc(bytes('{"not":"ifc"}'))).toBe(false);
    expect(isProbablyIfc(bytes('<!DOCTYPE html><html></html>'))).toBe(false);
  });
});
