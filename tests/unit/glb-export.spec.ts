import { describe, expect, it } from 'vitest';

import { GLB_MAGIC, hasActiveClipping, isValidGlb } from '../../src/core/glb-export';

/** Builds a minimal, well-formed GLB header (12-byte header + given total length). */
function makeGlb(opts: { magicOk?: boolean; version?: number; declaredLen?: number; totalLen?: number } = {}): ArrayBuffer {
  const totalLen = opts.totalLen ?? 12;
  const buffer = new ArrayBuffer(totalLen);
  const view = new DataView(buffer);
  const bytes = new Uint8Array(buffer);
  if (opts.magicOk === false) {
    bytes.set([0x00, 0x00, 0x00, 0x00], 0);
  } else {
    bytes.set([0x67, 0x6c, 0x54, 0x46], 0); // 'glTF'
  }
  view.setUint32(4, opts.version ?? 2, true);
  view.setUint32(8, opts.declaredLen ?? totalLen, true);
  return buffer;
}

describe('glb-export · isValidGlb', () => {
  it('accepts a well-formed GLB header (magic + version 2 + matching length)', () => {
    expect(isValidGlb(makeGlb())).toBe(true);
    expect(isValidGlb(makeGlb({ totalLen: 64 }))).toBe(true);
  });

  it('exposes the little-endian glTF magic constant', () => {
    expect(GLB_MAGIC).toBe(0x46546c67);
  });

  it('rejects a buffer that is too short', () => {
    expect(isValidGlb(new ArrayBuffer(4))).toBe(false);
    expect(isValidGlb(new ArrayBuffer(0))).toBe(false);
  });

  it('rejects a wrong magic', () => {
    expect(isValidGlb(makeGlb({ magicOk: false }))).toBe(false);
  });

  it('rejects a non-2 version', () => {
    expect(isValidGlb(makeGlb({ version: 1 }))).toBe(false);
  });

  it('rejects a declared length that does not match the buffer', () => {
    expect(isValidGlb(makeGlb({ totalLen: 32, declaredLen: 999 }))).toBe(false);
  });
});

describe('glb-export · hasActiveClipping', () => {
  it('is true only when the clipper is enabled AND has planes', () => {
    expect(hasActiveClipping(true, 2)).toBe(true);
    expect(hasActiveClipping(true, 0)).toBe(false);
    expect(hasActiveClipping(false, 3)).toBe(false);
    expect(hasActiveClipping(false, 0)).toBe(false);
  });
});
