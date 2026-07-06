import { describe, expect, it } from 'vitest';

import { encodeQr, type QrMatrix } from '../../src/core/qrcode';

function isSquare(m: QrMatrix): boolean {
  return m.length > 0 && m.every((row) => row.length === m.length);
}

/** A 7x7 finder pattern (dark border ring + dark 3x3 center) at (row,col). */
function hasFinder(m: QrMatrix, row: number, col: number): boolean {
  for (let r = 0; r < 7; r += 1) {
    for (let c = 0; c < 7; c += 1) {
      const onBorder = r === 0 || r === 6 || c === 0 || c === 6;
      const inCenter = r >= 2 && r <= 4 && c >= 2 && c <= 4;
      const expectDark = onBorder || inCenter;
      if (m[row + r][col + c] !== expectDark) return false;
    }
  }
  return true;
}

describe('qrcode · structure', () => {
  it('produces a square version-1 (21x21) matrix for short input', () => {
    const m = encodeQr('hi', 'M');
    expect(isSquare(m)).toBe(true);
    expect(m.length).toBe(21); // version 1
  });

  it('places all three finder patterns (top-left, top-right, bottom-left)', () => {
    const m = encodeQr('https://btc-ifc-viewer-2.vercel.app/embed.html?m=abc', 'M');
    const size = m.length;
    expect(hasFinder(m, 0, 0)).toBe(true);
    expect(hasFinder(m, 0, size - 7)).toBe(true);
    expect(hasFinder(m, size - 7, 0)).toBe(true);
  });

  it('has correct timing patterns (alternating on row/col 6)', () => {
    const m = encodeQr('timing-check', 'M');
    const size = m.length;
    for (let i = 8; i < size - 8; i += 1) {
      expect(m[6][i]).toBe(i % 2 === 0);
      expect(m[i][6]).toBe(i % 2 === 0);
    }
  });

  it('sets the dark module at (4*version+9, 8)', () => {
    const m = encodeQr('dark-module', 'M');
    const size = m.length;
    expect(m[size - 8][8]).toBe(true);
  });

  it('grows the version with the data length', () => {
    const short = encodeQr('a', 'M').length;
    const long = encodeQr('x'.repeat(200), 'M').length;
    expect(long).toBeGreaterThan(short);
  });

  it('is deterministic for the same input', () => {
    const a = encodeQr('deterministic', 'M');
    const b = encodeQr('deterministic', 'M');
    expect(a).toEqual(b);
  });

  it('encodes a long Blob-CDN embed URL without throwing', () => {
    const url =
      'https://btc-ifc-viewer-2.vercel.app/embed.html?m=' +
      encodeURIComponent('https://xyz.public.blob.vercel-storage.com/frags/abcdef1234567890.frag') +
      '&id=abcdef1234567890';
    expect(() => encodeQr(url, 'M')).not.toThrow();
    const m = encodeQr(url, 'M');
    expect(isSquare(m)).toBe(true);
  });

  it('throws for data exceeding version-20 capacity', () => {
    expect(() => encodeQr('x'.repeat(5000), 'M')).toThrow();
  });
});
