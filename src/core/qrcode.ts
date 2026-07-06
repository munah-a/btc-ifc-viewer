/**
 * Minimal, dependency-free QR Code generator (W4.4).
 *
 * C1 forbids a runtime CDN and we add no npm dep, so the share dialog needs its
 * own QR encoder. This is a compact byte-mode implementation (Reed–Solomon EC,
 * mask evaluation, format/version info) sufficient for the URLs the share dialog
 * produces. Returns a square boolean matrix (`true` = dark module); the caller
 * renders it (we draw to a <canvas>). Pure and unit-tested.
 *
 * Scope: byte (8-bit) mode, EC level chosen per capacity, versions 1–20 (up to
 * ~850 bytes at level L) — comfortably covers embed URLs incl. a long Blob-CDN
 * `?m=`. Throws if the data does not fit version 20.
 *
 * Algorithm reference: ISO/IEC 18004. Implementation adapted to be small and
 * self-contained (no external tables beyond the standard's constants).
 */

export type QrMatrix = boolean[][];

export type EcLevel = 'L' | 'M' | 'Q' | 'H';

/**
 * Renders a QR matrix to a canvas with a quiet zone. Dark modules use
 * `foreground`, the rest `background`. Sizes the drawing to fit the canvas'
 * current width (square). Browser-only (needs a 2D context).
 */
export function renderQrToCanvas(
  canvas: HTMLCanvasElement,
  text: string,
  opts: { ecLevel?: EcLevel; quietZone?: number; foreground?: string; background?: string } = {},
): void {
  const matrix = encodeQr(text, opts.ecLevel ?? 'M');
  const quiet = opts.quietZone ?? 4;
  const modules = matrix.length + quiet * 2;
  const ctx = canvas.getContext('2d');
  if (!ctx) return;

  const pixelSize = canvas.width;
  const scale = Math.max(1, Math.floor(pixelSize / modules));
  const drawSize = scale * modules;
  const offset = Math.floor((pixelSize - drawSize) / 2);

  ctx.fillStyle = opts.background ?? '#ffffff';
  ctx.fillRect(0, 0, pixelSize, pixelSize);
  ctx.fillStyle = opts.foreground ?? '#000000';
  for (let r = 0; r < matrix.length; r += 1) {
    for (let c = 0; c < matrix.length; c += 1) {
      if (!matrix[r][c]) continue;
      ctx.fillRect(offset + (c + quiet) * scale, offset + (r + quiet) * scale, scale, scale);
    }
  }
}

/** Public API: encode `text` (UTF-8, byte mode) into a QR module matrix. */
export function encodeQr(text: string, ecLevel: EcLevel = 'M'): QrMatrix {
  const data = utf8Bytes(text);
  const version = chooseVersion(data.length, ecLevel);
  const bits = buildDataBits(data, version, ecLevel);
  const codewords = bitsToCodewords(bits, version, ecLevel);
  const finalCodewords = interleaveWithEc(codewords, version, ecLevel);
  return buildMatrix(finalCodewords, version, ecLevel);
}

// ── UTF-8 ──
function utf8Bytes(text: string): number[] {
  return [...new TextEncoder().encode(text)];
}

// ── Capacity tables (byte-mode data capacity in bytes, versions 1..20) ──
// Source: ISO/IEC 18004 capacity tables (byte mode), one row per EC level.
const BYTE_CAPACITY: Record<EcLevel, number[]> = {
  L: [17, 32, 53, 78, 106, 134, 154, 192, 230, 271, 321, 367, 425, 458, 520, 586, 644, 718, 792, 858],
  M: [14, 26, 42, 62, 84, 106, 122, 152, 180, 213, 251, 287, 331, 362, 412, 450, 504, 560, 624, 666],
  Q: [11, 20, 32, 46, 60, 74, 86, 108, 130, 151, 177, 203, 241, 258, 292, 322, 364, 394, 442, 482],
  H: [7, 14, 24, 34, 44, 58, 64, 84, 98, 119, 137, 155, 177, 194, 220, 250, 280, 310, 338, 382],
};

// EC codewords per block + block counts, versions 1..20, per EC level.
// Format: [ecCodewordsPerBlock, numBlocksGroup1, dataCodewordsPerBlockGroup1, numBlocksGroup2, dataCodewordsPerBlockGroup2]
const EC_BLOCKS: Record<EcLevel, number[][]> = {
  L: [
    [7, 1, 19, 0, 0], [10, 1, 34, 0, 0], [15, 1, 55, 0, 0], [20, 1, 80, 0, 0], [26, 1, 108, 0, 0],
    [18, 2, 68, 0, 0], [20, 2, 78, 0, 0], [24, 2, 97, 0, 0], [30, 2, 116, 0, 0], [18, 2, 68, 2, 69],
    [20, 4, 81, 0, 0], [24, 2, 92, 2, 93], [26, 4, 107, 0, 0], [30, 3, 115, 1, 116], [22, 5, 87, 1, 88],
    [24, 5, 98, 1, 99], [28, 1, 107, 5, 108], [30, 5, 120, 1, 121], [28, 3, 113, 4, 114], [28, 3, 107, 5, 108],
  ],
  M: [
    [10, 1, 16, 0, 0], [16, 1, 28, 0, 0], [26, 1, 44, 0, 0], [18, 2, 32, 0, 0], [24, 2, 43, 0, 0],
    [16, 4, 27, 0, 0], [18, 4, 31, 0, 0], [22, 2, 38, 2, 39], [22, 3, 36, 2, 37], [26, 4, 43, 1, 44],
    [30, 1, 50, 4, 51], [22, 6, 36, 2, 37], [22, 8, 37, 1, 38], [24, 4, 40, 5, 41], [24, 5, 41, 5, 42],
    [28, 7, 45, 3, 46], [28, 10, 46, 1, 47], [26, 9, 43, 4, 44], [26, 3, 44, 11, 45], [26, 3, 41, 13, 42],
  ],
  Q: [
    [13, 1, 13, 0, 0], [22, 1, 22, 0, 0], [18, 2, 17, 0, 0], [26, 2, 24, 0, 0], [18, 2, 15, 2, 16],
    [24, 4, 19, 0, 0], [18, 2, 14, 4, 15], [22, 4, 18, 2, 19], [20, 4, 16, 4, 17], [24, 6, 19, 2, 20],
    [28, 4, 22, 4, 23], [26, 4, 20, 6, 21], [24, 8, 20, 4, 21], [20, 11, 16, 5, 17], [30, 5, 24, 7, 25],
    [24, 15, 19, 2, 20], [28, 1, 22, 15, 23], [28, 17, 22, 1, 23], [26, 17, 21, 4, 22], [30, 15, 24, 5, 25],
  ],
  H: [
    [17, 1, 9, 0, 0], [28, 1, 16, 0, 0], [22, 2, 13, 0, 0], [16, 4, 9, 0, 0], [22, 2, 11, 2, 12],
    [28, 4, 15, 0, 0], [26, 4, 13, 1, 14], [26, 4, 14, 2, 15], [24, 4, 12, 4, 13], [28, 6, 15, 2, 16],
    [24, 3, 12, 8, 13], [28, 7, 14, 4, 15], [22, 12, 11, 4, 12], [24, 11, 12, 5, 13], [24, 11, 12, 7, 13],
    [30, 3, 15, 13, 16], [28, 2, 14, 17, 15], [28, 2, 14, 19, 15], [26, 9, 13, 16, 14], [28, 15, 15, 10, 16],
  ],
};

function chooseVersion(dataLen: number, ec: EcLevel): number {
  const caps = BYTE_CAPACITY[ec];
  for (let v = 0; v < caps.length; v += 1) {
    if (dataLen <= caps[v]) return v + 1;
  }
  throw new Error(`QR: data too long (${dataLen} bytes) for version 20 at EC ${ec}`);
}

// ── Bit buffer ──
class BitBuffer {
  bits: number[] = [];
  put(value: number, length: number): void {
    for (let i = length - 1; i >= 0; i -= 1) this.bits.push((value >>> i) & 1);
  }
  get length(): number {
    return this.bits.length;
  }
}

function charCountBits(version: number): number {
  // Byte mode: 8 bits for versions 1–9, 16 for 10–40.
  return version <= 9 ? 8 : 16;
}

function totalDataCodewords(version: number, ec: EcLevel): number {
  const [ecPerBlock, g1, d1, g2, d2] = EC_BLOCKS[ec][version - 1];
  void ecPerBlock;
  return g1 * d1 + g2 * d2;
}

function buildDataBits(data: number[], version: number, ec: EcLevel): BitBuffer {
  const buffer = new BitBuffer();
  buffer.put(0b0100, 4); // byte mode indicator
  buffer.put(data.length, charCountBits(version));
  for (const byte of data) buffer.put(byte, 8);

  const capacityBits = totalDataCodewords(version, ec) * 8;
  // Terminator (up to 4 zero bits).
  const remaining = capacityBits - buffer.length;
  buffer.put(0, Math.min(4, Math.max(0, remaining)));
  // Pad to a byte boundary.
  while (buffer.length % 8 !== 0) buffer.bits.push(0);
  return buffer;
}

function bitsToCodewords(buffer: BitBuffer, version: number, ec: EcLevel): number[] {
  const codewords: number[] = [];
  for (let i = 0; i < buffer.length; i += 8) {
    let byte = 0;
    for (let j = 0; j < 8; j += 1) byte = (byte << 1) | buffer.bits[i + j];
    codewords.push(byte);
  }
  // Pad codewords 0xEC / 0x11 alternating to fill capacity.
  const need = totalDataCodewords(version, ec);
  const pads = [0xec, 0x11];
  let p = 0;
  while (codewords.length < need) {
    codewords.push(pads[p % 2]);
    p += 1;
  }
  return codewords;
}

// ── Galois field (GF(256)) for Reed–Solomon ──
const GF_EXP = new Array<number>(512);
const GF_LOG = new Array<number>(256);
(() => {
  let x = 1;
  for (let i = 0; i < 255; i += 1) {
    GF_EXP[i] = x;
    GF_LOG[x] = i;
    x <<= 1;
    if (x & 0x100) x ^= 0x11d;
  }
  for (let i = 255; i < 512; i += 1) GF_EXP[i] = GF_EXP[i - 255];
})();

function gfMul(a: number, b: number): number {
  if (a === 0 || b === 0) return 0;
  return GF_EXP[GF_LOG[a] + GF_LOG[b]];
}

function rsGeneratorPoly(degree: number): number[] {
  let poly = [1];
  for (let i = 0; i < degree; i += 1) {
    const next = new Array<number>(poly.length + 1).fill(0);
    for (let j = 0; j < poly.length; j += 1) {
      next[j] ^= poly[j];
      next[j + 1] ^= gfMul(poly[j], GF_EXP[i]);
    }
    poly = next;
  }
  return poly;
}

function rsEncode(data: number[], ecCount: number): number[] {
  const gen = rsGeneratorPoly(ecCount);
  const result = new Array<number>(ecCount).fill(0);
  for (const byte of data) {
    const factor = byte ^ result[0];
    result.shift();
    result.push(0);
    for (let j = 0; j < gen.length - 1; j += 1) {
      result[j] ^= gfMul(gen[j + 1], factor);
    }
  }
  return result;
}

function interleaveWithEc(dataCodewords: number[], version: number, ec: EcLevel): number[] {
  const [ecPerBlock, g1, d1, g2, d2] = EC_BLOCKS[ec][version - 1];
  const blocks: { data: number[]; ecc: number[] }[] = [];
  let offset = 0;
  for (let b = 0; b < g1; b += 1) {
    const chunk = dataCodewords.slice(offset, offset + d1);
    offset += d1;
    blocks.push({ data: chunk, ecc: rsEncode(chunk, ecPerBlock) });
  }
  for (let b = 0; b < g2; b += 1) {
    const chunk = dataCodewords.slice(offset, offset + d2);
    offset += d2;
    blocks.push({ data: chunk, ecc: rsEncode(chunk, ecPerBlock) });
  }

  const result: number[] = [];
  const maxData = Math.max(d1, d2);
  for (let i = 0; i < maxData; i += 1) {
    for (const block of blocks) if (i < block.data.length) result.push(block.data[i]);
  }
  for (let i = 0; i < ecPerBlock; i += 1) {
    for (const block of blocks) result.push(block.ecc[i]);
  }
  return result;
}

// ── Matrix construction ──
function sizeForVersion(version: number): number {
  return version * 4 + 17;
}

function buildMatrix(codewords: number[], version: number, ec: EcLevel): QrMatrix {
  const size = sizeForVersion(version);
  const modules: (boolean | null)[][] = Array.from({ length: size }, () =>
    new Array<boolean | null>(size).fill(null),
  );

  placeFinderPatterns(modules, size);
  placeSeparators(modules, size);
  placeAlignmentPatterns(modules, version);
  placeTimingPatterns(modules, size);
  // Dark module.
  modules[size - 8][8] = true;
  reserveFormatInfo(modules, size);
  if (version >= 7) reserveVersionInfo(modules, size);

  placeData(modules, codewords, size);

  // Try all 8 masks, pick the lowest penalty.
  let best: { mask: number; grid: boolean[][]; penalty: number } | null = null;
  for (let mask = 0; mask < 8; mask += 1) {
    const grid = applyMask(modules, mask, size);
    writeFormatInfo(grid, ec, mask, size);
    if (version >= 7) writeVersionInfo(grid, version, size);
    const penalty = evaluatePenalty(grid, size);
    if (!best || penalty < best.penalty) best = { mask, grid, penalty };
  }
  return best!.grid;
}

function setModule(m: (boolean | null)[][], r: number, c: number, v: boolean): void {
  m[r][c] = v;
}

function placeFinderPattern(m: (boolean | null)[][], row: number, col: number): void {
  for (let r = -1; r <= 7; r += 1) {
    for (let c = -1; c <= 7; c += 1) {
      const rr = row + r;
      const cc = col + c;
      if (rr < 0 || cc < 0 || rr >= m.length || cc >= m.length) continue;
      const isBorder = r >= 0 && r <= 6 && (c === 0 || c === 6);
      const isBorderH = c >= 0 && c <= 6 && (r === 0 || r === 6);
      const isCenter = r >= 2 && r <= 4 && c >= 2 && c <= 4;
      setModule(m, rr, cc, isBorder || isBorderH || isCenter);
    }
  }
}

function placeFinderPatterns(m: (boolean | null)[][], size: number): void {
  placeFinderPattern(m, 0, 0);
  placeFinderPattern(m, 0, size - 7);
  placeFinderPattern(m, size - 7, 0);
}

function placeSeparators(m: (boolean | null)[][], size: number): void {
  for (let i = 0; i < 8; i += 1) {
    // top-left
    if (m[7][i] === null) m[7][i] = false;
    if (m[i][7] === null) m[i][7] = false;
    // top-right
    if (m[7][size - 1 - i] === null) m[7][size - 1 - i] = false;
    if (m[i][size - 8] === null) m[i][size - 8] = false;
    // bottom-left
    if (m[size - 8][i] === null) m[size - 8][i] = false;
    if (m[size - 1 - i][7] === null) m[size - 1 - i][7] = false;
  }
}

// Alignment-pattern center coordinates per version (ISO/IEC 18004 Annex E).
const ALIGN_POS: number[][] = [
  [], [6, 18], [6, 22], [6, 26], [6, 30], [6, 34], [6, 22, 38], [6, 24, 42], [6, 26, 46], [6, 28, 50],
  [6, 30, 54], [6, 32, 58], [6, 34, 62], [6, 26, 46, 66], [6, 26, 48, 70], [6, 26, 50, 74], [6, 30, 54, 78],
  [6, 30, 56, 82], [6, 30, 58, 86], [6, 34, 62, 90],
];

function placeAlignmentPatterns(m: (boolean | null)[][], version: number): void {
  if (version < 2) return;
  const positions = ALIGN_POS[version - 1];
  for (const r of positions) {
    for (const c of positions) {
      // Skip where finder patterns already are.
      if ((r === 6 && c === 6) || (r === 6 && c === positions[positions.length - 1]) ||
        (c === 6 && r === positions[positions.length - 1])) continue;
      if (m[r][c] !== null) continue;
      for (let dr = -2; dr <= 2; dr += 1) {
        for (let dc = -2; dc <= 2; dc += 1) {
          const ring = Math.max(Math.abs(dr), Math.abs(dc));
          setModule(m, r + dr, c + dc, ring !== 1);
        }
      }
    }
  }
}

function placeTimingPatterns(m: (boolean | null)[][], size: number): void {
  for (let i = 8; i < size - 8; i += 1) {
    if (m[6][i] === null) m[6][i] = i % 2 === 0;
    if (m[i][6] === null) m[i][6] = i % 2 === 0;
  }
}

function reserveFormatInfo(m: (boolean | null)[][], size: number): void {
  for (let i = 0; i < 9; i += 1) {
    if (m[8][i] === null) m[8][i] = false;
    if (m[i][8] === null) m[i][8] = false;
  }
  for (let i = 0; i < 8; i += 1) {
    if (m[8][size - 1 - i] === null) m[8][size - 1 - i] = false;
    if (m[size - 1 - i][8] === null) m[size - 1 - i][8] = false;
  }
}

function reserveVersionInfo(m: (boolean | null)[][], size: number): void {
  for (let i = 0; i < 6; i += 1) {
    for (let j = 0; j < 3; j += 1) {
      if (m[i][size - 11 + j] === null) m[i][size - 11 + j] = false;
      if (m[size - 11 + j][i] === null) m[size - 11 + j][i] = false;
    }
  }
}

function placeData(m: (boolean | null)[][], codewords: number[], size: number): void {
  const bits: number[] = [];
  for (const cw of codewords) for (let i = 7; i >= 0; i -= 1) bits.push((cw >>> i) & 1);

  let bitIndex = 0;
  let upward = true;
  for (let col = size - 1; col > 0; col -= 2) {
    if (col === 6) col -= 1; // skip the timing column
    for (let i = 0; i < size; i += 1) {
      const row = upward ? size - 1 - i : i;
      for (let c = 0; c < 2; c += 1) {
        const cc = col - c;
        if (m[row][cc] !== null) continue;
        const bit = bitIndex < bits.length ? bits[bitIndex] : 0;
        bitIndex += 1;
        m[row][cc] = bit === 1;
      }
    }
    upward = !upward;
  }
}

// Function-pattern mask: which modules are reserved (not data) — recomputed by
// checking a fresh grid built with only the patterns.
function isFunctionModule(size: number, version: number, r: number, c: number): boolean {
  const fresh: (boolean | null)[][] = Array.from({ length: size }, () =>
    new Array<boolean | null>(size).fill(null),
  );
  placeFinderPatterns(fresh, size);
  placeSeparators(fresh, size);
  placeAlignmentPatterns(fresh, version);
  placeTimingPatterns(fresh, size);
  fresh[size - 8][8] = true;
  reserveFormatInfo(fresh, size);
  if (version >= 7) reserveVersionInfo(fresh, size);
  return fresh[r][c] !== null;
}

function applyMask(m: (boolean | null)[][], mask: number, size: number): boolean[][] {
  const version = (size - 17) / 4;
  const out: boolean[][] = Array.from({ length: size }, () => new Array<boolean>(size).fill(false));
  for (let r = 0; r < size; r += 1) {
    for (let c = 0; c < size; c += 1) {
      const base = m[r][c] === true;
      if (isFunctionModule(size, version, r, c)) {
        out[r][c] = base;
        continue;
      }
      out[r][c] = base !== maskCondition(mask, r, c);
    }
  }
  return out;
}

function maskCondition(mask: number, r: number, c: number): boolean {
  switch (mask) {
    case 0: return (r + c) % 2 === 0;
    case 1: return r % 2 === 0;
    case 2: return c % 3 === 0;
    case 3: return (r + c) % 3 === 0;
    case 4: return (Math.floor(r / 2) + Math.floor(c / 3)) % 2 === 0;
    case 5: return ((r * c) % 2) + ((r * c) % 3) === 0;
    case 6: return (((r * c) % 2) + ((r * c) % 3)) % 2 === 0;
    case 7: return (((r + c) % 2) + ((r * c) % 3)) % 2 === 0;
    default: return false;
  }
}

// Format info (EC level + mask) with BCH(15,5) + mask 0x5412.
const EC_FORMAT_BITS: Record<EcLevel, number> = { L: 0b01, M: 0b00, Q: 0b11, H: 0b10 };

function writeFormatInfo(grid: boolean[][], ec: EcLevel, mask: number, size: number): void {
  const data = (EC_FORMAT_BITS[ec] << 3) | mask;
  let bch = data << 10;
  const g = 0b10100110111;
  for (let i = 4; i >= 0; i -= 1) {
    if ((bch >>> (10 + i)) & 1) bch ^= g << i;
  }
  const format = ((data << 10) | bch) ^ 0b101010000010010;
  const bits: number[] = [];
  for (let i = 14; i >= 0; i -= 1) bits.push((format >>> i) & 1);

  // Around top-left finder.
  for (let i = 0; i <= 5; i += 1) grid[8][i] = bits[i] === 1;
  grid[8][7] = bits[6] === 1;
  grid[8][8] = bits[7] === 1;
  grid[7][8] = bits[8] === 1;
  for (let i = 9; i <= 14; i += 1) grid[14 - i][8] = bits[i] === 1;

  // Around top-right + bottom-left.
  for (let i = 0; i <= 7; i += 1) grid[size - 1 - i][8] = bits[i] === 1;
  for (let i = 8; i <= 14; i += 1) grid[8][size - 15 + i] = bits[i] === 1;
  grid[size - 8][8] = true; // dark module stays
}

// Version info (BCH(18,6)) for versions >= 7.
function writeVersionInfo(grid: boolean[][], version: number, size: number): void {
  let bch = version << 12;
  const g = 0b1111100100101;
  for (let i = 5; i >= 0; i -= 1) {
    if ((bch >>> (12 + i)) & 1) bch ^= g << i;
  }
  const versionBits = (version << 12) | bch;
  for (let i = 0; i < 18; i += 1) {
    const bit = ((versionBits >>> i) & 1) === 1;
    const r = Math.floor(i / 3);
    const c = i % 3;
    grid[r][size - 11 + c] = bit;
    grid[size - 11 + c][r] = bit;
  }
}

// Penalty scoring (ISO/IEC 18004 §8.8.2) for mask selection.
function evaluatePenalty(grid: boolean[][], size: number): number {
  let penalty = 0;

  // Rule 1: runs of 5+ same-color in row/col.
  for (let r = 0; r < size; r += 1) {
    penalty += runPenalty(grid[r]);
  }
  for (let c = 0; c < size; c += 1) {
    const col = grid.map((row) => row[c]);
    penalty += runPenalty(col);
  }

  // Rule 2: 2x2 blocks of same color.
  for (let r = 0; r < size - 1; r += 1) {
    for (let c = 0; c < size - 1; c += 1) {
      const v = grid[r][c];
      if (v === grid[r][c + 1] && v === grid[r + 1][c] && v === grid[r + 1][c + 1]) penalty += 3;
    }
  }

  // Rule 3: finder-like 1:1:3:1:1 patterns.
  const pat1 = [true, false, true, true, true, false, true, false, false, false, false];
  const pat2 = [false, false, false, false, true, false, true, true, true, false, true];
  for (let r = 0; r < size; r += 1) {
    for (let c = 0; c <= size - 11; c += 1) {
      if (matches(grid[r], c, pat1) || matches(grid[r], c, pat2)) penalty += 40;
    }
  }
  for (let c = 0; c < size; c += 1) {
    const col = grid.map((row) => row[c]);
    for (let r = 0; r <= size - 11; r += 1) {
      if (matches(col, r, pat1) || matches(col, r, pat2)) penalty += 40;
    }
  }

  // Rule 4: dark/light balance.
  let dark = 0;
  for (const row of grid) for (const cell of row) if (cell) dark += 1;
  const ratio = (dark * 100) / (size * size);
  const k = Math.floor(Math.abs(ratio - 50) / 5);
  penalty += k * 10;

  return penalty;
}

function runPenalty(line: boolean[]): number {
  let penalty = 0;
  let runColor = line[0];
  let runLen = 1;
  for (let i = 1; i < line.length; i += 1) {
    if (line[i] === runColor) {
      runLen += 1;
    } else {
      if (runLen >= 5) penalty += 3 + (runLen - 5);
      runColor = line[i];
      runLen = 1;
    }
  }
  if (runLen >= 5) penalty += 3 + (runLen - 5);
  return penalty;
}

function matches(line: boolean[], start: number, pattern: boolean[]): boolean {
  for (let i = 0; i < pattern.length; i += 1) {
    if (line[start + i] !== pattern[i]) return false;
  }
  return true;
}
