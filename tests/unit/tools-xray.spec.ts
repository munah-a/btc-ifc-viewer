import { describe, expect, it } from 'vitest';
import { computeXrayOpacity, XRAY_OPACITY_FACTOR } from '../../src/tools/xray';

describe('computeXrayOpacity', () => {
  it('returns 1 (fully opaque) when x-ray is off and base is opaque', () => {
    expect(computeXrayOpacity(1, false)).toBe(1);
  });

  it('dims by the x-ray factor when enabled', () => {
    // 1 * 0.28 = 0.28
    expect(computeXrayOpacity(1, true)).toBeCloseTo(XRAY_OPACITY_FACTOR, 5);
  });

  it('clamps a dimmed value into the visible-but-transparent band [0.02, 0.999]', () => {
    // A tiny base under x-ray must not vanish entirely.
    expect(computeXrayOpacity(0.01, true)).toBeGreaterThanOrEqual(0.02);
    // A mid base under x-ray stays strictly transparent.
    const dimmed = computeXrayOpacity(0.9, true);
    expect(dimmed).toBeGreaterThanOrEqual(0.02);
    expect(dimmed).toBeLessThanOrEqual(0.999);
  });

  it('snaps effectively-opaque (>= 0.999) to exactly 1', () => {
    expect(computeXrayOpacity(0.9995, false)).toBe(1);
    expect(computeXrayOpacity(0.999, false)).toBe(1);
  });

  it('passes the base opacity through unchanged when x-ray is off', () => {
    expect(computeXrayOpacity(0.5, false)).toBeCloseTo(0.5, 5);
  });
});
