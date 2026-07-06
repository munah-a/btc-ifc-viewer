/**
 * X-ray tool — pure opacity math.
 *
 * The per-model opacity application (calling the fragments model's
 * setOpacity/resetOpacity and tracking the last-applied value to skip no-ops)
 * stays in the orchestrator. This module holds the DOM-free/engine-free
 * computation of the target opacity so it is unit-testable.
 */

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

/** The x-ray dimming factor applied to a model's base opacity when active. */
export const XRAY_OPACITY_FACTOR = 0.28;

/**
 * The opacity to apply to a model given its base opacity and whether x-ray is
 * on. Returns 1 for effectively-opaque (so the caller resets to fully opaque),
 * otherwise a value clamped into the visible-but-transparent band [0.02, 0.999]
 * and rounded to 3 dp (matching the original inline logic exactly).
 */
export function computeXrayOpacity(baseOpacity: number, xrayEnabled: boolean): number {
  const base = clamp(baseOpacity, 0, 1);
  const effective = xrayEnabled ? clamp(base * XRAY_OPACITY_FACTOR, 0, 1) : base;
  return effective >= 0.999 ? 1 : clamp(Number(effective.toFixed(3)), 0.02, 0.999);
}
