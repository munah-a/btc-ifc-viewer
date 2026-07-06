import { expect, test } from '@playwright/test';

/**
 * AUDIT A18 regression guard: hydrated inline icons must actually render.
 * The prior DOMParser+adoption path produced mis-namespaced <svg>s that were
 * 0×0 with black fill — invisible across the whole rebranded chrome, yet
 * silent (no console error) so only screenshots caught it. This asserts a
 * hydrated tool-rail icon has real size and follows currentColor.
 */
test.describe('icon hydration (A18)', () => {
  test('hydrated icons render at nonzero size in the SVG namespace', async ({ page }) => {
    await page.goto('/');
    await page.waitForFunction(
      () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
      undefined,
      { timeout: 60_000 },
    );
    const info = await page.evaluate(() => {
      const btn = [...document.querySelectorAll('.rail-btn')].find((b) => b.getBoundingClientRect().width > 0);
      const svg = btn?.querySelector('svg') ?? null;
      const rect = svg?.getBoundingClientRect();
      return {
        found: !!svg,
        ns: svg?.namespaceURI ?? null,
        w: rect ? Math.round(rect.width) : 0,
        h: rect ? Math.round(rect.height) : 0,
        fill: svg ? getComputedStyle(svg as unknown as Element).fill : null,
      };
    });
    expect(info.found).toBe(true);
    expect(info.ns).toBe('http://www.w3.org/2000/svg');
    expect(info.w).toBeGreaterThan(0);
    expect(info.h).toBeGreaterThan(0);
    // currentColor honored → not the initial black fill.
    expect(info.fill).not.toBe('rgb(0, 0, 0)');
  });
});
