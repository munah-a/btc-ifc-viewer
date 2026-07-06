import path from 'node:path';

import AxeBuilder from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';

// W3.5 accessibility smoke (AUDIT U3/U5/U6/U7/U8/U11). Runs axe-core against the
// rebuilt shell in both themes, with a model loaded, and fails on any
// serious/critical violation. This is a smoke — not a full manual audit — but it
// catches the regression classes the U-series flagged (missing names, bad
// contrast on core surfaces, broken roles).

const ifcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc');
const CI = Boolean(process.env.CI);
const viewerUrl = '/';

const waitForReady = async (page: Page): Promise<void> => {
  await page.goto(viewerUrl);
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
    undefined,
    { timeout: CI ? 120_000 : 60_000 },
  );
};

const loadModel = async (page: Page): Promise<void> => {
  await page.setInputFiles('#fileInput', ifcPath);
  await page.waitForFunction(
    () => {
      const api = window.__viewerTestApi;
      return !!api && api.modelCount() === 1 && api.indexedModelCount() === 1;
    },
    undefined,
    { timeout: CI ? 300_000 : 180_000 },
  );
};

// Only fail on serious/critical outcomes — the canvas + third-party WebGL DOM
// can surface minor best-practice notes we do not own.
const runAxe = async (page: Page): Promise<void> => {
  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa'])
    // The viewport canvas is a WebGL surface with no semantic content.
    .exclude('#viewer-container')
    .analyze();
  const serious = results.violations.filter(
    (v) => v.impact === 'serious' || v.impact === 'critical',
  );
  const summary = serious
    .map((v) => `${v.id} (${v.impact}): ${v.nodes.length} node(s) — ${v.help}`)
    .join('\n');
  expect(serious, `Accessibility violations:\n${summary}`).toEqual([]);
};

test.describe('accessibility smoke', () => {
  test('dark theme (empty + loaded) has no serious axe violations', async ({ page }) => {
    await waitForReady(page);
    await runAxe(page); // empty state
    await loadModel(page);
    // Exercise a couple of tabs so their panels are in the a11y tree.
    await page.click('#tab-properties');
    await page.click('#tab-explorer');
    await runAxe(page);
  });

  test('light theme (loaded) has no serious axe violations', async ({ page }) => {
    await waitForReady(page);
    await page.click('#btnThemeToggle');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
    await loadModel(page);
    await runAxe(page);
  });
});
