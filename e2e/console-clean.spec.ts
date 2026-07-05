import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

/**
 * Program acceptance criteria §4 (IMPLEMENTATION_PLAN): zero browser console
 * errors AND warnings across boot → model load → feature exercise, repeated
 * with a responsive viewport sweep (tablet 768px, phone 375px — panels are
 * hidden ≤1023px until W3; the sweep asserts no crash/console noise, not
 * layout). Permanent regression guard added in W1.
 *
 * The suppressed three.js shader warning (AUDIT A5) never reaches the console
 * (app-level filter) and its proper fix is W2.3 scope.
 */

const ifcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc');

const waitForStatus = async (page: Page, text: string, timeout = 30_000): Promise<void> => {
  await page.waitForFunction(
    (expected) => (document.querySelector('#statusText')?.textContent || '').includes(expected),
    text,
    { timeout },
  );
};

const settleFrames = async (page: Page): Promise<void> => {
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
};

test.describe('console cleanliness (§4 program acceptance)', () => {
  test('boot, load, W1 feature exercise and viewport sweep stay console-clean', async ({ browser }) => {
    test.setTimeout(10 * 60 * 1000);
    const context = await browser.newContext({
      acceptDownloads: true,
      viewport: { width: 1600, height: 1000 },
    });
    const page = await context.newPage();

    const violations: string[] = [];
    page.on('console', (message) => {
      const type = message.type();
      if (type === 'error' || type === 'warning') {
        violations.push(`console.${type}: ${message.text()}`);
      }
    });
    page.on('pageerror', (error) => violations.push(`pageerror: ${error.message}`));

    // Boot (match the full init message — static HTML ships plain 'Ready')
    await page.goto('/');
    await waitForStatus(page, 'Ready - load IFC', 60_000);

    // Model load (conversion + registration + indexing)
    await page.setInputFiles('#fileInput', ifcPath);
    await page.waitForFunction(
      () => {
        const viewer = (window as any).__viewer;
        return !!viewer && viewer.federatedModels?.size === 1 && viewer.modelIndices?.size === 1;
      },
      undefined,
      { timeout: 180_000 },
    );
    // A15: both status slots populated side by side.
    await expect(page.locator('#loadInfo')).toHaveText(/Loaded in/);

    // F1: search renders results.
    await page.fill('#searchInput', 'wall');
    await page.click('#btnSearch');
    await waitForStatus(page, 'Search found');
    await page.click('#btnClearSearch');

    // F2: viewpoint capture (render-before-capture + thumbnail encode).
    await page.click('.tab-btn[data-tab="viewpoints"]');
    await page.fill('#viewpointName', 'Console sweep');
    await page.click('#btnSaveViewpoint');
    await waitForStatus(page, 'Saved viewpoint: Console sweep');
    await page.locator('[data-viewpoint-id]').first().click();
    await page.click('#btnApplySelectedViewpoint');
    await waitForStatus(page, 'Applied viewpoint: Console sweep');
    await page.click('#btnDeleteSelectedViewpoint');
    await page.click('.confirm-btn-confirm');
    await waitForStatus(page, 'Viewpoint deleted');

    // F3: X-ray/edges toggles.
    await page.click('#btnTransparency');
    await waitForStatus(page, 'X-ray enabled');
    await page.click('#btnWireframe');
    await waitForStatus(page, 'Edge overlay enabled');
    await page.click('#btnTransparency');
    await page.click('#btnWireframe');
    await waitForStatus(page, 'Edge overlay disabled');

    // Sections (F10 shared path) + measurements.
    await page.click('#btnSectionX');
    await waitForStatus(page, 'Section plane added');
    await page.click('#btnClearSections');
    await waitForStatus(page, 'Sections cleared');
    await page.click('#btnMeasureLength');
    await waitForStatus(page, 'Length measurement enabled');
    await page.keyboard.press('Escape');
    await waitForStatus(page, 'Active tool canceled');

    // F8: theme round-trip + background preset via the View menu.
    await page.locator('.menu-dropdown', { has: page.locator('#toggleTheme') }).locator('.menu-item').click();
    await page.click('#toggleTheme');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
    await page.click('#toggleTheme');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
    await page.click('[data-bg-preset="#0b1220"]');
    await waitForStatus(page, 'Background color set to #0b1220');
    await page.locator('.app-titlebar').click();

    // Filters (F11-adjacent safe path: no selection → show-all message).
    await page.click('.tab-btn[data-tab="explorer"]');
    await page.click('#btnApplyFilters');
    await waitForStatus(page, 'No filters selected');

    // F9/U8: issue lifecycle from a selection.
    await page.evaluate(async () => {
      const viewer = (window as any).__viewer;
      const [modelId, index] = Array.from(viewer.modelIndices.entries())[0] as [string, any];
      const firstId = Array.from(index.allIds)[0];
      await viewer.selectSingleItem(modelId, firstId, false);
    });
    await page.click('.tab-btn[data-tab="issues"]');
    await page.fill('#issueTitle', 'Console sweep issue');
    await page.click('#btnCreateIssue');
    await waitForStatus(page, 'Issue created');
    await page.locator('[data-issue-id]').first().click();
    await page.click('#btnDeleteIssue');
    await page.click('.confirm-btn-confirm');
    await waitForStatus(page, 'Issue deleted');

    // F2: screenshot export (canvas readback path).
    const download = page.waitForEvent('download');
    await page.click('#btnExportScreenshot');
    await download;
    await waitForStatus(page, 'Screenshot exported');

    // Responsive sweep (§4): tablet and phone visit with interactions — the
    // app must not crash or log; panels hidden ≤1023px is expected pre-W3.
    for (const size of [{ width: 768, height: 1024 }, { width: 375, height: 812 }]) {
      await page.setViewportSize(size);
      await settleFrames(page);
      await page.keyboard.press('f'); // fit-to-model via keyboard
      await page.mouse.click(size.width / 2, size.height / 2); // canvas raycast
      await settleFrames(page);
    }
    await page.setViewportSize({ width: 1600, height: 1000 });
    await settleFrames(page);

    expect(violations).toEqual([]);
    await context.close();
  });
});
