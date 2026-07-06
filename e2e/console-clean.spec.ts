import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

import { applyCpuThrottle } from './_cpu-throttle';
// Pulls in the `Window.__viewerTestApi` global augmentation (T6 contract).
import '../src/core/test-api';

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

// GitHub's 2-core SwiftShader runner is ~2-3x slower than a dev GPU box, so
// this comprehensive sweep (load + ~25 interactions + viewport passes) needs a
// generously CI-scaled budget — a fast local pass does NOT prove CI timing.
const CI = Boolean(process.env.CI);

// Browser/GL-driver-emitted diagnostics — NOT application console calls. The
// WebGL layer logs a "GPU stall due to ReadPixels" performance hint whenever
// the canvas is read back (screenshot export, viewpoint snapshots — both are
// deliberate features); it fires on SwiftShader (CI) and some real GPUs alike
// and the app cannot suppress it without globally monkey-patching console (the
// AUDIT A5 anti-pattern we are removing). The gate stays strict for every other
// error/warning; only this exact driver-performance class is ignored.
const isEnvironmentalNoise = (text: string): boolean =>
  /GL Driver Message \(OpenGL, Performance/i.test(text)
  || /GPU stall due to ReadPixels/i.test(text);

const waitForStatus = async (page: Page, text: string, timeout = CI ? 45_000 : 30_000): Promise<void> => {
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
    // ~5min local (many-core SwiftShader) → ~15min on GitHub's 2-core runner
    // (render-bound flows run ~3x slower there); 30min ceiling = ample margin.
    test.setTimeout(CI ? 30 * 60 * 1000 : 10 * 60 * 1000);
    const context = await browser.newContext({
      acceptDownloads: true,
      viewport: { width: 1600, height: 1000 },
    });
    const page = await context.newPage();
    await applyCpuThrottle(page);

    const violations: string[] = [];
    page.on('console', (message) => {
      const type = message.type();
      if ((type === 'error' || type === 'warning') && !isEnvironmentalNoise(message.text())) {
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
        const api = window.__viewerTestApi;
        return !!api && api.modelCount() === 1 && api.indexedModelCount() === 1;
      },
      undefined,
      { timeout: 180_000 },
    );
    // A15: both status slots populated side by side.
    await expect(page.locator('#loadInfo')).toHaveText(/Loaded in/);

    // F1: search renders results (live-on-input, debounced).
    await page.fill('#searchInput', 'wall');
    await waitForStatus(page, 'Search found');
    await page.click('#btnClearSearch');

    // F2: viewpoint capture (render-before-capture + thumbnail encode).
    await page.click('#tab-viewpoints');
    await page.fill('#viewpointName', 'Console sweep');
    await page.click('#btnSaveViewpoint');
    await waitForStatus(page, 'Saved viewpoint: Console sweep');
    await page.locator('[data-viewpoint-id]').first().locator('[data-viewpoint-action="apply"]').click();
    await waitForStatus(page, 'Applied viewpoint: Console sweep');
    await page.locator('[data-viewpoint-id]').first().locator('[data-viewpoint-action="delete"]').click();
    await page.click('#confirmOk');
    await waitForStatus(page, 'Viewpoint deleted');

    // F3: X-ray/edges toggles (rail buttons).
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

    // F8: theme round-trip (topbar toggle, per-theme background memory).
    await page.click('#btnThemeToggle');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
    await page.click('#btnThemeToggle');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');

    // F9/U8: issue lifecycle from a selection.
    await page.evaluate(async () => {
      const context = window.__viewerTestApi?.firstModelContext();
      if (context) await window.__viewerTestApi?.selectItem(context.modelId, context.firstItemId, false);
    });
    await page.click('#tab-issues');
    await page.fill('#issueTitle', 'Console sweep issue');
    await page.click('#btnCreateIssue');
    await waitForStatus(page, 'Issue created');
    await page.locator('[data-issue-id]').first().locator('[data-issue-action="delete"]').click();
    await page.click('#confirmOk');
    await waitForStatus(page, 'Issue deleted');

    // F2: screenshot export (canvas readback path).
    const download = page.waitForEvent('download');
    await page.click('#btnExportScreenshot');
    await download;
    await waitForStatus(page, 'Screenshot exported');

    // Responsive sweep (§4 + U1): tablet drawer + phone bottom-sheet/More sheet
    // must be reachable and interactive without console noise.
    // Tablet (768): open the panel drawer via a tab, then dismiss via scrim.
    await page.setViewportSize({ width: 768, height: 1024 });
    await settleFrames(page);
    await page.click('#tab-explorer');
    await expect(page.locator('#btc-viewer-root')).toHaveClass(/panel-open/);
    await page.click('#scrim');
    await settleFrames(page);

    // Phone (375): bottom nav opens the tree sheet; More opens view settings.
    await page.setViewportSize({ width: 375, height: 812 });
    await settleFrames(page);
    await page.click('[data-mobile-nav="tree"]');
    await expect(page.locator('#btc-viewer-root')).toHaveClass(/panel-open/);
    await page.click('[data-mobile-nav="more"]');
    await expect(page.locator('#btc-viewer-root')).toHaveClass(/sheet-open/);
    await page.click('#btnCloseSheet');
    await page.click('#mobileFab'); // fit-to-model FAB
    await settleFrames(page);

    await page.setViewportSize({ width: 1600, height: 1000 });
    await settleFrames(page);

    expect(violations).toEqual([]);
    await context.close();
  });
});
