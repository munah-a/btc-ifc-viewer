import fs from 'node:fs';
import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

// W4.7: e2e for the chromeless /embed entry and the share/GLB flow. Uses route
// interception so nothing hits live Vercel storage — the model URL and the API
// are all fulfilled locally. The 24 prior specs are unaffected (separate files).

const CI = Boolean(process.env.CI);
// Converting even the smaller fixture IFC in headless SwiftShader is slow on the
// 2-core CI runner (AUDIT T11) — the embed load path does a full IFC→fragments
// conversion, so budget generously there.
const CONVERT_TIMEOUT = CI ? 240_000 : 120_000;
const READY_TIMEOUT = CI ? 120_000 : 60_000;

const smallIfc = fs.readFileSync(path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc'));
// An ALLOWED (Vercel Blob CDN) model URL — passes the S8 allowlist; intercepted
// below so no real network hit occurs. (A non-allowed host is exercised
// separately by the S8 "blocked" test.)
const MODEL_URL = 'https://testblob.public.blob.vercel-storage.com/frags/sample.ifc';
// A disallowed foreign host for the S8 rejection test.
const FOREIGN_URL = 'https://evil.example.com/frags/sample.ifc';

/** Fulfills the model fetch with the fixture IFC bytes (embed converts client-side). */
async function routeModel(page: Page): Promise<void> {
  await page.route(MODEL_URL, (route) =>
    route.fulfill({
      status: 200,
      headers: { 'content-type': 'application/octet-stream', 'access-control-allow-origin': '*' },
      body: smallIfc,
    }),
  );
}

function collectConsoleErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('console', (msg) => {
    if (msg.type() === 'error') errors.push(msg.text());
  });
  page.on('pageerror', (err) => errors.push(err.message));
  return errors;
}

test.describe('embed · load by URL', () => {
  test('poster → activate → loads the model → shows controls, console clean', async ({ page }) => {
    test.setTimeout(CONVERT_TIMEOUT + 60_000);
    const errors = collectConsoleErrors(page);
    await routeModel(page);

    await page.goto(`/embed.html?m=${encodeURIComponent(MODEL_URL)}`);

    // Poster shows (WebGL not created yet).
    const poster = page.locator('#embedPoster');
    await expect(poster).toBeVisible();

    // Activate → engine bootstraps + model converts.
    await poster.click();

    // Loading overlay appears, then the model finishes and controls reveal.
    await page.waitForFunction(() => !document.getElementById('embedControls')?.hidden, undefined, {
      timeout: CONVERT_TIMEOUT,
    });

    // Canvas exists and the error state is NOT shown.
    await expect(page.locator('#embed-viewer canvas')).toBeAttached();
    await expect(page.locator('#embedError')).toBeHidden();

    // The BTC badge deep-links back to the full app with the same model URL.
    const badgeHref = await page.locator('#embedBadge').getAttribute('href');
    expect(badgeHref).toContain(encodeURIComponent(MODEL_URL));

    // Fit button works without error.
    await page.locator('#embedFit').click();

    // No console errors across the whole flow (§4).
    expect(errors, `console errors: ${errors.join(' | ')}`).toEqual([]);
  });

  test('missing ?m= shows the friendly no-model state (no engine spun up)', async ({ page }) => {
    await page.goto('/embed.html');
    await expect(page.locator('#embedError')).toBeVisible();
    await expect(page.locator('#embedPoster')).toBeHidden();
    // No canvas — the engine was never created.
    await expect(page.locator('#embed-viewer canvas')).toHaveCount(0);
  });

  test('expired / 404 model shows the expired error state with an open-in-viewer link', async ({ page }) => {
    await page.route(MODEL_URL, (route) => route.fulfill({ status: 404, body: 'gone' }));
    await page.goto(`/embed.html?m=${encodeURIComponent(MODEL_URL)}`);
    await page.locator('#embedPoster').click();
    await expect(page.locator('#embedError')).toBeVisible({ timeout: READY_TIMEOUT });
    const openHref = await page.locator('#embedErrorOpen').getAttribute('href');
    expect(openHref).toContain(encodeURIComponent(MODEL_URL));
  });

  test('S8: a disallowed foreign ?m= host is blocked (no poster, no engine, no fetch)', async ({ page }) => {
    let fetched = false;
    await page.route(FOREIGN_URL, (route) => {
      fetched = true;
      return route.fulfill({ status: 200, body: smallIfc });
    });
    await page.goto(`/embed.html?m=${encodeURIComponent(FOREIGN_URL)}`);
    // Blocked up front: error shown, poster hidden, engine never created.
    await expect(page.locator('#embedError')).toBeVisible();
    await expect(page.locator('#embedPoster')).toBeHidden();
    await expect(page.locator('#embed-viewer canvas')).toHaveCount(0);
    // The foreign URL was never fetched.
    expect(fetched).toBe(false);
  });
});
