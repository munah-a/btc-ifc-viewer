import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

import '../src/core/test-api';

/**
 * PWA offline (W5.5 / C6). Verifies the service worker precaches the app shell
 * and self-hosted runtime assets, then — with the network cut — the app STILL
 * boots and a previously-cached model restores from IndexedDB (offline field use
 * on a tablet). C1: the SW never depends on a CDN.
 */

const ifcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc');
const viewerUrl = '/';
const CI = Boolean(process.env.CI);
const MODEL_TIMEOUT = CI ? 300_000 : 180_000;
const RESTORE_TIMEOUT = CI ? 240_000 : 120_000;

const waitForAppReady = async (page: Page): Promise<void> => {
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
    undefined,
    { timeout: CI ? 120_000 : 60_000 },
  );
};

test.describe('PWA offline (W5.5 / C6)', () => {
  test('service worker activates, then the app boots + restores a model OFFLINE', async ({ browser }) => {
    test.setTimeout(CI ? 20 * 60 * 1000 : 8 * 60 * 1000);
    const context = await browser.newContext({ viewport: { width: 1280, height: 720 } });
    const page = await context.newPage();

    // 1. First online visit — register + activate the SW, load a model (populates
    //    the SW shell cache + the IndexedDB frag cache).
    await page.goto(viewerUrl);
    await waitForAppReady(page);

    // Wait for the SW to reach 'activated' (registration happens on window load).
    // Controlling the current page can lag behind activation, so gate on the
    // registration's active worker state, not navigator.controller.
    await page.waitForFunction(
      async () => {
        if (!('serviceWorker' in navigator)) return false;
        const reg = await navigator.serviceWorker.getRegistration();
        return reg?.active?.state === 'activated';
      },
      undefined,
      { timeout: 60_000 },
    );

    await page.setInputFiles('#fileInput', ifcPath);
    await page.waitForFunction(() => (window.__viewerTestApi?.modelCount() ?? 0) === 1, undefined, {
      timeout: MODEL_TIMEOUT,
    });
    // Give the SW a beat to finish caching the shell assets it fetched.
    await page.waitForTimeout(1500);

    // 2. Go OFFLINE and reload. The shell must boot from the SW cache and the
    //    model must restore from IndexedDB — with zero network available.
    await context.setOffline(true);
    await page.reload();
    await waitForAppReady(page);

    // App shell booted offline (status reached Ready).
    // The cached model auto-restores from IndexedDB (no network / no re-fetch).
    await page.waitForFunction(() => (window.__viewerTestApi?.modelCount() ?? 0) === 1, undefined, {
      timeout: RESTORE_TIMEOUT,
    });

    expect(await page.evaluate(() => window.__viewerTestApi?.modelCount() ?? 0)).toBe(1);

    await context.setOffline(false);
    await context.close();
  });
});
