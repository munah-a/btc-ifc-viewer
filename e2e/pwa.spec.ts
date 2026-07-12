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
    // The offline restore needs BOTH async post-load steps to have finished
    // before the reload, or there is nothing to restore (pre-existing flake on
    // loaded machines — modelCount flips long before them):
    // (a) persist-on-load — runs at the END of registerModel, right before the
    //     'Model loaded successfully' status, so gate on the status;
    await page.waitForFunction(
      () => (document.querySelector('#statusText')?.textContent || '').includes('Model loaded successfully'),
      undefined,
      { timeout: MODEL_TIMEOUT },
    );
    // (b) the fire-and-forget `.frag` write into the IndexedDB cache (~8MB).
    await page.waitForFunction(
      async () =>
        new Promise<boolean>((resolve) => {
          const open = indexedDB.open('btc-viewer-frag-cache');
          open.onerror = () => resolve(false);
          open.onsuccess = () => {
            const db = open.result;
            try {
              const request = db.transaction('frags', 'readonly').objectStore('frags').count();
              request.onsuccess = () => {
                resolve(request.result >= 1);
                db.close();
              };
              request.onerror = () => {
                resolve(false);
                db.close();
              };
            } catch {
              resolve(false);
              db.close();
            }
          };
        }),
      undefined,
      { timeout: MODEL_TIMEOUT },
    );
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

  test('P1 (W5-fixups): offline /embed navigation serves the CHROMELESS embed shell, not the full app', async ({
    browser,
  }) => {
    test.setTimeout(CI ? 10 * 60 * 1000 : 5 * 60 * 1000);
    const context = await browser.newContext({ viewport: { width: 1280, height: 720 } });
    const page = await context.newPage();

    // Online visit to register + activate the SW (precaches index.html AND
    // embed.html per P1). Also visit /embed once online so it's in the SW cache
    // via the navigate handler too.
    await page.goto(viewerUrl);
    await waitForAppReady(page);
    await page.waitForFunction(
      async () => {
        if (!('serviceWorker' in navigator)) return false;
        const reg = await navigator.serviceWorker.getRegistration();
        return reg?.active?.state === 'activated';
      },
      undefined,
      { timeout: 60_000 },
    );
    // Let the install precache (index.html + embed.html) settle.
    await page.waitForTimeout(1500);

    // Cut the network and navigate to the embed route. The path-aware navigate
    // fallback (matchShell -> pickShellName) must serve embed.html, so the
    // chromeless embed root is present and the full-app status bar is NOT.
    await context.setOffline(true);
    await page.goto('/embed.html');
    await expect(page.locator('#btc-embed-root')).toHaveCount(1, { timeout: 30_000 });
    expect(await page.locator('#statusText').count()).toBe(0);

    await context.setOffline(false);
    await context.close();
  });

  test('P2 (W5-fixups): a 5xx navigation does not poison the cached shell', async ({ browser }) => {
    test.setTimeout(CI ? 10 * 60 * 1000 : 5 * 60 * 1000);
    const context = await browser.newContext({ viewport: { width: 1280, height: 720 } });
    const page = await context.newPage();

    // Online visit to register + activate the SW and cache a good shell for '/'.
    await page.goto(viewerUrl);
    await waitForAppReady(page);
    await page.waitForFunction(
      async () => {
        if (!('serviceWorker' in navigator)) return false;
        const reg = await navigator.serviceWorker.getRegistration();
        return reg?.active?.state === 'activated';
      },
      undefined,
      { timeout: 60_000 },
    );
    await page.waitForTimeout(1500);

    // Force the next top-level navigation to return a 5xx. The SW navigate branch
    // must NOT cache.put this error (shouldCacheNavigation guard) — it should fall
    // back to the precached good shell instead of poisoning it.
    await page.route('**/', (route, req) => {
      if (req.resourceType() === 'document') {
        return route.fulfill({ status: 503, contentType: 'text/html', body: 'upstream down' });
      }
      return route.continue();
    });
    await page.goto(viewerUrl);
    // The SW served the precached good shell despite the 503 — the app boots.
    await waitForAppReady(page);
    await page.unroute('**/');

    // Now go fully offline and reload: the cached shell must STILL be the good one
    // (the 503 never overwrote it). App reaches Ready with no network.
    await context.setOffline(true);
    await page.reload();
    await waitForAppReady(page);
    expect(await page.locator('#statusText').count()).toBe(1);

    await context.setOffline(false);
    await context.close();
  });
});
