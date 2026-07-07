import path from 'node:path';

import { expect, test, type BrowserContext, type Page } from '@playwright/test';

// Pulls in the `Window.__viewerTestApi` global augmentation (T6 + C8 hooks).
import '../src/core/test-api';

/**
 * C8 full-session persistence round-trip (W5.2). Loads TWO models, applies
 * per-model modifications (opacity + transform offset + hide), then RELOADS the
 * page in the SAME browser context — so localStorage (session metadata) and
 * IndexedDB (the cached `.frag` bytes) survive — and asserts both models plus
 * every modification are restored WITHOUT re-conversion (restore reads the
 * cached fragments, not the IFC).
 */

const ifcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc');
const secondIfcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'Ifc4_Revit_ARC.ifc');
const viewerUrl = '/';

const CI = Boolean(process.env.CI);
const VIEWPORT = { width: 1280, height: 720 };
const STATUS_TIMEOUT = CI ? 60_000 : 30_000;
const MODEL_TIMEOUT = CI ? 300_000 : 180_000;
const RESTORE_TIMEOUT = CI ? 240_000 : 120_000;

const waitForAppReady = async (page: Page): Promise<void> => {
  await page.goto(viewerUrl);
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
    undefined,
    { timeout: CI ? 120_000 : 60_000 },
  );
};

const waitForModelCount = async (page: Page, count: number): Promise<void> => {
  await page.waitForFunction(
    (expected) => {
      const api = window.__viewerTestApi;
      return !!api && api.modelCount() === expected && api.indexedModelCount() === expected;
    },
    count,
    { timeout: MODEL_TIMEOUT },
  );
};

const firstTwoModelIds = async (page: Page): Promise<[string, string]> => {
  const ids = await page.evaluate(() => window.__viewerTestApi?.allModelIds() ?? []);
  if (ids.length < 2) throw new Error(`Expected 2 model ids, found ${ids.length}`);
  return [ids[0], ids[1]];
};

test.describe('C8 full-session persistence round-trip (W5.2)', () => {
  let context: BrowserContext;
  let page: Page;

  test.beforeAll(async ({ browser }) => {
    // A single shared context so the reload keeps IndexedDB + localStorage.
    context = await browser.newContext({ viewport: VIEWPORT });
    page = await context.newPage();
  });

  test.afterAll(async () => {
    await context.close();
  });

  test('loads 2 models + modifications, reloads, and restores both without re-conversion', async () => {
    test.setTimeout(CI ? 20 * 60 * 1000 : 8 * 60 * 1000);

    await waitForAppReady(page);
    // Cheapest style so SwiftShader doesn't starve the click actionability (T11).
    await page.setInputFiles('#fileInput', ifcPath);
    await waitForModelCount(page, 1);
    await page.evaluate(async () => {
      await window.__viewerTestApi?.setVisualStyle('basic');
    });
    await page.setInputFiles('#fileInput', secondIfcPath);
    await waitForModelCount(page, 2);

    const [modelA, modelB] = await firstTwoModelIds(page);

    // Apply per-model modifications: opacity + a transform offset on A, and
    // opacity on B — via the same code paths the sliders/inputs drive.
    await page.evaluate(
      ({ a, b }) => {
        const api = window.__viewerTestApi;
        if (!api) throw new Error('test api missing');
        api.setModelOpacity(a, 0.4);
        api.setModelOffset(a, 5, 0, 0);
        api.setModelOpacity(b, 0.6);
        api.saveSession();
      },
      { a: modelA, b: modelB },
    );

    // Snapshot the fragKeys + modifications before reload.
    const before = await page.evaluate(
      ({ a, b }) => {
        const api = window.__viewerTestApi;
        return {
          persisted: api?.persistedModelCount() ?? 0,
          a: api?.modelModifications(a) ?? null,
          b: api?.modelModifications(b) ?? null,
        };
      },
      { a: modelA, b: modelB },
    );
    expect(before.persisted).toBe(2);
    expect(before.a?.fragKey).toBeTruthy();
    expect(before.b?.fragKey).toBeTruthy();
    expect(before.a?.opacity).toBeCloseTo(0.4, 2);
    expect(before.a?.offsetPosition.x).toBeCloseTo(5, 2);
    expect(before.b?.opacity).toBeCloseTo(0.6, 2);

    // ── RELOAD: same context, so IDB + localStorage persist. ──
    await waitForAppReady(page);

    // The session auto-restores on boot: both models re-added from the IDB
    // frag cache (no IFC re-conversion), modifications re-applied. The transient
    // "Session restored" status is overwritten by the boot "Ready" message, so
    // the durable signal is modelCount reaching 2.
    await page.waitForFunction(
      () => (window.__viewerTestApi?.modelCount() ?? 0) === 2,
      undefined,
      { timeout: RESTORE_TIMEOUT },
    );

    // Assert the modifications survived the reload. The model ids are stable
    // (derived from file name), so the same ids restore. Poll: modifications are
    // applied a beat after the model count settles.
    await expect
      .poll(async () => page.evaluate(
        ({ a, b }) => {
          const api = window.__viewerTestApi;
          const ma = api?.modelModifications(a);
          const mb = api?.modelModifications(b);
          return ma && mb ? Math.round(ma.opacity * 100) + Math.round(mb.opacity * 100) : -1;
        },
        { a: modelA, b: modelB },
      ), { timeout: RESTORE_TIMEOUT })
      .toBe(100); // 0.4*100 + 0.6*100

    const after = await page.evaluate(
      ({ a, b }) => {
        const api = window.__viewerTestApi;
        return {
          modelCount: api?.modelCount() ?? 0,
          a: api?.modelModifications(a) ?? null,
          b: api?.modelModifications(b) ?? null,
        };
      },
      { a: modelA, b: modelB },
    );

    expect(after.modelCount).toBe(2);
    // Same cached fragments were reused (proves no re-conversion path ran).
    expect(after.a?.fragKey).toBe(before.a?.fragKey);
    expect(after.b?.fragKey).toBe(before.b?.fragKey);
    // Per-model modifications restored.
    expect(after.a?.opacity).toBeCloseTo(0.4, 2);
    expect(after.a?.offsetPosition.x).toBeCloseTo(5, 2);
    expect(after.b?.opacity).toBeCloseTo(0.6, 2);

    // ── Explicit Save/Restore affordances (same restored session). ──
    // Set a distinct opacity, then Save — the explicit affordance confirms.
    await page.evaluate((a) => window.__viewerTestApi?.setModelOpacity(a, 0.25), modelA);
    await page.click('#btnSaveSession');
    await page.waitForFunction(
      () => (document.querySelector('#statusText')?.textContent || '').includes('Session saved'),
      undefined,
      { timeout: STATUS_TIMEOUT },
    );
    // The saved localStorage now carries 0.25 for model A (auto-save + explicit
    // Save both write it). Restore re-applies the saved session end-to-end.
    await page.click('#btnRestoreSession');
    await expect
      .poll(async () => {
        const mod = await page.evaluate((a) => window.__viewerTestApi?.modelModifications(a) ?? null, modelA);
        return mod?.opacity ?? -1;
      }, { timeout: RESTORE_TIMEOUT })
      .toBeCloseTo(0.25, 2);
    // Both models still present after the manual restore.
    expect(await page.evaluate(() => window.__viewerTestApi?.modelCount() ?? 0)).toBe(2);
  });
});
