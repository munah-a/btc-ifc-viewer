import fs from 'node:fs';
import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

/**
 * Rebrand visual-fidelity capture (program acceptance criterion 4). Not part of
 * the gate — runs only with CAPTURE=1. Shoots the rebranded shell across
 * viewport × theme × state so the PO/user can confirm it matches the Claude
 * Design mockup. page.screenshot() captures the compositor (real SwiftShader
 * render), so the 3D model appears even headless.
 */
const ifcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc');
const outDir = path.join(process.cwd(), 'docs', 'design', 'shots');

const shot = async (page: Page, name: string): Promise<void> => {
  fs.mkdirSync(outDir, { recursive: true });
  await page.screenshot({ path: path.join(outDir, `${name}.png`) });
};

const boot = async (page: Page): Promise<void> => {
  await page.goto('/');
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
    undefined,
    { timeout: 60_000 },
  );
};

const load = async (page: Page): Promise<void> => {
  await page.setInputFiles('#fileInput', ifcPath);
  await page.waitForFunction(
    () => {
      const a = (window as unknown as { __viewerTestApi?: { modelCount(): number; indexedModelCount(): number } }).__viewerTestApi;
      return !!a && a.modelCount() === 1 && a.indexedModelCount() === 1;
    },
    undefined,
    { timeout: 300_000 },
  );
  // Let a few frames render so the model is visible in the shot.
  await page.evaluate(() => new Promise((r) => setTimeout(r, 2500)));
};

const setTheme = async (page: Page, theme: 'dark' | 'light'): Promise<void> => {
  const current = await page.evaluate(() => document.documentElement.getAttribute('data-theme'));
  if (current === theme) return;
  // Prefer the real toggle (also swaps canvas bg); fall back to the attribute.
  const toggle = page.locator('#toggleTheme, [data-action="theme"], [aria-label*="theme" i]').first();
  if (await toggle.count()) {
    await toggle.click().catch(() => undefined);
  }
  const after = await page.evaluate(() => document.documentElement.getAttribute('data-theme'));
  if (after !== theme) {
    await page.evaluate((t) => document.documentElement.setAttribute('data-theme', t), theme);
  }
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
};

test('capture rebrand screenshots', async ({ browser }) => {
  test.skip(!process.env.CAPTURE, 'Capture scaffold — run with CAPTURE=1 only.');
  test.setTimeout(10 * 60 * 1000);

  // ---- Desktop 1440x900 ----
  const desktop = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const dp = await desktop.newPage();
  await boot(dp);
  await setTheme(dp, 'dark');
  await shot(dp, 'desktop-01-empty-dark');
  await setTheme(dp, 'light');
  await shot(dp, 'desktop-02-empty-light');
  await setTheme(dp, 'dark');
  await load(dp);
  await shot(dp, 'desktop-03-loaded-explorer-dark');
  // Properties tab with a selection
  await dp.evaluate(async () => {
    const a = (window as unknown as { __viewerTestApi?: { selectFirstItemPerModel(): Promise<void> } }).__viewerTestApi;
    await a?.selectFirstItemPerModel();
  });
  await dp.locator('[data-tab="properties"], #tab-properties, [role="tab"]:has-text("Properties")').first().click().catch(() => undefined);
  await dp.evaluate(() => new Promise((r) => setTimeout(r, 600)));
  await shot(dp, 'desktop-04-loaded-properties-dark');
  await setTheme(dp, 'light');
  await shot(dp, 'desktop-05-loaded-light');
  await desktop.close();

  // ---- Tablet 768x1024 ----
  const tablet = await browser.newContext({ viewport: { width: 768, height: 1024 } });
  const tp = await tablet.newPage();
  await boot(tp);
  await shot(tp, 'tablet-01-empty-dark');
  await load(tp);
  await shot(tp, 'tablet-02-loaded-dark');
  await tablet.close();

  // ---- Phone 390x844 ----
  const phone = await browser.newContext({ viewport: { width: 390, height: 844 } });
  const pp = await phone.newPage();
  await boot(pp);
  await shot(pp, 'phone-01-empty-dark');
  await load(pp);
  await shot(pp, 'phone-02-loaded-dark');
  // Open the mobile bottom sheet if a nav exists
  await pp.locator('[data-mobile-nav] button, .mobile-nav button, .bottom-nav button').first().click().catch(() => undefined);
  await pp.evaluate(() => new Promise((r) => setTimeout(r, 500)));
  await shot(pp, 'phone-03-sheet-dark');
  await phone.close();

  expect(fs.readdirSync(outDir).filter((f) => f.endsWith('.png')).length).toBeGreaterThanOrEqual(10);
});
