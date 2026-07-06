import { expect, test, type Page } from '@playwright/test';

// W3.7 (C7): EN⇄DE language switch via the real top-bar control. No model load
// needed — the static shell + JS-rendered labels localize on their own, which
// keeps this spec fast. EN is the default (existing 22 specs assert English).

const viewerUrl = '/';
const CI = Boolean(process.env.CI);
const READY_TIMEOUT = CI ? 120_000 : 60_000;

const waitForAppReady = async (page: Page): Promise<void> => {
  await page.goto(viewerUrl);
  // init() ends with the English 'Ready - load IFC model(s)' status (EN default).
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
    undefined,
    { timeout: READY_TIMEOUT },
  );
};

test.describe('i18n — EN/DE language switch (C7)', () => {
  test('boots in English, switches to German and back via the real control', async ({ page }) => {
    await waitForAppReady(page);

    // --- English defaults ---
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await expect(page.locator('#langCode')).toHaveText('EN');
    // Static shell string (data-i18n) + JS-set label.
    await expect(page.locator('#btnUpload')).toContainText('Load IFC');
    await expect(page.locator('#viewLabel')).toHaveText('Orbit · perspective');
    await expect(page.locator('#panelTitle')).toHaveText('Explorer');
    // A localized aria/title attribute (data-i18n-attr).
    await expect(page.locator('#btnFitAll')).toHaveAttribute('aria-label', 'Fit all (F)');

    // --- Switch to German via the real toggle ---
    await page.click('#btnLangToggle');

    await expect(page.locator('html')).toHaveAttribute('lang', 'de');
    await expect(page.locator('#langCode')).toHaveText('DE');
    // Static shell re-hydrated to German.
    await expect(page.locator('#btnUpload')).toContainText('IFC laden');
    await expect(page.locator('#panelTitle')).toHaveText('Explorer'); // same word in DE
    await expect(page.locator('#btnFitAll')).toHaveAttribute('aria-label', 'Alles einpassen (F)');
    // JS-rendered label re-rendered to German.
    await expect(page.locator('#viewLabel')).toHaveText('Orbit · perspektivisch');
    // Status line reflects the switch (German).
    await expect(page.locator('#statusText')).toHaveText('Sprache: Deutsch');

    // --- Switch back to English ---
    await page.click('#btnLangToggle');

    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await expect(page.locator('#langCode')).toHaveText('EN');
    await expect(page.locator('#btnUpload')).toContainText('Load IFC');
    await expect(page.locator('#viewLabel')).toHaveText('Orbit · perspective');
    await expect(page.locator('#statusText')).toHaveText('Language: English');
  });

  test('persists the language choice across a reload', async ({ page }) => {
    await waitForAppReady(page);
    await expect(page.locator('#langCode')).toHaveText('EN');

    // Switch to German, then reload.
    await page.click('#btnLangToggle');
    await expect(page.locator('html')).toHaveAttribute('lang', 'de');

    await page.reload();
    await page.waitForFunction(
      () => (document.querySelector('#statusText')?.textContent || '').includes('Bereit'),
      undefined,
      { timeout: READY_TIMEOUT },
    );

    // Restored in German without any user action.
    await expect(page.locator('html')).toHaveAttribute('lang', 'de');
    await expect(page.locator('#langCode')).toHaveText('DE');
    await expect(page.locator('#btnUpload')).toContainText('IFC laden');

    // Reset to English so this spec leaves the persisted default clean for
    // any specs that share storage state.
    await page.click('#btnLangToggle');
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
  });
});
