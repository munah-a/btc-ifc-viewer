import fs from 'node:fs';
import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

import '../src/core/test-api';

// W4.7: share-dialog flow (mocked upload/delete) + GLB export, in the full app.
// The hosting API is intercepted so nothing hits live Vercel storage.

const CI = Boolean(process.env.CI);
const READY_TIMEOUT = CI ? 120_000 : 60_000;
const CONVERT_TIMEOUT = CI ? 300_000 : 180_000;

const smallIfc = fs.readFileSync(path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc'));

const MOCK_PUBLISH = {
  id: 'testid123456',
  embedUrl: 'https://btc-ifc-viewer-2.vercel.app/embed.html?m=https%3A%2F%2Fcdn%2Fx.frag&id=testid123456',
  viewerUrl: 'https://btc-ifc-viewer-2.vercel.app/?m=https%3A%2F%2Fcdn%2Fx.frag',
  fragUrl: 'https://cdn.test.local/frags/testid123456.frag',
  deleteToken: 'delete-token-xyz',
  expiresAt: '2026-07-13T03:00:00.000Z',
};

async function mockApi(page: Page): Promise<void> {
  await page.route('**/api/uploads', (route) =>
    route.fulfill({
      status: 201,
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(MOCK_PUBLISH),
    }),
  );
  await page.route('**/api/e/**', (route) => {
    if (route.request().method() === 'DELETE') {
      return route.fulfill({
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ id: MOCK_PUBLISH.id, deleted: true }),
      });
    }
    return route.continue();
  });
}

async function loadModel(page: Page): Promise<void> {
  await page.goto('/');
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
    undefined,
    { timeout: READY_TIMEOUT },
  );
  await page.evaluate((bytes) => {
    const file = new File([new Uint8Array(bytes)], 'school_str.ifc', { type: 'application/octet-stream' });
    const input = document.getElementById('fileInput') as HTMLInputElement;
    const dt = new DataTransfer();
    dt.items.add(file);
    input.files = dt.files;
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }, Array.from(smallIfc));
  await page.waitForFunction(() => (window.__viewerTestApi?.modelCount() ?? 0) >= 1, undefined, {
    timeout: CONVERT_TIMEOUT,
  });
}

test.describe('share dialog', () => {
  test('publish → shows embed link, iframe, QR, expiry; delete clears it (mocked API)', async ({ page }) => {
    test.setTimeout(CONVERT_TIMEOUT + 60_000);
    await mockApi(page);
    await loadModel(page);

    // Open the share dialog.
    await page.locator('#btnShare').click();
    await expect(page.locator('#shareDialog')).toBeVisible();

    // Publish.
    await page.locator('#sharePublish').click();
    await expect(page.locator('#sharePublished')).toBeVisible({ timeout: 30_000 });

    // Embed link + iframe snippet populated with the mocked embed URL.
    await expect(page.locator('#shareEmbedUrl')).toHaveValue(MOCK_PUBLISH.embedUrl);
    const iframe = await page.locator('#shareIframe').inputValue();
    expect(iframe).toContain('<iframe');
    expect(iframe).toContain(MOCK_PUBLISH.embedUrl.replace(/&/g, '&amp;').split('&amp;')[0]);

    // QR canvas rendered (non-blank: has some dark pixels).
    const qrHasContent = await page.evaluate(() => {
      const c = document.getElementById('shareQr') as HTMLCanvasElement;
      const ctx = c.getContext('2d')!;
      const data = ctx.getImageData(0, 0, c.width, c.height).data;
      let dark = 0;
      for (let i = 0; i < data.length; i += 4) if (data[i] < 128) dark += 1;
      return dark > 50;
    });
    expect(qrHasContent).toBe(true);

    // Expiry shown.
    await expect(page.locator('#shareExpiry')).toContainText('13.07.2026');

    // Delete (mocked): confirm dialog → confirm → result hides.
    await page.locator('#shareDelete').click();
    // The app's native confirm dialog appears; click its confirm button.
    await page.locator('#confirmOk').click();
    await expect(page.locator('#sharePublished')).toBeHidden({ timeout: 15_000 });
  });

  test('PowerPoint tab shows the how-to steps and GLB button', async ({ page }) => {
    test.setTimeout(READY_TIMEOUT + 30_000);
    await page.goto('/');
    await page.waitForFunction(
      () => (document.querySelector('#statusText')?.textContent || '').includes('Ready - load IFC'),
      undefined,
      { timeout: READY_TIMEOUT },
    );
    await page.locator('#btnShare').click();
    await page.locator('#shareTabPp').click();
    await expect(page.locator('#sharePanelPp')).toBeVisible();
    await expect(page.locator('.share-steps li')).toHaveCount(2);
    await expect(page.locator('#sharePpGlb')).toBeVisible();
  });
});

test.describe('GLB export', () => {
  test('exports a non-empty, valid binary GLB from a loaded model', async ({ page }) => {
    test.setTimeout(CONVERT_TIMEOUT + 60_000);
    await loadModel(page);
    const result = await page.evaluate(() => window.__viewerTestApi!.exportGlbBytes());
    expect(result.valid).toBe(true);
    expect(result.byteLength).toBeGreaterThan(1000);
  });
});
