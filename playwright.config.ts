import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './e2e',
  timeout: 12 * 60 * 1000,
  expect: {
    timeout: 20 * 1000,
  },
  forbidOnly: Boolean(process.env.CI),
  fullyParallel: false,
  retries: process.env.CI ? 2 : 0,
  reporter: [
    ['list'],
    ['html', { open: 'never' }],
  ],
  workers: 1,
  use: {
    baseURL: 'http://127.0.0.1:4173/btc-ifc-viewer/',
    headless: true,
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
    video: 'retain-on-failure',
    viewport: { width: 1600, height: 1000 },
  },
  outputDir: 'test-results',
  projects: [
    {
      name: 'chromium',
      use: {
        browserName: 'chromium',
      },
    },
  ],
  webServer: {
    command: 'npm run dev:e2e',
    reuseExistingServer: !process.env.CI,
    timeout: 120 * 1000,
    url: 'http://127.0.0.1:4173/btc-ifc-viewer/',
  },
});
