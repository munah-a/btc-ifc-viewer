import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './e2e',
  timeout: 12 * 60 * 1000,
  expect: {
    // SwiftShader on the 2-core CI runner is render-starved (AUDIT T11) —
    // give assertions more headroom there while keeping local feedback fast.
    timeout: process.env.CI ? 30 * 1000 : 20 * 1000,
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
    // Bound every action/navigation: Playwright's default is unlimited, which
    // turned a single non-actionable click into a full test-timeout hang (T10).
    // CI gets a longer bound: SwiftShader renders slowly enough that the
    // element-stability actionability check can legitimately exceed 15 s (T11).
    actionTimeout: process.env.CI ? 45 * 1000 : 15 * 1000,
    navigationTimeout: 30 * 1000,
    baseURL: 'http://127.0.0.1:4173/',
    headless: true,
    // Newer Chromium refuses software WebGL without --enable-unsafe-swiftshader
    // (unknown args are ignored harmlessly on older builds); --disable-gpu
    // keeps the render path consistent on the GPU-less runner (T11). CI-only
    // so local runs keep the real GPU.
    launchOptions: process.env.CI
      ? { args: ['--enable-unsafe-swiftshader', '--disable-gpu'] }
      : {},
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
    video: 'retain-on-failure',
    // 1280x720 (was 1600x1000) halves the software-rasterizer pixel cost (T11).
    viewport: { width: 1280, height: 720 },
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
  // Build + serve the production artifact with test hooks (AUDIT T4).
  webServer: {
    command: 'npm run e2e:serve',
    reuseExistingServer: !process.env.CI,
    timeout: 300 * 1000,
    url: 'http://127.0.0.1:4173/',
  },
});
