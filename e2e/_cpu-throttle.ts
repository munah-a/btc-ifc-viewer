import { type Page } from '@playwright/test';

/**
 * Emulate a slow CPU so a fast dev box better reproduces GitHub's 2-core runner
 * timing (AUDIT T12: a green local e2e on GPU hardware does NOT prove CI wall
 * clock — timeout-bound tests can still fail on the slower runner). Opt-in via
 * `E2E_CPU_THROTTLE=<rate>` (e.g. 3). No-op when unset, so normal local runs
 * stay fast. Chromium-only (the suite's single project).
 *
 * LIMITATION (measured T12): CDP throttling slows the page main thread only —
 * it does NOT throttle SwiftShader's software-GL threads or the fragments
 * worker, which are the actual bottleneck on the runner. So render-bound flows
 * (the console-clean sweep) barely slow down here; their CI budget must instead
 * be set with a wide margin (~3x local). Throttling still catches CPU-bound
 * (parse/layout/JS) slowness.
 *
 * Usage in a fixture, right after the page is created:
 *   await applyCpuThrottle(page);
 */
export async function applyCpuThrottle(page: Page): Promise<void> {
  const rate = Number(process.env.E2E_CPU_THROTTLE ?? '0');
  if (!Number.isFinite(rate) || rate <= 1) return;
  const client = await page.context().newCDPSession(page);
  await client.send('Emulation.setCPUThrottlingRate', { rate });
}
