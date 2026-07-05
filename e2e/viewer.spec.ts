import fs from 'node:fs';
import path from 'node:path';

import { expect, test as base, type Page } from '@playwright/test';

type CameraPosition = {
  x: number;
  y: number;
  z: number;
};

type CameraState = {
  position: CameraPosition;
  target: CameraPosition;
};

type ModelContext = {
  modelId: string;
  searchTerm: string;
  firstItemId: number;
};

const ifcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'school_str.ifc');
const secondIfcPath = path.join(process.cwd(), 'e2e', 'fixtures', 'Ifc4_Revit_ARC.ifc');
// Resolved against playwright.config.ts baseURL (production preview at '/').
const viewerUrl = '/';

const waitForAppReady = async (page: Page): Promise<void> => {
  await page.goto(viewerUrl);
  // init() ends with the 'Ready - load IFC model(s)' status; nothing races it
  // anymore since W1.5 deleted the F4 grid hack.
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready'),
    undefined,
    { timeout: 60_000 },
  );
};

// A model is fully registered when its index exists (W1.4: indexing is the
// last await inside registration; the old registeringModelIds set is gone).
const waitForModelCount = async (page: Page, count: number): Promise<void> => {
  await page.waitForFunction(
    (expected) => {
      const viewer = (window as any).__viewer;
      const elementText = document.querySelector('#elementCount')?.textContent || '';
      return !!viewer
        && viewer.federatedModels?.size === expected
        && viewer.modelIndices?.size === expected
        && elementText !== '0 elements';
    },
    count,
    { timeout: 180_000 },
  );
};

const waitForModelReady = async (page: Page): Promise<void> => {
  await waitForModelCount(page, 1);
};

const waitForStatus = async (page: Page, text: string, timeout = 20_000): Promise<void> => {
  await page.waitForFunction(
    (expected) => (document.querySelector('#statusText')?.textContent || '').includes(expected),
    text,
    { timeout },
  );
};

// Two rAF ticks — lets the layout/render pipeline settle after viewport changes.
const waitForLayoutSettle = async (page: Page): Promise<void> => {
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
};

// NOTE (AUDIT T10): the old monolith's `openDock` helper clicked
// `[data-dock-toggle][title=...]`, but since the ui-overhaul merge those
// toggles are display:none at desktop widths (all dock tools render inline in
// the toolbar), so the click could never become actionable and hung for the
// whole test timeout. Tool buttons are now clicked directly.

const normalizeVector = (vector: CameraPosition): CameraPosition => {
  const length = Math.hypot(vector.x, vector.y, vector.z) || 1;
  return {
    x: vector.x / length,
    y: vector.y / length,
    z: vector.z / length,
  };
};

const dotProduct = (a: CameraPosition, b: CameraPosition): number => (
  a.x * b.x + a.y * b.y + a.z * b.z
);

const getCameraState = async (page: Page): Promise<CameraState> => page.evaluate(() => {
  const camera = (window as any).__world.camera.three;
  const target = camera.position.clone();
  (window as any).__world.camera.controls.getTarget(target);
  return {
    position: { x: camera.position.x, y: camera.position.y, z: camera.position.z },
    target: { x: target.x, y: target.y, z: target.z },
  };
});

const getCameraPosition = async (page: Page): Promise<CameraPosition> => page.evaluate(() => {
  const position = (window as any).__world.camera.three.position;
  return { x: position.x, y: position.y, z: position.z };
});

const getCameraDirection = async (page: Page): Promise<CameraPosition> => {
  const state = await getCameraState(page);
  return normalizeVector({
    x: state.position.x - state.target.x,
    y: state.position.y - state.target.y,
    z: state.position.z - state.target.z,
  });
};

const cameraChanged = (before: CameraPosition, after: CameraPosition): boolean => (
  Math.abs(before.x - after.x) > 0.01
  || Math.abs(before.y - after.y) > 0.01
  || Math.abs(before.z - after.z) > 0.01
);

// State wait (T5): polls until the camera animation settles on the expected
// direction — replaces the fixed 750 ms sleeps of the old monolith.
const waitForCameraDirection = async (
  page: Page,
  expected: CameraPosition,
  threshold = 0.985,
): Promise<void> => {
  await expect
    .poll(
      async () => dotProduct(await getCameraDirection(page), normalizeVector(expected)),
      { timeout: 15_000 },
    )
    .toBeGreaterThanOrEqual(threshold);
};

const waitForCameraMove = async (page: Page, before: CameraPosition): Promise<void> => {
  await expect
    .poll(async () => cameraChanged(before, await getCameraPosition(page)), { timeout: 15_000 })
    .toBeTruthy();
};

const clickVisibleCubeTarget = async (page: Page, key: string): Promise<void> => {
  const target = page.locator(`[data-cube-target="${key}"][data-visible="true"]`).first();
  await expect(target).toBeVisible();
  await target.click();
};

const getExpectedCubeDirection = async (
  page: Page,
  vector: readonly [number, number, number],
): Promise<CameraPosition> => {
  const expected = await page.evaluate((localVector) => {
    const viewer = (window as any).__viewer;
    const models = viewer?.federatedModels ? Array.from(viewer.federatedModels.values()) : [];
    const anchor = viewer?.activeGizmoModelId
      ? viewer.federatedModels.get(viewer.activeGizmoModelId)
      : models.find((model: any) => model.visible) || models[0];

    if (!anchor) return null;

    const basis = anchor.object.getWorldQuaternion(anchor.object.quaternion.clone());
    const direction = anchor.object.position.clone();
    direction.set(localVector[0], localVector[1], localVector[2]).normalize().applyQuaternion(basis).normalize();
    return { x: direction.x, y: direction.y, z: direction.z };
  }, vector);

  if (!expected) throw new Error('Expected cube direction unavailable');
  return expected;
};

const getModelContext = async (page: Page): Promise<ModelContext> => {
  const context = await page.evaluate(() => {
    const viewer = (window as any).__viewer;
    const firstEntry = viewer?.modelIndices ? Array.from(viewer.modelIndices.entries())[0] : null;
    if (!firstEntry) return null;

    const [modelId, index] = firstEntry as [string, any];
    let firstNamed: { id: number; name: string } | null = null;

    for (const [id, name] of index.itemNames.entries()) {
      if (!index.allIds.has(id) || !name) continue;
      firstNamed = { id, name };
      break;
    }

    return {
      modelId,
      searchTerm: firstNamed?.name || Array.from(index.classes.keys())[0] || '',
      firstItemId: firstNamed?.id || Array.from(index.allIds)[0],
    };
  });

  if (!context) throw new Error('Model context unavailable');
  return context as ModelContext;
};

const findItemByNameKeyword = async (
  page: Page,
  keyword: string,
): Promise<{ modelId: string; localId: number } | null> => page.evaluate((needle) => {
  const viewer = (window as any).__viewer;
  if (!viewer?.modelIndices) return null;

  for (const [modelId, index] of viewer.modelIndices.entries()) {
    for (const [localId, name] of index.itemNames.entries()) {
      if (!index.allIds.has(localId) || !name) continue;
      if (name.toLowerCase().includes(needle)) {
        return { modelId, localId };
      }
    }
  }

  return null;
}, keyword.toLowerCase());

const ensureSingleSelection = async (page: Page): Promise<void> => {
  const context = await getModelContext(page);
  await page.evaluate(
    async ({ modelId, localId }) => {
      await (window as any).__viewer.selectSingleItem(modelId, localId, false);
    },
    { modelId: context.modelId, localId: context.firstItemId },
  );
  await page.waitForFunction(
    () => (document.querySelector('#selectionCount')?.textContent || '').startsWith('1 selected'),
    undefined,
    { timeout: 20_000 },
  );
};

// Shared loaded-model fixture (T5): the school_str.ifc conversion (~1 min) runs
// once per worker; every describe block below reuses the same page. Tests run
// sequentially (workers=1, fullyParallel=false) in file order. On CI retry a
// fresh worker rebuilds the fixture, so each test establishes its own
// selection/panel preconditions.
type WorkerFixtures = {
  appPage: Page;
};

// NonNullable<unknown> === the empty fixture set ({}), spelled so the
// no-empty-object-type lint rule stays happy.
const test = base.extend<NonNullable<unknown>, WorkerFixtures>({
  appPage: [
    async ({ browser }, use) => {
      const context = await browser.newContext({
        acceptDownloads: true,
        viewport: { width: 1600, height: 1000 },
      });
      const page = await context.newPage();
      await waitForAppReady(page);
      await page.setInputFiles('#fileInput', ifcPath);
      await waitForModelReady(page);
      await use(page);
      await context.close();
    },
    { scope: 'worker' },
  ],
});

test.describe('shell & boot', () => {
  test('boots to ready state with the empty-state upload affordance', async ({ page }) => {
    await waitForAppReady(page);

    const emptyUploadChooser = page.waitForEvent('filechooser');
    await page.click('#btnUploadEmpty');
    await emptyUploadChooser;

    await expect(page.locator('.header-center')).toHaveCount(0);
  });

  test('loads the model, exports a screenshot and cycles panel tabs', async ({ appPage: page }, testInfo) => {
    const headerUploadChooser = page.waitForEvent('filechooser');
    await page.click('#btnUpload');
    await headerUploadChooser;

    await expect(page.locator('#elementCount')).not.toHaveText(/0 elements/);
    await expect(page.locator('#visibleCount')).not.toHaveText(/0 visible/);
    // AUDIT A15: load metrics live in their own status slot and coexist with
    // the FPS monitor instead of being overwritten within a second.
    await expect(page.locator('#loadInfo')).toHaveText(/Loaded in .+s \| .+MB/);
    await expect(page.locator('#perfInfo')).toHaveText(/\d+ FPS/);

    const screenshotDownload = page.waitForEvent('download');
    await page.click('#btnExportScreenshot');
    expect((await screenshotDownload).suggestedFilename()).toMatch(/\.png$/);

    for (const tab of ['explorer', 'models', 'properties', 'viewpoints', 'issues', 'help']) {
      await page.click(`.tab-btn[data-tab="${tab}"]`);
      await expect(page.locator(`#panel-${tab}`)).toHaveClass(/active/);
    }
    await page.click('.tab-btn[data-tab="explorer"]');
    // Viewport screenshot: element screenshots wait for two stable rAF frames,
    // which starves on the ~2 FPS headless WebGL page (P6). These artifacts
    // are never pixel-compared (T5) — the viewport capture is enough.
    await page.screenshot({ path: testInfo.outputPath('view-cube-home.png') });
  });
});

test.describe('camera & navigation', () => {
  test('switches navigation modes', async ({ appPage: page }) => {
    await page.click('#btnModePlan');
    await expect(page.locator('#btnModePlan')).toHaveClass(/active/);

    await page.click('#btnModeFirstPerson');
    await expect(page.locator('#btnModeFirstPerson')).toHaveClass(/active/);

    await page.click('#btnModeOrbit');
    await expect(page.locator('#btnModeOrbit')).toHaveClass(/active/);
  });

  test('view cube and preset views drive the camera', async ({ appPage: page }) => {
    const beforeFront = await getCameraPosition(page);
    const expectedFrontDirection = await getExpectedCubeDirection(page, [0, 0, 1]);
    await clickVisibleCubeTarget(page, 'front');
    await waitForCameraDirection(page, expectedFrontDirection);
    expect(cameraChanged(beforeFront, await getCameraPosition(page))).toBeTruthy();

    const expectedCornerDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
    await clickVisibleCubeTarget(page, 'top-front-right');
    await waitForCameraDirection(page, expectedCornerDirection, 0.975);

    const expectedHomeDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
    await page.click('#cubeHome');
    await waitForCameraDirection(page, expectedHomeDirection, 0.975);

    const beforeTop = await getCameraPosition(page);
    await page.click('#btnTop');
    await waitForCameraMove(page, beforeTop);
  });
});

test.describe('selection & search', () => {
  test('toggles selection modes', async ({ appPage: page }) => {
    await page.click('#btnSelectMulti');
    await expect(page.locator('#btnSelectMulti')).toHaveClass(/active/);

    await page.click('#btnSelectSingle');
    await expect(page.locator('#btnSelectSingle')).toHaveClass(/active/);
  });

  // AUDIT F1 regression: search used to crash on any hit (raw ItemAttribute
  // objects fed into escapeHtml) — fixed in W1.1 by unwrapping to primitives.
  test('search finds elements and selects from results', async ({ appPage: page }) => {
    const modelContext = await getModelContext(page);
    await page.fill('#searchInput', modelContext.searchTerm.split(':')[0].trim());
    await page.click('#btnSearch');
    await expect(page.locator('.result-item').first()).toBeVisible();
    await page.locator('.result-item').first().click();
    await expect(page.locator('#selectionCount')).toHaveText(/1 selected/);
  });

  // AUDIT F11: a disjoint class∩level filter combination used to silently
  // hide the entire model — it must warn and leave visibility untouched.
  test('disjoint class and level filters warn instead of hiding everything', async ({ appPage: page }) => {
    const disjoint = await page.evaluate(() => {
      const viewer = (window as any).__viewer;
      const firstEntry = Array.from(viewer.modelIndices.values())[0] as any;
      if (!firstEntry) return null;
      for (const [className, classIds] of firstEntry.classes.entries()) {
        for (const [levelName, levelIds] of firstEntry.levels.entries()) {
          let overlaps = false;
          for (const id of classIds) {
            if (levelIds.has(id)) {
              overlaps = true;
              break;
            }
          }
          if (!overlaps) return { className, levelName };
        }
      }
      return null;
    });
    // Hard expectation (gate forbids skipped tests): the structural fixture
    // has storey-specific classes, so a disjoint pair must exist.
    expect(disjoint).not.toBeNull();
    if (!disjoint) return;

    const visibleBefore = await page.locator('#visibleCount').textContent();
    await page.check(`input[data-filter-type="class"][value="${disjoint.className}"]`);
    await page.check(`input[data-filter-type="level"][value="${disjoint.levelName}"]`);
    await page.click('#btnApplyFilters');
    await expect(page.locator('.toast-warning')).toBeVisible();
    await waitForStatus(page, 'Selected filters have no elements in common');
    await expect(page.locator('#visibleCount')).toHaveText(visibleBefore || '');

    await page.click('#btnClearFilters');
    await waitForStatus(page, 'Filters reset');
  });
});

test.describe('properties panel', () => {
  test('shows, filters and scrolls properties for selections', async ({ appPage: page }, testInfo) => {
    await ensureSingleSelection(page);
    await page.click('.tab-btn[data-tab="properties"]');
    await expect(page.locator('#propsContent')).toBeVisible();
    await expect(page.locator('#propName')).not.toHaveText('-');
    expect(await page.locator('.prop-section').count()).toBeGreaterThan(3);
    await expect(page.locator('.prop-section').first()).toHaveAttribute('open', '');
    await page.locator('.prop-section-summary').first().click();
    await expect(page.locator('.prop-section').first()).not.toHaveAttribute('open', '');
    await page.locator('.prop-section-summary').first().click();
    await expect(page.locator('.prop-section').first()).toHaveAttribute('open', '');

    const thicknessItem = await findItemByNameKeyword(page, 'floor');
    if (!thicknessItem) throw new Error('No floor-like element found in the IFC test model');
    await page.evaluate(
      async ({ modelId, localId }) => {
        await (window as any).__viewer.selectSingleItem(modelId, localId, false);
      },
      thicknessItem,
    );
    await page.click('.tab-btn[data-tab="properties"]');
    await expect(page.locator('#propsContent')).toBeVisible();
    await page.fill('#propFilterInput', 'thickness');
    await expect(page.locator('#panel-properties')).toContainText(/Thickness/i);
    // Viewport screenshot — see the stability note in the shell describe block.
    await page.screenshot({ path: testInfo.outputPath('properties-panel.png') });
    await page.fill('#propFilterInput', 'center x');
    await expect(page.locator('#panel-properties')).toContainText(/Center X/i);
    await page.fill('#propFilterInput', '');

    const scrollStressItem = await findItemByNameKeyword(page, 'tapered')
      || await findItemByNameKeyword(page, 'insulation')
      || await findItemByNameKeyword(page, 'deck')
      || thicknessItem;
    await page.evaluate(
      async ({ modelId, localId }) => {
        await (window as any).__viewer.selectSingleItem(modelId, localId, false);
      },
      scrollStressItem,
    );
    await page.click('.tab-btn[data-tab="properties"]');
    await expect(page.locator('#propsContent')).toBeVisible();
    await expect(page.locator('#panel-properties')).toContainText(/Materials/i);
    await expect(page.locator('#panel-properties')).toContainText(/Raw IFC/i);
    await page.waitForFunction(
      () => document.querySelectorAll('#panel-properties [data-prop-row]:not([hidden])').length > 20,
      undefined,
      { timeout: 20_000 },
    );

    await page.setViewportSize({ width: 1600, height: 500 });
    await waitForLayoutSettle(page);
    await page.evaluate(() => {
      const closedSections = Array.from(document.querySelectorAll<HTMLElement>('.prop-section:not([open]) .prop-section-summary'));
      for (const summary of closedSections) summary.click();
    });
    await page.waitForFunction(
      () => document.querySelectorAll('#panel-properties [data-prop-row]:not([hidden])').length > 40,
      undefined,
      { timeout: 20_000 },
    );
    const propertiesScrollState = await page.evaluate(() => {
      const sections = document.querySelector('#propSections');
      const panel = document.querySelector('#panel-properties');
      if (!sections) return null;
      return {
        panelOverflowY: panel ? getComputedStyle(panel).overflowY : null,
        sectionsOverflowY: getComputedStyle(sections).overflowY,
        sectionCount: document.querySelectorAll('.prop-section').length,
        openSectionCount: document.querySelectorAll('.prop-section[open]').length,
      };
    });
    expect(propertiesScrollState).not.toBeNull();
    expect(propertiesScrollState?.panelOverflowY).toBe('hidden');
    expect(propertiesScrollState?.sectionsOverflowY).toBe('auto');
    expect(propertiesScrollState?.openSectionCount).toBe(propertiesScrollState?.sectionCount);
    await expect(page.locator('#viewerDock')).toBeVisible();
    const dockBounds = await page.locator('#viewerDock').boundingBox();
    const viewport = page.viewportSize();
    expect(dockBounds).not.toBeNull();
    expect(viewport).not.toBeNull();
    if (!dockBounds || !viewport) {
      throw new Error('Viewer dock bounds unavailable');
    }
    expect(dockBounds.y + dockBounds.height).toBeLessThanOrEqual(viewport.height);

    // Restore the shared fixture page for subsequent describe blocks.
    await page.setViewportSize({ width: 1600, height: 1000 });
    await waitForLayoutSettle(page);
    await page.click('.tab-btn[data-tab="explorer"]');
  });
});

test.describe('visibility, measure, section & visual tools', () => {
  test('hide/show/isolate/reset visibility round-trip', async ({ appPage: page }) => {
    await ensureSingleSelection(page);
    const visibleBeforeReset = await page.locator('#visibleCount').textContent();

    await page.click('#btnHide');
    await waitForStatus(page, 'Selection hidden');

    await page.click('#btnShow');
    await waitForStatus(page, 'Selection shown');

    await page.click('#btnIsolate');
    await waitForStatus(page, 'Selection isolated');

    await page.click('#btnResetVisibility');
    await waitForStatus(page, 'Visibility reset');
    await expect(page.locator('#visibleCount')).toHaveText(visibleBeforeReset || '');
  });

  test('measurement, section and visual style tools toggle', async ({ appPage: page }) => {
    await page.click('#btnMeasureLength');
    await waitForStatus(page, 'Length measurement enabled');

    await page.click('#btnMeasureArea');
    await waitForStatus(page, 'Area measurement enabled');

    await page.click('#btnClearMeasurements');
    await waitForStatus(page, 'Measurements cleared');

    await page.click('#btnSectionX');
    await waitForStatus(page, 'Section plane added');

    await page.click('#btnSectionBox');
    await waitForStatus(page, 'Section box created');

    await page.click('#btnClearSections');
    await waitForStatus(page, 'Sections cleared');

    await page.click('#btnTransparency');
    await waitForStatus(page, 'X-ray enabled');

    await page.click('#btnWireframe');
    await waitForStatus(page, 'Edge overlay enabled');

    await page.click('#btnTransparency');
    await waitForStatus(page, 'X-ray disabled');

    await page.click('#btnWireframe');
    await waitForStatus(page, 'Edge overlay disabled');
  });
});

test.describe('models panel', () => {
  test('visual style, grid and background settings apply', async ({ appPage: page }) => {
    // Style/grid/background controls live only in the View menubar dropdown
    // since the ui-overhaul (AUDIT U2/T10) — open it before interacting.
    // Inline menu controls stopPropagation, so the menu stays open between
    // steps; a titlebar click closes it again at the end.
    await page.locator('.menu-dropdown', { has: page.locator('#visualStyleSelect') }).locator('.menu-item').click();
    await expect(page.locator('#visualStyleSelect')).toBeVisible();

    await page.selectOption('#visualStyleSelect', 'basic');
    await waitForStatus(page, 'Visual style: Basic');

    await page.selectOption('#visualStyleSelect', 'color-pen-shadows');
    await waitForStatus(page, 'Visual style: Color Pen Shadows');

    const gridWasChecked = await page.locator('#toggleGrid').isChecked();
    await page.click('#toggleGrid');
    await waitForStatus(page, gridWasChecked ? 'Grid hidden' : 'Grid enabled');

    await page.click('#toggleGrid');
    await waitForStatus(page, gridWasChecked ? 'Grid enabled' : 'Grid hidden');

    await page.click('[data-bg-preset="#c6d5e8"]');
    await waitForStatus(page, 'Background color set to #c6d5e8');
    await expect(page.locator('#backgroundColorInput')).toHaveValue('#c6d5e8');

    // AUDIT F8: a custom background survives a theme round-trip — the theme
    // toggle swaps per-theme memory instead of force-resetting defaults.
    await page.click('[data-bg-preset="#1b1f24"]');
    await waitForStatus(page, 'Background color set to #1b1f24');
    await page.click('#toggleTheme');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
    await expect(page.locator('#backgroundColorInput')).toHaveValue('#c6d5e8');
    await page.click('#toggleTheme');
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
    await expect(page.locator('#backgroundColorInput')).toHaveValue('#1b1f24');

    // Restore the default dark background for subsequent tests.
    await page.click('[data-bg-preset="#0b1220"]');
    await waitForStatus(page, 'Background color set to #0b1220');

    await page.locator('.app-titlebar').click();
    await expect(page.locator('#visualStyleSelect')).toBeHidden();
  });

  test('per-model visibility, opacity and transforms round-trip', async ({ appPage: page }, testInfo) => {
    await page.click('.tab-btn[data-tab="models"]');

    await page.click('[data-model-action="toggle-visibility"]');
    await waitForStatus(page, 'Hidden: school_str.ifc');

    await page.click('[data-model-action="toggle-visibility"]');
    await waitForStatus(page, 'Shown: school_str.ifc');

    const opacityInput = page.locator('input[data-model-opacity]').first();
    await opacityInput.fill('65');
    await opacityInput.dispatchEvent('change');
    await expect(page.locator('[data-opacity-value]').first()).toHaveText('65%');

    const xTransform = page.locator('input[data-transform="px"]').first();
    await xTransform.fill('1.5');
    await xTransform.dispatchEvent('change');
    await waitForStatus(page, 'Updated transform: school_str.ifc');

    const yRotation = page.locator('input[data-transform="ry"]').first();
    await yRotation.fill('90');
    await yRotation.dispatchEvent('change');
    await waitForStatus(page, 'Updated transform: school_str.ifc');

    await page.click('.tab-btn[data-tab="explorer"]');
    // Return to a known camera first: cube hotspots only render when facing
    // the camera (data-visible), and the preceding tests leave the camera on
    // a top-down view where the 'front' hotspot is hidden.
    const expectedHomeDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
    await page.click('#cubeHome');
    await waitForCameraDirection(page, expectedHomeDirection, 0.975);

    const expectedRotatedFrontDirection = await getExpectedCubeDirection(page, [0, 0, 1]);
    await clickVisibleCubeTarget(page, 'front');
    await waitForCameraDirection(page, expectedRotatedFrontDirection, 0.98);

    const expectedRotatedHomeDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
    await page.click('#cubeHome');
    await waitForCameraDirection(page, expectedRotatedHomeDirection, 0.975);
    // Viewport screenshot — see the stability note in the shell describe block.
    await page.screenshot({ path: testInfo.outputPath('view-cube-rotated.png') });

    await page.click('.tab-btn[data-tab="models"]');
    await page.locator('[data-model-action="reset"]').first().click();
    await waitForStatus(page, 'Reset transform: school_str.ifc');
    await expect(page.locator('input[data-transform="px"]').first()).toHaveValue('0.00');
    await expect(page.locator('input[data-transform="ry"]').first()).toHaveValue('0.0');

    await page.click('.tab-btn[data-tab="explorer"]');
    const beforeBrowserFit = await getCameraPosition(page);
    await page.locator('[data-browser-action="fit-model"]').first().click();
    await waitForCameraMove(page, beforeBrowserFit);
  });
});

test.describe('federation & load lifecycle', () => {
  // AUDIT A6 (metadata keyed by model id, no FIFO mis-attribution) and
  // F6 (per-model unload frees engine + viewer state).
  test('second model federates with correct metadata and unloads cleanly', async ({ appPage: page }) => {
    const elementsBefore = await page.locator('#elementCount').textContent();

    // AUDIT F3: X-ray/edges must survive a model load (setVisualStyle used to
    // wipe them on every registration).
    await page.click('#btnTransparency');
    await waitForStatus(page, 'X-ray enabled');
    await page.click('#btnWireframe');
    await waitForStatus(page, 'Edge overlay enabled');

    await page.setInputFiles('#fileInput', secondIfcPath);
    await waitForModelCount(page, 2);

    await expect(page.locator('#btnTransparency')).toHaveClass(/active/);
    await expect(page.locator('#btnWireframe')).toHaveClass(/active/);
    expect(await page.evaluate(() => {
      const viewer = (window as any).__viewer;
      return { xray: viewer.xrayEnabled, edges: viewer.edgesEnabled };
    })).toEqual({ xray: true, edges: true });

    // Restore toggles before the panel assertions below.
    await page.click('#btnTransparency');
    await waitForStatus(page, 'X-ray disabled');
    await page.click('#btnWireframe');
    await waitForStatus(page, 'Edge overlay disabled');

    await page.click('.tab-btn[data-tab="models"]');
    await expect(page.locator('.federated-model')).toHaveCount(2);
    // Metadata attribution (A6): each card carries its own file name.
    await expect(page.locator('.federated-model-name-btn').nth(0)).toContainText('school_str.ifc');
    await expect(page.locator('.federated-model-name-btn').nth(1)).toContainText('Ifc4_Revit_ARC.ifc');

    const elementsWithTwo = await page.locator('#elementCount').textContent();
    expect(elementsWithTwo).not.toBe(elementsBefore);

    // AUDIT F9: issues capture the full multi-model selection (not just the
    // first model's elements) while both models are loaded.
    await page.evaluate(() => {
      const viewer = (window as any).__viewer;
      for (const key of Object.keys(viewer.selectedItems)) delete viewer.selectedItems[key];
      for (const [modelId, index] of viewer.modelIndices.entries()) {
        const firstId = Array.from(index.allIds)[0];
        if (typeof firstId === 'number') viewer.selectedItems[modelId] = new Set([firstId]);
      }
    });
    await page.click('.tab-btn[data-tab="issues"]');
    await page.fill('#issueTitle', 'Multi-model issue');
    await page.click('#btnCreateIssue');
    await waitForStatus(page, 'Issue created');
    const captured = await page.evaluate(() => {
      const issue = (window as any).__viewer.issues[0];
      return {
        models: Object.keys(issue.elementsByModel ?? {}).length,
        legacyModelId: typeof issue.modelId === 'string',
      };
    });
    expect(captured.models).toBe(2);
    expect(captured.legacyModelId).toBe(true);
    await page.locator('[data-issue-id]').first().click();
    await page.click('#btnDeleteIssue');
    await page.click('.confirm-btn-confirm');
    await page.waitForFunction(
      () => !(document.querySelector('#issuesList')?.textContent || '').includes('Multi-model issue'),
      undefined,
      { timeout: 20_000 },
    );
    await page.click('.tab-btn[data-tab="models"]');

    // F6: unload the second model via the federation panel action.
    await page.locator('[data-model-action="unload"]').nth(1).click();
    await page.click('.confirm-btn-confirm');
    await waitForStatus(page, 'Model unloaded: Ifc4_Revit_ARC.ifc', 30_000);

    await expect(page.locator('.federated-model')).toHaveCount(1);
    await expect(page.locator('#elementCount')).toHaveText(elementsBefore || '');
    const engineState = await page.evaluate(() => {
      const viewer = (window as any).__viewer;
      return {
        fragmentsCount: viewer.fragments.list.size,
        federatedCount: viewer.federatedModels.size,
        indexCount: viewer.modelIndices.size,
        objectCount: viewer.modelObjects.length,
      };
    });
    expect(engineState).toEqual({ fragmentsCount: 1, federatedCount: 1, indexCount: 1, objectCount: 1 });

    await page.click('.tab-btn[data-tab="explorer"]');
  });
});

test.describe('error surfacing (U4)', () => {
  test('failed imports and loads surface as toasts and overlay error state', async ({ appPage: page }, testInfo) => {
    // Invalid state import → error toast (catch paths must not stay silent).
    const badStatePath = testInfo.outputPath('bad-state.json');
    fs.writeFileSync(badStatePath, 'this is not json');
    await page.setInputFiles('#importStateInput', badStatePath);
    await expect(page.locator('.toast-error')).toBeVisible();
    await waitForStatus(page, 'Import failed');
    // AUDIT U11: toasts anchor bottom-right, clear of the view cube (top-right).
    const toastBox = await page.locator('.toast-error').boundingBox();
    const viewportSize = page.viewportSize();
    expect(toastBox).not.toBeNull();
    expect(viewportSize).not.toBeNull();
    if (toastBox && viewportSize) {
      expect(toastBox.y).toBeGreaterThan(viewportSize.height / 2);
    }
    // Toasts auto-dismiss — wait so later assertions see a clean slate.
    await page.waitForSelector('.toast-error', { state: 'detached', timeout: 10_000 });

    // Corrupt IFC → overlay error state with Retry/Dismiss (U4).
    const badIfcPath = testInfo.outputPath('corrupt.ifc');
    fs.writeFileSync(badIfcPath, 'NOT-AN-IFC-FILE');
    await page.setInputFiles('#fileInput', badIfcPath);
    await expect(page.locator('#loadingOverlay')).toHaveClass(/is-error/, { timeout: 60_000 });
    await expect(page.locator('#loadingErrorActions')).toBeVisible();
    await expect(page.locator('.toast-error')).toBeVisible();

    // Retry replays the failed file and fails into the same error state.
    await page.click('#btnRetryLoad');
    await expect(page.locator('#loadingOverlay')).toHaveClass(/is-error/, { timeout: 60_000 });

    // Dismiss returns to the normal (model still loaded) viewer.
    await page.click('#btnDismissLoadError');
    await expect(page.locator('#loadingOverlay')).toBeHidden();
    await expect(page.locator('#emptyState')).toBeHidden();
  });
});

test.describe('viewpoints, issues & state persistence', () => {
  test('viewpoint/issue lifecycle and state export/import round-trip', async ({ appPage: page, browser }, testInfo) => {
    await ensureSingleSelection(page);

    await page.click('.tab-btn[data-tab="viewpoints"]');
    await page.fill('#viewpointName', 'QA View');
    await page.click('#btnSaveViewpoint');
    await waitForStatus(page, 'Saved viewpoint: QA View', 30_000);

    // AUDIT F2 regression: the snapshot must be a real (non-blank) capture —
    // decode it and assert pixel variance; size/mime prove the ≤320px JPEG
    // thumbnail contract. A blank transparent capture decodes to zero spread.
    const snapshotInfo = await page.evaluate(async () => {
      const viewpoint = (window as any).__viewer.viewpoints[0];
      if (!viewpoint?.snapshot) return null;
      const image = new Image();
      await new Promise((resolve, reject) => {
        image.onload = resolve;
        image.onerror = reject;
        image.src = viewpoint.snapshot;
      });
      const canvas = document.createElement('canvas');
      canvas.width = image.naturalWidth;
      canvas.height = image.naturalHeight;
      const context = canvas.getContext('2d');
      if (!context) return null;
      context.drawImage(image, 0, 0);
      const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
      let min = 255;
      let max = 0;
      for (let i = 0; i < data.length; i += 4) {
        const luminance = (data[i] + data[i + 1] + data[i + 2]) / 3;
        if (luminance < min) min = luminance;
        if (luminance > max) max = luminance;
      }
      return {
        mime: viewpoint.snapshot.slice(5, viewpoint.snapshot.indexOf(';')),
        width: image.naturalWidth,
        height: image.naturalHeight,
        spread: max - min,
      };
    });
    expect(snapshotInfo).not.toBeNull();
    expect(snapshotInfo?.mime).toBe('image/jpeg');
    expect(Math.max(snapshotInfo?.width ?? 0, snapshotInfo?.height ?? 0)).toBeLessThanOrEqual(320);
    expect(snapshotInfo?.spread ?? 0).toBeGreaterThan(25);
    // Thumbnail is rendered in the viewpoint list.
    await expect(page.locator('.viewpoint-thumb').first()).toBeVisible();

    await page.locator('[data-viewpoint-id]').first().click();
    await page.click('#btnApplySelectedViewpoint');
    await waitForStatus(page, 'Applied viewpoint: QA View', 30_000);

    await ensureSingleSelection(page);
    await page.click('.tab-btn[data-tab="issues"]');
    await page.fill('#issueTitle', 'QA Issue');
    await page.fill('#issueDescription', 'Created during automated QA');
    await page.fill('#issueAssignee', 'Automation');
    await page.click('#btnCreateIssue');
    await waitForStatus(page, 'Issue created');
    await expect(page.locator('[data-issue-id]').first()).toContainText('QA Issue');

    await page.locator('[data-issue-id]').first().click();
    await page.fill('#issueCommentInput', 'Follow-up note');
    await page.click('#btnAddIssueComment');
    await waitForStatus(page, 'Comment added');
    await expect(page.locator('#issueComments')).toContainText('Follow-up note');

    const importChooser = page.waitForEvent('filechooser');
    await page.click('#btnImportState');
    await importChooser;

    const exportDownload = page.waitForEvent('download');
    await page.click('#btnExportState');
    const exportedState = await exportDownload;
    expect(exportedState.suggestedFilename()).toMatch(/\.json$/);

    const exportedStatePath = testInfo.outputPath('viewer-state.json');
    await exportedState.saveAs(exportedStatePath);

    await page.click('.tab-btn[data-tab="viewpoints"]');
    await page.locator('[data-viewpoint-id]').first().click();
    await page.click('#btnDeleteSelectedViewpoint');
    // Destructive actions open the app's confirm dialog (AUDIT U8) — accept it.
    await page.click('.confirm-btn-confirm');
    await page.waitForFunction(
      () => !(document.querySelector('#viewpointList')?.textContent || '').includes('QA View'),
      undefined,
      { timeout: 20_000 },
    );

    await page.click('.tab-btn[data-tab="issues"]');
    await page.locator('[data-issue-id]').first().click();
    await page.click('#btnDeleteIssue');
    await page.click('.confirm-btn-confirm');
    await page.waitForFunction(
      () => !(document.querySelector('#issuesList')?.textContent || '').includes('QA Issue'),
      undefined,
      { timeout: 20_000 },
    );

    const importContext = await browser.newContext({
      acceptDownloads: true,
      viewport: { width: 1600, height: 1000 },
    });
    const importPage = await importContext.newPage();

    try {
      await waitForAppReady(importPage);
      await importPage.setInputFiles('#importStateInput', exportedStatePath);
      await waitForStatus(importPage, 'Viewer data imported', 30_000);

      await importPage.click('.tab-btn[data-tab="viewpoints"]');
      await expect(importPage.locator('#viewpointList')).toContainText('QA View');

      await importPage.click('.tab-btn[data-tab="issues"]');
      await expect(importPage.locator('#issuesList')).toContainText('QA Issue');
    } finally {
      await importPage.close();
      await importContext.close();
    }
  });
});
