import path from 'node:path';

import { expect, test, type Page } from '@playwright/test';

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

const ifcPath = path.join(process.cwd(), 'public', 'school_str.ifc');
const viewerUrl = 'http://127.0.0.1:4173/btc-ifc-viewer/';

const waitForAppReady = async (page: Page): Promise<void> => {
  await page.goto(viewerUrl);
  await page.waitForFunction(
    () => (document.querySelector('#statusText')?.textContent || '').includes('Ready'),
    undefined,
    { timeout: 60_000 },
  );
};

const waitForModelReady = async (page: Page): Promise<void> => {
  await page.waitForFunction(
    () => {
      const viewer = (window as any).__viewer;
      const elementText = document.querySelector('#elementCount')?.textContent || '';
      return !!viewer
        && viewer.federatedModels?.size > 0
        && viewer.registeringModelIds?.size === 0
        && elementText !== '0 elements';
    },
    undefined,
    { timeout: 180_000 },
  );
};

const waitForStatus = async (page: Page, text: string, timeout = 20_000): Promise<void> => {
  await page.waitForFunction(
    (expected) => (document.querySelector('#statusText')?.textContent || '').includes(expected),
    text,
    { timeout },
  );
};

const openDock = async (page: Page, title: string): Promise<void> => {
  await page.click(`[data-dock-toggle][title="${title}"]`);
};

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

const directionMatches = (actual: CameraPosition, expected: CameraPosition, threshold = 0.985): boolean => (
  dotProduct(normalizeVector(actual), normalizeVector(expected)) >= threshold
);

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
  return expected as CameraPosition;
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

test('viewer smoke regression', async ({ browser, page }, testInfo) => {
  test.slow();

  await waitForAppReady(page);

  const emptyUploadChooser = page.waitForEvent('filechooser');
  await page.click('#btnUploadEmpty');
  await emptyUploadChooser;

  await page.setInputFiles('#fileInput', ifcPath);
  await waitForModelReady(page);

  const headerUploadChooser = page.waitForEvent('filechooser');
  await page.click('#btnUpload');
  await headerUploadChooser;

  await expect(page.locator('.header-center')).toHaveCount(0);
  await expect(page.locator('#elementCount')).not.toHaveText(/0 elements/);
  await expect(page.locator('#visibleCount')).not.toHaveText(/0 visible/);

  const screenshotDownload = page.waitForEvent('download');
  await page.click('#btnExportScreenshot');
  expect((await screenshotDownload).suggestedFilename()).toMatch(/\.png$/);

  for (const tab of ['explorer', 'models', 'properties', 'viewpoints', 'issues', 'help']) {
    await page.click(`.tab-btn[data-tab="${tab}"]`);
    await expect(page.locator(`#panel-${tab}`)).toHaveClass(/active/);
  }
  await page.click('.tab-btn[data-tab="explorer"]');
  await page.locator('#viewCube').screenshot({ path: testInfo.outputPath('view-cube-home.png') });

  await openDock(page, 'Navigate');
  await page.click('#btnModePlan');
  await expect(page.locator('#btnModePlan')).toHaveClass(/active/);

  await openDock(page, 'Navigate');
  await page.click('#btnModeFirstPerson');
  await expect(page.locator('#btnModeFirstPerson')).toHaveClass(/active/);

  await openDock(page, 'Navigate');
  await page.click('#btnModeOrbit');
  await expect(page.locator('#btnModeOrbit')).toHaveClass(/active/);

  const beforeFront = await getCameraPosition(page);
  const expectedFrontDirection = await getExpectedCubeDirection(page, [0, 0, 1]);
  await clickVisibleCubeTarget(page, 'front');
  await page.waitForTimeout(750);
  const afterFront = await getCameraPosition(page);
  expect(cameraChanged(beforeFront, afterFront)).toBeTruthy();
  expect(directionMatches(await getCameraDirection(page), expectedFrontDirection)).toBeTruthy();

  const expectedCornerDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
  await clickVisibleCubeTarget(page, 'top-front-right');
  await page.waitForTimeout(750);
  expect(directionMatches(await getCameraDirection(page), expectedCornerDirection, 0.975)).toBeTruthy();

  const expectedHomeDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
  await page.click('#cubeHome');
  await page.waitForTimeout(750);
  expect(directionMatches(await getCameraDirection(page), expectedHomeDirection, 0.975)).toBeTruthy();

  await openDock(page, 'Navigate');
  await page.click('#btnTop');
  await page.waitForTimeout(750);
  const afterTop = await getCameraPosition(page);
  expect(cameraChanged(afterFront, afterTop)).toBeTruthy();

  await openDock(page, 'Select');
  await page.click('#btnSelectMulti');
  await expect(page.locator('#btnSelectMulti')).toHaveClass(/active/);

  await openDock(page, 'Select');
  await page.click('#btnSelectSingle');
  await expect(page.locator('#btnSelectSingle')).toHaveClass(/active/);

  const modelContext = await getModelContext(page);
  await page.fill('#searchInput', modelContext.searchTerm.split(':')[0].trim());
  await page.click('#btnSearch');
  await expect(page.locator('.result-item').first()).toBeVisible();
  await page.locator('.result-item').first().click();
  await expect(page.locator('#selectionCount')).toHaveText(/1 selected/);

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
  await page.locator('#panel-properties').screenshot({ path: testInfo.outputPath('properties-panel.png') });
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
  await page.waitForTimeout(250);
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
    const sections = document.querySelector('#propSections') as HTMLElement | null;
    const panel = document.querySelector('#panel-properties') as HTMLElement | null;
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
  await page.click('.tab-btn[data-tab="explorer"]');

  await ensureSingleSelection(page);
  const visibleBeforeReset = await page.locator('#visibleCount').textContent();

  await openDock(page, 'Select');
  await page.click('#btnHide');
  await waitForStatus(page, 'Selection hidden');

  await openDock(page, 'Select');
  await page.click('#btnShow');
  await waitForStatus(page, 'Selection shown');

  await openDock(page, 'Select');
  await page.click('#btnIsolate');
  await waitForStatus(page, 'Selection isolated');

  await openDock(page, 'Select');
  await page.click('#btnResetVisibility');
  await waitForStatus(page, 'Visibility reset');
  await expect(page.locator('#visibleCount')).toHaveText(visibleBeforeReset || '');

  await openDock(page, 'Measure');
  await page.click('#btnMeasureLength');
  await waitForStatus(page, 'Length measurement enabled');

  await openDock(page, 'Measure');
  await page.click('#btnMeasureArea');
  await waitForStatus(page, 'Area measurement enabled');

  await openDock(page, 'Measure');
  await page.click('#btnClearMeasurements');
  await waitForStatus(page, 'Measurements cleared');

  await openDock(page, 'Section');
  await page.click('#btnSectionX');
  await waitForStatus(page, 'Section plane added');

  await openDock(page, 'Section');
  await page.click('#btnSectionBox');
  await waitForStatus(page, 'Section box created');

  await openDock(page, 'Section');
  await page.click('#btnClearSections');
  await waitForStatus(page, 'Sections cleared');

  await openDock(page, 'Visual');
  await page.click('#btnTransparency');
  await waitForStatus(page, 'X-ray enabled');

  await openDock(page, 'Visual');
  await page.click('#btnWireframe');
  await waitForStatus(page, 'Edge overlay enabled');

  await openDock(page, 'Visual');
  await page.click('#btnTransparency');
  await waitForStatus(page, 'X-ray disabled');

  await openDock(page, 'Visual');
  await page.click('#btnWireframe');
  await waitForStatus(page, 'Edge overlay disabled');

  await page.click('.tab-btn[data-tab="models"]');
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
  const expectedRotatedFrontDirection = await getExpectedCubeDirection(page, [0, 0, 1]);
  await clickVisibleCubeTarget(page, 'front');
  await page.waitForTimeout(750);
  expect(directionMatches(await getCameraDirection(page), expectedRotatedFrontDirection, 0.98)).toBeTruthy();

  const expectedRotatedHomeDirection = await getExpectedCubeDirection(page, [1, 1, 1]);
  await page.click('#cubeHome');
  await page.waitForTimeout(750);
  expect(directionMatches(await getCameraDirection(page), expectedRotatedHomeDirection, 0.975)).toBeTruthy();
  await page.locator('#viewCube').screenshot({ path: testInfo.outputPath('view-cube-rotated.png') });

  await page.click('.tab-btn[data-tab="models"]');
  await page.locator('[data-model-action="reset"]').first().click();
  await waitForStatus(page, 'Reset transform: school_str.ifc');
  await expect(page.locator('input[data-transform="px"]').first()).toHaveValue('0.00');
  await expect(page.locator('input[data-transform="ry"]').first()).toHaveValue('0.0');

  await page.click('.tab-btn[data-tab="explorer"]');
  const beforeBrowserFit = await getCameraPosition(page);
  await page.locator('[data-browser-action="fit-model"]').first().click();
  await page.waitForTimeout(750);
  const afterBrowserFit = await getCameraPosition(page);
  expect(cameraChanged(beforeBrowserFit, afterBrowserFit)).toBeTruthy();

  await page.click('.tab-btn[data-tab="viewpoints"]');
  await page.fill('#viewpointName', 'QA View');
  await page.click('#btnSaveViewpoint');
  await waitForStatus(page, 'Saved viewpoint: QA View', 30_000);
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
  await page.waitForFunction(
    () => !(document.querySelector('#viewpointList')?.textContent || '').includes('QA View'),
    undefined,
    { timeout: 20_000 },
  );

  await page.click('.tab-btn[data-tab="issues"]');
  await page.locator('[data-issue-id]').first().click();
  await page.click('#btnDeleteIssue');
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
