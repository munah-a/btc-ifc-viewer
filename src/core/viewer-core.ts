/**
 * Engine core: the ThatOpen/three bootstrap and a couple of engine-adjacent
 * lifecycle helpers, extracted from viewer.ts so the full app AND the future
 * chromeless /embed entry (W4) share one engine setup.
 *
 * `bootstrapEngine` builds the OBC world (SimpleScene + OrthoPerspectiveCamera +
 * PostproductionRenderer), the grid, the local-wasm IfcLoader, the local-worker
 * FragmentsManager, the Clipper/Hider/Raycaster, the length/area measurements,
 * the marker manager and the transform-controls gizmo — then returns the raw
 * handles. It deliberately does NOT wire the app's `this`-coupled callbacks
 * (model registration, gizmo → panel re-render, camera → fragments update): the
 * orchestrator wires those against the returned handles so behaviour is
 * identical to the pre-extraction inline `initEngine`.
 *
 * `ShaderWarningFilter` is the A5 scoped console.warn filter (install/uninstall
 * paired so destroy() restores the original). `createFpsMonitor` is the rAF FPS
 * counter that writes to the status bar.
 */
import * as THREE from 'three';
import * as OBC from '@thatopen/components';
import * as OBCF from '@thatopen/components-front';
import { TransformControls } from 'three/examples/jsm/controls/TransformControls.js';

export interface EngineHandles {
  components: OBC.Components;
  world: OBC.SimpleWorld<OBC.SimpleScene, OBC.OrthoPerspectiveCamera, OBCF.PostproductionRenderer>;
  ifcLoader: OBC.IfcLoader;
  fragments: OBC.FragmentsManager;
  clipper: OBC.Clipper;
  hider: OBC.Hider;
  raycaster: ReturnType<OBC.Raycasters['get']>;
  lengthMeasurement: OBCF.LengthMeasurement;
  areaMeasurement: OBCF.AreaMeasurement;
  markerManager: OBCF.Marker;
  transformControls: TransformControls;
  transformControlsHelper: THREE.Object3D;
  gridHelper: THREE.Object3D;
}

export interface BootstrapEngineOptions {
  /** Container the PostproductionRenderer renders into. */
  container: HTMLElement;
  /** Initial scene background color (hex). */
  backgroundColor: string;
  /** Initial grid visibility. */
  gridVisible: boolean;
  /**
   * Live read of the current background color for the PEN-mode composer clear
   * hook (the color can change at runtime via the theme/background picker).
   */
  getBackgroundColor: () => string;
  /** Narrows world.renderer to the PostproductionRenderer, or null. */
  getPostproductionRenderer: () => OBCF.PostproductionRenderer | null;
}

/**
 * Builds and initializes the engine. Async: awaits camera framing, the IFC
 * loader wasm setup and the fragments worker fetch (self-hosted assets — no
 * CDN, C1/A2). Returns the engine handles for the orchestrator to wire.
 */
export async function bootstrapEngine(options: BootstrapEngineOptions): Promise<EngineHandles> {
  const { container, backgroundColor, gridVisible, getBackgroundColor, getPostproductionRenderer } = options;

  const components = new OBC.Components();
  const worlds = components.get(OBC.Worlds);

  const world = worlds.create<OBC.SimpleScene, OBC.OrthoPerspectiveCamera, OBCF.PostproductionRenderer>();
  world.scene = new OBC.SimpleScene(components);
  world.scene.setup();
  world.scene.three.background = new THREE.Color(backgroundColor);

  world.renderer = new OBCF.PostproductionRenderer(components, container);

  // Hook into the render loop to clear composer targets before each frame in PEN mode.
  // PEN mode (PostproductionAspect.PEN) skips the BasePass, so the EffectComposer's
  // read/write buffers never get cleared — causing ghost lines from previous frames.
  world.renderer.onBeforeUpdate.add(() => {
    const postRenderer = getPostproductionRenderer();
    const post = postRenderer?.postproduction;
    if (!post?.enabled || !post.composer) return;
    // PEN = 1, PEN_SHADOWS = 2 — these use EdgeDetectionPass without a prior clear
    const isPenStyle = post.style === OBCF.PostproductionAspect.PEN
      || post.style === OBCF.PostproductionAspect.PEN_SHADOWS;
    if (!isPenStyle) return;
    const renderer = postRenderer!.three;
    const bgColor = new THREE.Color(getBackgroundColor());
    renderer.setClearColor(bgColor, 1);
    renderer.setRenderTarget(post.composer.renderTarget1);
    renderer.clear();
    renderer.setRenderTarget(post.composer.renderTarget2);
    renderer.clear();
    renderer.setRenderTarget(null);
  });
  world.camera = new OBC.OrthoPerspectiveCamera(components);
  await world.camera.controls.setLookAt(18, 18, 18, 0, 0, 0);

  components.init();

  const grids = components.get(OBC.Grids);
  const grid = grids.create(world);
  const gridHelper = grid as unknown as THREE.Object3D;
  gridHelper.visible = gridVisible;
  if (grid.material?.uniforms?.uColor) grid.material.uniforms.uColor.value = new THREE.Color(0x25334a);

  // Self-hosted runtime assets (A2/P2): web-ifc.wasm and the fragments
  // worker are vendored from node_modules into public/ by
  // scripts/vendor-assets.mjs (prebuild/predev) — no CDN at runtime (C1).
  const ifcLoader = components.get(OBC.IfcLoader);
  await ifcLoader.setup({
    autoSetWasm: false,
    wasm: {
      path: import.meta.env.BASE_URL,
      absolute: true,
    },
  });
  ifcLoader.settings.webIfc.CIRCLE_SEGMENTS = 24;

  const fragments = components.get(OBC.FragmentsManager);
  const workerUrl = `${import.meta.env.BASE_URL}worker.mjs`;
  const fetchedWorker = await fetch(workerUrl);
  const workerBlob = await fetchedWorker.blob();
  const workerFile = new File([workerBlob], 'worker.mjs', { type: 'text/javascript' });
  fragments.init(URL.createObjectURL(workerFile));
  fragments.core.settings.graphicsQuality = 1;

  fragments.core.models.materials.list.onItemSet.add(({ value: material }) => {
    if (!('isLodMaterial' in material && (material as unknown as { isLodMaterial: boolean }).isLodMaterial)) {
      const cast = material as THREE.Material & {
        polygonOffset?: boolean;
        polygonOffsetFactor?: number;
        polygonOffsetUnits?: number;
      };
      cast.polygonOffset = true;
      cast.polygonOffsetFactor = 1;
      cast.polygonOffsetUnits = 1;
    }
  });

  const clipper = components.get(OBC.Clipper);
  clipper.enabled = false;
  const hider = components.get(OBC.Hider);
  const raycasters = components.get(OBC.Raycasters);
  const raycaster = raycasters.get(world);

  const lengthMeasurement = components.get(OBCF.LengthMeasurement);
  lengthMeasurement.world = world;
  lengthMeasurement.enabled = false;

  const areaMeasurement = components.get(OBCF.AreaMeasurement);
  areaMeasurement.world = world;
  areaMeasurement.enabled = false;

  const markerManager = components.get(OBCF.Marker);
  markerManager.threshold = 64;
  markerManager.autoCluster = true;

  const transformControls = new TransformControls(world.camera.three, world.renderer.three.domElement);
  transformControls.setSize(0.75);
  transformControls.setSpace('world');
  transformControls.enabled = false;
  const transformControlsHelper = transformControls.getHelper();
  transformControlsHelper.visible = false;
  world.scene.three.add(transformControlsHelper);

  return {
    components,
    world,
    ifcLoader,
    fragments,
    clipper,
    hider,
    raycaster,
    lengthMeasurement,
    areaMeasurement,
    markerManager,
    transformControls,
    transformControlsHelper,
    gridHelper,
  };
}

/**
 * P6 (W5.3): switches the world's PostproductionRenderer to on-demand (MANUAL)
 * rendering and returns a `requestRender()` the orchestrator calls on any visual
 * change. In MANUAL mode the renderer only composites a frame when `needsUpdate`
 * is set, so an idle viewport stops burning the GPU on continuous postprocessing
 * (battery — a field/tablet win, C6) while a static frame stays on the canvas
 * (required for e2e/page.screenshot capture).
 *
 * Visual parity: `turnOffOnManualMode` drops the heavy postprocessing DURING
 * navigation for fluidity and restores it `manualModeDelay` ms after the last
 * change — so the settled frame carries full outlines/gloss exactly as before.
 *
 * The renderer keeps rendering while the camera controls animate (their `update`
 * event fires each frame and re-arms `needsUpdate`). Mesh streaming (LOD) and
 * all other visual changes are re-armed by the orchestrator calling the returned
 * `requestRender()` — after each `fragments.core.update()`, on each model's
 * `onViewUpdated`, and on selection/tool/marker changes.
 *
 * Returns a `requestRender` fn AND a `stop` to detach the camera listener
 * (destroy()).
 */
export function enableOnDemandRendering(engine: EngineHandles): {
  requestRender: () => void;
  stop: () => void;
} {
  const renderer = engine.world.renderer as OBCF.PostproductionRenderer & {
    mode: OBC.RendererMode;
    needsUpdate: boolean;
    turnOffOnManualMode?: boolean;
    manualModeDelay?: number;
  };

  const requestRender = (): void => {
    renderer.needsUpdate = true;
  };

  renderer.mode = OBC.RendererMode.MANUAL;
  // Keep full postprocessing on the settled frame; only relax it while moving.
  renderer.turnOffOnManualMode = true;
  renderer.manualModeDelay = 200;
  requestRender();

  // Camera animation frames + user navigation → re-arm each tick while moving.
  const onCameraUpdate = (): void => requestRender();
  engine.world.camera.controls.addEventListener('update', onCameraUpdate);

  return {
    requestRender,
    stop: () => {
      engine.world.camera.controls.removeEventListener('update', onCameraUpdate);
    },
  };
}

/**
 * Visual-style → PostproductionAspect mapping (W5.1). Kept in viewer-core so the
 * `OBCF.PostproductionAspect` enum — a runtime *value* from @thatopen — stays
 * inside this dynamically-imported engine module. Were viewer.ts to reference the
 * enum directly, its static import of @thatopen would pin the (large) engine
 * chunk into the initial shell, defeating the P1 code-split.
 *
 * Applies the style's aspect + the outline/gloss flags to a postproduction
 * instance. `post` is typed loosely because @thatopen does not export the
 * postproduction shape as a public type; the members used here are stable.
 */
export type PostproductionStyle =
  | 'basic'
  | 'pen'
  | 'color-pen'
  | 'color-shadows'
  | 'color-pen-shadows';

interface PostproductionLike {
  style: number;
  outlinesEnabled: boolean;
  glossEnabled: boolean;
}

export function applyPostproductionStyle(post: PostproductionLike, style: PostproductionStyle): void {
  switch (style) {
    case 'basic':
      post.style = OBCF.PostproductionAspect.COLOR;
      break;
    case 'pen':
      post.style = OBCF.PostproductionAspect.PEN;
      post.outlinesEnabled = true;
      break;
    case 'color-pen':
      post.style = OBCF.PostproductionAspect.COLOR_PEN;
      post.outlinesEnabled = true;
      break;
    case 'color-shadows':
      post.style = OBCF.PostproductionAspect.COLOR_SHADOWS;
      post.glossEnabled = true;
      break;
    case 'color-pen-shadows':
    default:
      post.style = OBCF.PostproductionAspect.COLOR_PEN_SHADOWS;
      post.outlinesEnabled = true;
      post.glossEnabled = true;
      break;
  }
}

// A5 ShaderWarningFilter now lives in the dependency-free core/engine-lite.ts
// (W5.1) so the app shell can install it before the heavy engine chunk loads.
// Re-exported here for consumers that already import it from viewer-core.
export { ShaderWarningFilter } from './engine-lite';

/**
 * Real-rendered-frame FPS counter (A15, W5.3). Counts the renderer's actual
 * `onAfterUpdate` firings — which happen ONLY on a real `three.render()` — and
 * samples the count into `output` once a second. Under on-demand rendering this
 * reads the true rendered rate (≈0 when idle, the real value while interacting),
 * not the rAF cadence the old counter measured (AUDIT A15). The 1 s sampler runs
 * off setInterval, so it needs no rAF of its own.
 *
 * `renderer` is loosely typed to the `onAfterUpdate` event it exposes so this
 * stays decoupled from the concrete renderer class.
 */
export function createFpsMonitor(
  output: HTMLElement,
  renderer?: { onAfterUpdate: { add(cb: () => void): void; remove(cb: () => void): void } },
): { stop: () => void } {
  let frameCount = 0;
  const onFrame = (): void => {
    frameCount += 1;
  };

  // Fallback (no renderer given): count rAF ticks as before.
  let frameId: number | null = null;
  if (renderer) {
    renderer.onAfterUpdate.add(onFrame);
  } else {
    const tick = (): void => {
      frameCount += 1;
      frameId = requestAnimationFrame(tick);
    };
    frameId = requestAnimationFrame(tick);
  }

  const sampler = setInterval(() => {
    output.textContent = `${frameCount} FPS`;
    frameCount = 0;
  }, 1000);

  return {
    stop: () => {
      clearInterval(sampler);
      if (renderer) renderer.onAfterUpdate.remove(onFrame);
      if (frameId !== null) {
        cancelAnimationFrame(frameId);
        frameId = null;
      }
    },
  };
}
