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
 * A5: scoped console.warn filter. Suppresses ONLY the known three.js WebGL
 * program info-log noise; everything else passes through. install/uninstall are
 * paired so destroy() can restore the original console.warn rather than leaving
 * it permanently monkey-patched.
 */
export class ShaderWarningFilter {
  private originalConsoleWarn: typeof console.warn | null = null;
  private installed = false;

  install(): void {
    if (this.installed) return;
    const originalWarn = console.warn.bind(console);
    this.originalConsoleWarn = originalWarn;
    console.warn = (...args: unknown[]) => {
      const header = typeof args[0] === 'string' ? args[0] : '';
      const payload = args
        .map((entry) => (typeof entry === 'string' ? entry : ''))
        .join(' ');
      const isThreeProgramLog = header.includes('THREE.WebGLProgram: Program Info Log:');
      const isKnownNoise = payload.includes('dyn_index_vec4_float4_int');
      if (isThreeProgramLog && isKnownNoise) return;
      originalWarn(...args);
    };
    this.installed = true;
  }

  uninstall(): void {
    if (!this.installed || !this.originalConsoleWarn) return;
    console.warn = this.originalConsoleWarn;
    this.originalConsoleWarn = null;
    this.installed = false;
  }
}

/**
 * rAF FPS counter (A15). Writes `<n> FPS` into `output` roughly once a second.
 * Returns a stop() to cancel the loop (wired to destroy()).
 */
export function createFpsMonitor(output: HTMLElement): { stop: () => void } {
  let frameCount = 0;
  let lastTs = performance.now();
  let frameId: number | null = null;

  const tick = (): void => {
    frameCount += 1;
    const now = performance.now();
    if (now - lastTs >= 1000) {
      output.textContent = `${frameCount} FPS`;
      frameCount = 0;
      lastTs = now;
    }
    frameId = requestAnimationFrame(tick);
  };
  frameId = requestAnimationFrame(tick);

  return {
    stop: () => {
      if (frameId !== null) {
        cancelAnimationFrame(frameId);
        frameId = null;
      }
    },
  };
}
