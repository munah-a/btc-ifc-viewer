/**
 * Chromeless embed entry (W4.1).
 *
 * A minimal, self-contained viewer for boards / decks / docs. It reuses the
 * SHARED engine bootstrap (core/viewer-core.ts) — the exact same OBC world the
 * full app uses — but ships none of the app chrome: just a canvas, orbit, fit
 * and fullscreen, a BTC badge and an "Open in viewer" link.
 *
 * Design decisions per the plan:
 *  • Poster + click-to-activate: the WebGL context (and the whole engine) is NOT
 *    created until the user activates, so a board with many embeds does not blow
 *    the browser's WebGL-context budget.
 *  • On-demand posture (a subset of P6): the embed uses the cheapest visual
 *    style (no SMAA/outline postprocessing) and only pushes fragment updates on
 *    camera movement / model load / resize — no heavy continuous postprocessing.
 *  • Loads a model BY URL (?m=) — a browser-converted `.frag` fetched directly
 *    from the Blob CDN (C2: the server never processed it), or an `.ifc` URL
 *    which the embed converts client-side. The viewpoint (?vp=) restores camera,
 *    projection, section planes, x-ray/edges and a hidden-items summary.
 *  • initLanguage + hydrateIcons + hydrateI18n at bootstrap (C7 / C1).
 */
import * as THREE from 'three';

import { bootstrapEngine, type EngineHandles } from './core/viewer-core';
import { isProbablyIfc } from './core/ifc-format';
import { hydrateI18n, initLanguage, t } from './core/i18n';
import { decodeUrlState, encodeUrlState, type UrlState, type UrlViewpointState } from './core/url-state';
import type { FragmentsModelLike } from './core/fragments-model';
import { hydrateIcons } from './ui/icons';

const EMBED_MODEL_ID = 'embed-model';
const DEFAULT_BG = '#14151b';

interface EmbedDom {
  root: HTMLElement;
  viewer: HTMLElement;
  poster: HTMLButtonElement;
  loading: HTMLElement;
  loadingText: HTMLElement;
  loadingFill: HTMLElement;
  errorBox: HTMLElement;
  errorMsg: HTMLElement;
  errorOpen: HTMLAnchorElement;
  controls: HTMLElement;
  fitBtn: HTMLButtonElement;
  fullscreenBtn: HTMLButtonElement;
  badge: HTMLAnchorElement;
  oembedLink: HTMLLinkElement | null;
}

function query<T extends Element>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`[embed] missing element #${id}`);
  return el as unknown as T;
}

/** The full-app URL that "Open in viewer" / the badge point at (same model + view). */
function openInViewerUrl(state: UrlState): string {
  const origin = window.location.origin;
  return `${origin}/${encodeUrlState(state)}`;
}

/** Points the oEmbed discovery <link> at our provider for THIS page URL. */
function wireOEmbedDiscovery(link: HTMLLinkElement | null): void {
  if (!link) return;
  const pageUrl = window.location.href;
  link.href = `/api/oembed?format=json&url=${encodeURIComponent(pageUrl)}`;
}

class EmbedViewer {
  private engine: EngineHandles | null = null;
  private activated = false;
  private readonly state: UrlState;

  constructor(
    private readonly dom: EmbedDom,
    state: UrlState,
  ) {
    this.state = state;
  }

  /** Wires the chrome (poster, badge, error link) — no engine yet. */
  init(): void {
    // Badge + error "open in viewer" reopen the SAME model + view in the full app.
    const viewerUrl = openInViewerUrl(this.state);
    this.dom.badge.href = viewerUrl;
    this.dom.errorOpen.href = viewerUrl;

    if (!this.state.modelUrl) {
      // Nothing to show — surface a friendly state, no engine spun up.
      this.showError(t('embed.errorNoModel'));
      this.dom.poster.hidden = true;
      return;
    }

    this.dom.poster.addEventListener('click', () => void this.activate());
  }

  /** Creates the engine (WebGL context) and loads the model. Runs once. */
  private async activate(): Promise<void> {
    if (this.activated) return;
    this.activated = true;
    this.dom.poster.hidden = true;
    this.showLoading(t('embed.loading'), 0.06);

    try {
      const engine = await bootstrapEngine({
        container: this.dom.viewer,
        backgroundColor: DEFAULT_BG,
        gridVisible: false, // chromeless: no grid clutter
        getBackgroundColor: () => DEFAULT_BG,
        // Embed uses the cheapest render path — no PEN postprocessing hook needed.
        getPostproductionRenderer: () => null,
      });
      this.engine = engine;

      // On-demand posture: push a fragments update whenever the camera moves,
      // instead of relying on a heavy continuous postprocessing loop.
      engine.world.camera.controls.addEventListener('update', () => {
        void engine.fragments.core.update();
      });

      const model = await this.loadModel(engine, this.state.modelUrl!);
      await this.addModel(engine, model);

      if (this.state.viewpoint) await this.applyViewpoint(engine, this.state.viewpoint);
      else await this.fit();

      await engine.fragments.core.update(true);

      this.hideLoading();
      this.dom.controls.hidden = false;
    } catch (error) {
      console.error('[embed] load failed', error);
      this.hideLoading();
      this.showError(this.messageForError(error));
    }
  }

  /** Fetches the model bytes (progress) and loads them as .frag or (converting) .ifc. */
  private async loadModel(engine: EngineHandles, url: string): Promise<FragmentsModelLike> {
    const bytes = await this.fetchWithProgress(url);

    // A .frag is opaque fragments bytes (the normal hosting path). An .ifc URL
    // is converted client-side (C2 stays intact — conversion is in the browser).
    const looksIfc = /\.ifc(\?|#|$)/i.test(url) || isProbablyIfc(bytes);
    if (looksIfc) {
      this.showLoading(t('embed.loading'), 0.5);
      const model = (await engine.ifcLoader.load(bytes, true, EMBED_MODEL_ID, {
        processData: {
          progressCallback: (progress: number) => this.setLoadingProgress(0.5 + progress * 0.45),
        },
      })) as unknown as FragmentsModelLike;
      return model;
    }

    this.setLoadingProgress(0.85);
    const model = (await engine.fragments.core.load(bytes, {
      modelId: EMBED_MODEL_ID,
    })) as unknown as FragmentsModelLike;
    return model;
  }

  /** Adds the loaded model to the scene (mirrors the full app's onModelAdded core). */
  private async addModel(engine: EngineHandles, model: FragmentsModelLike): Promise<void> {
    model.useCamera(engine.world.camera.three);
    if (typeof model.graphicsQuality === 'number') model.graphicsQuality = 1;
    engine.world.scene.three.add(model.object);
    await engine.fragments.core.update(true);
  }

  /** Frames the model using the data-driven box (A17: reliable right after load). */
  private async fit(): Promise<void> {
    const engine = this.engine;
    if (!engine) return;
    const model = engine.fragments.list.get(EMBED_MODEL_ID) as unknown as FragmentsModelLike | undefined;
    const box = new THREE.Box3();
    if (model?.box) {
      box.copy(model.box).applyMatrix4(model.object.matrixWorld);
    } else {
      box.setFromObject(engine.world.scene.three);
    }
    if (box.isEmpty()) return;
    const sphere = box.getBoundingSphere(new THREE.Sphere());
    await engine.world.camera.controls.fitToSphere(sphere, true);
  }

  /** Applies the shared viewpoint (camera/projection/section/x-ray/edges/hidden). */
  private async applyViewpoint(engine: EngineHandles, vp: UrlViewpointState): Promise<void> {
    if (vp.camera) {
      if (engine.world.camera.projection.current !== vp.camera.projection) {
        await engine.world.camera.projection.set(vp.camera.projection);
      }
      await engine.world.camera.controls.setLookAt(
        vp.camera.position.x,
        vp.camera.position.y,
        vp.camera.position.z,
        vp.camera.target.x,
        vp.camera.target.y,
        vp.camera.target.z,
        true,
      );
    } else {
      await this.fit();
    }

    // Section planes.
    if (vp.clippingPlanes && vp.clippingPlanes.length > 0) {
      engine.clipper.enabled = true;
      for (const plane of vp.clippingPlanes) {
        engine.clipper.createFromNormalAndCoplanarPoint(
          engine.world,
          new THREE.Vector3(plane.normal.x, plane.normal.y, plane.normal.z),
          new THREE.Vector3(plane.origin.x, plane.origin.y, plane.origin.z),
        );
      }
    }

    // Hidden items (the capped id sample per model).
    if (vp.hidden) {
      const model = engine.fragments.list.get(EMBED_MODEL_ID) as unknown as FragmentsModelLike | undefined;
      if (model) {
        const idsToHide = new Set<number>();
        for (const entry of Object.values(vp.hidden)) {
          for (const id of entry.ids) idsToHide.add(id);
        }
        if (idsToHide.size > 0) {
          await engine.hider.set(false, { [EMBED_MODEL_ID]: idsToHide });
        }
      }
    }

    // X-ray via per-model opacity (cheap: no postprocessing).
    if (vp.xray) {
      const model = engine.fragments.list.get(EMBED_MODEL_ID) as unknown as FragmentsModelLike | undefined;
      await model?.setOpacity(undefined, 0.28);
    }

    await engine.fragments.core.update(true);
  }

  private async toggleFullscreen(): Promise<void> {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await this.dom.root.requestFullscreen();
    } catch (error) {
      console.debug('[embed] fullscreen unavailable', error);
    }
  }

  wireControls(): void {
    this.dom.fitBtn.addEventListener('click', () => void this.fit());
    this.dom.fullscreenBtn.addEventListener('click', () => void this.toggleFullscreen());
    // Keep the canvas sized to the container on resize / fullscreen change.
    window.addEventListener('resize', () => this.onResize());
    document.addEventListener('fullscreenchange', () => this.onResize());
  }

  private onResize(): void {
    const engine = this.engine;
    if (!engine) return;
    // The renderer/camera listen for their own resize; nudge a render so the
    // on-demand embed reflects the new size immediately.
    void engine.fragments.core.update();
  }

  // ── progress / states ──

  private async fetchWithProgress(url: string): Promise<Uint8Array> {
    let response: Response;
    try {
      response = await fetch(url, { mode: 'cors' });
    } catch {
      // fetch() rejects (network down, DNS, CORS/CSP block) — a network error.
      throw new EmbedError('network');
    }
    if (!response.ok) throw new EmbedError(response.status === 404 ? 'expired' : 'network');

    const total = Number(response.headers.get('content-length') ?? 0);
    if (!response.body || !total) {
      // No stream / unknown length — fall back to a single buffer read.
      const buffer = await response.arrayBuffer();
      this.setLoadingProgress(0.45);
      return new Uint8Array(buffer);
    }

    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let received = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      if (value) {
        chunks.push(value);
        received += value.length;
        this.setLoadingProgress(0.06 + (received / total) * 0.4);
      }
    }
    const out = new Uint8Array(received);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.length;
    }
    return out;
  }

  private showLoading(text: string, progress: number): void {
    this.dom.errorBox.hidden = true;
    this.dom.loading.hidden = false;
    this.dom.loadingText.textContent = text;
    this.setLoadingProgress(progress);
  }

  private setLoadingProgress(fraction: number): void {
    const pct = Math.max(2, Math.min(100, Math.round(fraction * 100)));
    this.dom.loadingFill.style.width = `${pct}%`;
  }

  private hideLoading(): void {
    this.dom.loading.hidden = true;
  }

  private showError(message: string): void {
    this.dom.loading.hidden = true;
    this.dom.errorBox.hidden = false;
    this.dom.errorMsg.textContent = message;
  }

  private messageForError(error: unknown): string {
    if (error instanceof EmbedError) {
      if (error.kind === 'expired') return t('embed.errorExpired');
      if (error.kind === 'network') return t('embed.errorNetwork');
    }
    return t('embed.errorGeneric');
  }
}

/** Typed load failure so the UI can distinguish expired (404) from network errors. */
class EmbedError extends Error {
  constructor(readonly kind: 'expired' | 'network') {
    super(kind);
    this.name = 'EmbedError';
  }
}

function bootstrap(): void {
  initLanguage();
  hydrateIcons(document);
  hydrateI18n(document);

  const dom: EmbedDom = {
    root: query('btc-embed-root'),
    viewer: query('embed-viewer'),
    poster: query('embedPoster'),
    loading: query('embedLoading'),
    loadingText: query('embedLoadingText'),
    loadingFill: query('embedLoadingFill'),
    errorBox: query('embedError'),
    errorMsg: query('embedErrorMsg'),
    errorOpen: query('embedErrorOpen'),
    controls: query('embedControls'),
    fitBtn: query('embedFit'),
    fullscreenBtn: query('embedFullscreen'),
    badge: query('embedBadge'),
    oembedLink: document.getElementById('oembedLink') as HTMLLinkElement | null,
  };

  wireOEmbedDiscovery(dom.oembedLink);

  const state = decodeUrlState(window.location.search);
  const viewer = new EmbedViewer(dom, state);
  viewer.wireControls();
  viewer.init();

  if (import.meta.env.VITE_E2E === 'true') {
    (window as unknown as { __embedTestApi?: unknown }).__embedTestApi = {
      version: 1,
      state,
    };
  }
}

bootstrap();
