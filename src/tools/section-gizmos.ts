/**
 * Section gizmos — Autodesk-Viewer-style direct manipulation for sectioning.
 *
 * `SectionPlaneGizmo`: the cutting plane is a model-sized translucent quad
 * with a border; a center handle kit (normal arrow + two rotation rings) sits
 * on it. Dragging the quad or the arrow slides the plane along its normal
 * (clamped to the model bounds); dragging a ring tilts the plane around the
 * corresponding in-plane axis — exactly the Autodesk Viewer section-plane UX.
 *
 * `SectionBoxGizmo`: a translucent box with outlined edges. Hovering a face
 * highlights it and shows the handle kit on it; dragging a face (or the
 * arrow) moves that face along its axis, dragging a ring rotates the whole
 * box about its center — the Autodesk Viewer section-box UX.
 *
 * The gizmos do NOT own any clipping state: they report changes through
 * callbacks and the orchestrator drives the @thatopen clipper planes (which
 * stay the single source of truth for rendering, persistence and viewpoints).
 *
 * Rendering note: every material here is a THREE.ShaderMaterial whose shaders
 * contain no clipping chunks — three.js only injects clipping-plane code into
 * built-in materials (or ShaderMaterials with `clipping: true`), so these
 * visuals are immune to the very clipping planes they control and never get
 * cut by their own section.
 *
 * Pointer notes: the pointerdown listener is registered on `window` with
 * `capture: true`, so it runs before the camera-controls listeners on the
 * canvas; when a handle is hit the event is stopped and the camera never
 * starts an orbit. Plain clicks (no movement) are not consumed — the browser
 * still fires `click`, so click-to-select keeps working through the
 * translucent quad/faces, like Autodesk. After a real drag the synthesized
 * click is swallowed so drag-end does not change the selection.
 */
import * as THREE from 'three';
import {
  axisDragOffset,
  axisRangeOfBox,
  intersectRayPlane,
  moveBoxFace,
  orientedSectionBoxPlanes,
  planeBasis,
  planeQuadRect,
  sectionBoxFaceCenter,
  sectionBoxFaceNormal,
  signedAngleAroundAxis,
  type OrientedSectionBox,
  type SectionRay,
} from './section';

/** Engine handles + render hooks the orchestrator lends to a gizmo. */
export interface SectionGizmoContext {
  scene: THREE.Scene;
  /** The ACTIVE camera (persp/ortho swaps at runtime) — read per event. */
  getCamera: () => THREE.PerspectiveCamera | THREE.OrthographicCamera;
  /** The renderer canvas (pointer events + NDC math). */
  domElement: HTMLElement;
  cameraControls: {
    enabled: boolean;
    addEventListener: (type: 'update', cb: () => void) => void;
    removeEventListener: (type: 'update', cb: () => void) => void;
  };
  requestRender: () => void;
  /** Drag lifecycle — the orchestrator wires the R1 render pump + fragment refreshes. */
  onDragStart: () => void;
  onDragMove: () => void;
  onDragEnd: () => void;
  /** Section accent (CSS hex) — defaults to the neutral steel tone when absent. */
  accentColor?: string;
}

interface GizmoHandle {
  id: string;
  hit: THREE.Object3D;
  cursor: string;
}

const VERTEX_SHADER = 'void main() { gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0); }';
const FRAGMENT_SHADER =
  'uniform vec3 uColor; uniform float uOpacity; void main() { gl_FragColor = vec4(uColor, uOpacity); }';

/** Flat translucent material with no clipping chunks (see module docs). */
function gizmoMaterial(color: THREE.ColorRepresentation, opacity: number, depthTest = true): THREE.ShaderMaterial {
  return new THREE.ShaderMaterial({
    uniforms: { uColor: { value: new THREE.Color(color) }, uOpacity: { value: opacity } },
    vertexShader: VERTEX_SHADER,
    fragmentShader: FRAGMENT_SHADER,
    transparent: true,
    depthWrite: false,
    depthTest,
    side: THREE.DoubleSide,
  });
}

function setMaterial(mesh: THREE.Mesh | THREE.Line, color: THREE.ColorRepresentation, opacity: number): void {
  const material = mesh.material as THREE.ShaderMaterial;
  (material.uniforms.uColor.value as THREE.Color).set(color);
  material.uniforms.uOpacity.value = opacity;
}

const NEUTRAL_ACCENT = '#7d95ad';
const HANDLE_RADIUS = 0.55;
/** Perspective gizmo size: world units per unit camera distance. */
const SCALE_PER_DISTANCE = 1 / 9;
const DRAG_CLICK_SUPPRESS_PX = 4;

const FACE_HANDLE_PATTERN = /^face-([0-5])$/;
const FACE_POINT_PATTERN = /^box-face-([0-5])(-out)?$/;

interface HandleKit {
  group: THREE.Group;
  arrow: THREE.Group;
  ringA: THREE.Mesh;
  ringB: THREE.Mesh;
  arrowHit: THREE.Mesh;
  ringAHit: THREE.Mesh;
  ringBHit: THREE.Mesh;
}

/**
 * The shared handle kit: an arrow along local +Z plus two rotation rings —
 * ringA rotates about local X (ring spans the Y/Z plane), ringB about local Y.
 * Invisible fat hit-proxies (`material.visible = false` keeps them raycastable
 * without rendering) make the thin visuals grabbable on touch (C6).
 */
function buildHandleKit(accent: THREE.Color): HandleKit {
  const group = new THREE.Group();

  const arrow = new THREE.Group();
  const shaft = new THREE.Mesh(new THREE.CylinderGeometry(0.02, 0.02, 1, 12), gizmoMaterial(accent, 0.95, false));
  shaft.rotation.x = Math.PI / 2;
  shaft.position.z = 0.5;
  const head = new THREE.Mesh(new THREE.ConeGeometry(0.07, 0.2, 16), gizmoMaterial(accent, 0.95, false));
  head.rotation.x = Math.PI / 2;
  head.position.z = 1.1;
  arrow.add(shaft, head);

  const ringA = new THREE.Mesh(new THREE.TorusGeometry(HANDLE_RADIUS, 0.015, 8, 48), gizmoMaterial(accent, 0.85, false));
  ringA.rotation.y = Math.PI / 2;
  const ringB = new THREE.Mesh(new THREE.TorusGeometry(HANDLE_RADIUS, 0.015, 8, 48), gizmoMaterial(accent, 0.85, false));
  ringB.rotation.x = Math.PI / 2;

  const hitMaterial = (): THREE.ShaderMaterial => {
    const material = gizmoMaterial(0xffffff, 0);
    material.visible = false;
    return material;
  };
  const arrowHit = new THREE.Mesh(new THREE.CylinderGeometry(0.14, 0.14, 1.4, 8), hitMaterial());
  arrowHit.rotation.x = Math.PI / 2;
  arrowHit.position.z = 0.65;
  const ringAHit = new THREE.Mesh(new THREE.TorusGeometry(HANDLE_RADIUS, 0.09, 6, 24), hitMaterial());
  ringAHit.rotation.y = Math.PI / 2;
  const ringBHit = new THREE.Mesh(new THREE.TorusGeometry(HANDLE_RADIUS, 0.09, 6, 24), hitMaterial());
  ringBHit.rotation.x = Math.PI / 2;

  for (const part of [arrow, ringA, ringB, arrowHit, ringAHit, ringBHit]) {
    part.renderOrder = 999;
    part.traverse((child) => {
      child.renderOrder = 999;
    });
    group.add(part);
  }
  return { group, arrow, ringA, ringB, arrowHit, ringAHit, ringBHit };
}

function disposeObject(root: THREE.Object3D): void {
  root.traverse((child) => {
    const mesh = child as THREE.Mesh;
    if (mesh.geometry) mesh.geometry.dispose();
    const material = mesh.material as THREE.Material | THREE.Material[] | undefined;
    if (Array.isArray(material)) material.forEach((m) => m.dispose());
    else material?.dispose();
  });
}

/**
 * Base gizmo: scene attachment, capture-phase pointer plumbing (hover +
 * drag lifecycle with camera-controls lockout and click suppression), camera
 * driven auto-scaling and disposal. Subclasses define the handles and what a
 * drag does.
 */
abstract class SectionGizmoBase {
  protected readonly ctx: SectionGizmoContext;
  protected readonly group = new THREE.Group();
  protected readonly accent: THREE.Color;
  protected handles: GizmoHandle[] = [];

  private readonly raycaster = new THREE.Raycaster();
  private hovered: GizmoHandle | null = null;
  private draggingHandle: GizmoHandle | null = null;
  private controlsEnabledBeforeDrag = true;
  private dragStartClient = { x: 0, y: 0 };
  private dragMovedPx = 0;
  private previousCursor = '';
  private disposed = false;
  private visibleState = true;

  constructor(ctx: SectionGizmoContext) {
    this.ctx = ctx;
    this.accent = new THREE.Color(ctx.accentColor || NEUTRAL_ACCENT);
    this.group.name = 'section-gizmo';
    ctx.scene.add(this.group);
    window.addEventListener('pointerdown', this.onPointerDown, true);
    ctx.domElement.addEventListener('pointermove', this.onHoverMove);
    ctx.cameraControls.addEventListener('update', this.onCameraUpdate);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    if (this.draggingHandle) this.finishDrag();
    window.removeEventListener('pointerdown', this.onPointerDown, true);
    this.ctx.domElement.removeEventListener('pointermove', this.onHoverMove);
    this.ctx.cameraControls.removeEventListener('update', this.onCameraUpdate);
    this.setHover(null);
    this.ctx.scene.remove(this.group);
    disposeObject(this.group);
    this.ctx.requestRender();
  }

  /** Whether the gizmo visuals are shown (the clipping itself is not affected). */
  get visible(): boolean {
    return this.visibleState;
  }

  /**
   * Shows/hides the gizmo visuals AND its pointer interactions, leaving the
   * clip planes untouched — the Autodesk "deactivate the section tool, keep
   * the cut" mode for inspecting the model unobstructed.
   */
  setVisible(visible: boolean): void {
    if (this.visibleState === visible) return;
    this.visibleState = visible;
    this.group.visible = visible;
    if (!visible) {
      if (this.draggingHandle) this.finishDrag();
      this.setHover(null);
    }
    this.ctx.requestRender();
  }

  /** World position of a named handle (test seam for e2e pointer drags). */
  abstract handleWorldPoint(id: string): THREE.Vector3 | null;

  protected abstract beginDrag(handle: GizmoHandle, hitPoint: THREE.Vector3, ray: SectionRay): boolean;
  protected abstract updateDrag(handle: GizmoHandle, ray: SectionRay): void;
  protected abstract applyHighlight(handle: GizmoHandle | null): void;
  protected abstract updateScale(): void;
  /** Optional subclass hook fired when a handle is hovered (e.g. box face focus). */
  protected onHandleHovered(_handle: GizmoHandle | null): void {}

  protected updateWorldMatrices(): void {
    this.group.updateMatrixWorld(true);
  }

  /** Gizmo scale factor at `anchor` for the current camera (persp + ortho). */
  protected scaleAt(anchor: THREE.Vector3): number {
    const camera = this.ctx.getCamera();
    if ((camera as THREE.PerspectiveCamera).isPerspectiveCamera) {
      const cameraPosition = camera.getWorldPosition(new THREE.Vector3());
      return Math.max(cameraPosition.distanceTo(anchor) * SCALE_PER_DISTANCE, 1e-4);
    }
    const ortho = camera as THREE.OrthographicCamera;
    const height = (ortho.top - ortho.bottom) / (ortho.zoom || 1);
    return Math.max(height * SCALE_PER_DISTANCE, 1e-4);
  }

  private pointerRay(event: PointerEvent): SectionRay | null {
    const rect = this.ctx.domElement.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return null;
    const ndc = new THREE.Vector2(
      ((event.clientX - rect.left) / rect.width) * 2 - 1,
      -(((event.clientY - rect.top) / rect.height) * 2 - 1),
    );
    this.raycaster.setFromCamera(ndc, this.ctx.getCamera());
    return { origin: this.raycaster.ray.origin.clone(), direction: this.raycaster.ray.direction.clone() };
  }

  private pick(event: PointerEvent): { handle: GizmoHandle; point: THREE.Vector3; ray: SectionRay } | null {
    const ray = this.pointerRay(event);
    if (!ray) return null;
    this.updateWorldMatrices();
    let best: { handle: GizmoHandle; point: THREE.Vector3; distance: number } | null = null;
    for (const handle of this.handles) {
      const hits = this.raycaster.intersectObject(handle.hit, false);
      if (hits.length > 0 && (!best || hits[0].distance < best.distance)) {
        best = { handle, point: hits[0].point.clone(), distance: hits[0].distance };
      }
    }
    return best ? { handle: best.handle, point: best.point, ray } : null;
  }

  private setCursor(cursor: string): void {
    this.ctx.domElement.style.cursor = cursor;
  }

  private setHover(handle: GizmoHandle | null): void {
    if (this.hovered === handle) return;
    this.hovered = handle;
    this.applyHighlight(handle);
    this.onHandleHovered(handle);
    if (handle) {
      this.previousCursor = '';
      this.setCursor(handle.cursor);
    } else {
      this.setCursor(this.previousCursor);
    }
    this.ctx.requestRender();
  }

  private readonly onHoverMove = (event: PointerEvent): void => {
    if (this.disposed || !this.visibleState || this.draggingHandle) return;
    const picked = this.pick(event);
    this.setHover(picked ? picked.handle : null);
  };

  private readonly onPointerDown = (event: PointerEvent): void => {
    if (this.disposed || !this.visibleState || this.draggingHandle) return;
    if (event.target !== this.ctx.domElement) return;
    if (event.pointerType === 'mouse' && event.button !== 0) return;
    const picked = this.pick(event);
    if (!picked) return;
    if (!this.beginDrag(picked.handle, picked.point, picked.ray)) return;
    // Consume the event so camera-controls never starts an orbit from it.
    event.stopImmediatePropagation();
    this.draggingHandle = picked.handle;
    this.dragStartClient = { x: event.clientX, y: event.clientY };
    this.dragMovedPx = 0;
    this.controlsEnabledBeforeDrag = this.ctx.cameraControls.enabled;
    this.ctx.cameraControls.enabled = false;
    this.setCursor('grabbing');
    window.addEventListener('pointermove', this.onDragMove, true);
    window.addEventListener('pointerup', this.onDragUp, true);
    window.addEventListener('pointercancel', this.onDragUp, true);
    this.ctx.onDragStart();
    this.ctx.requestRender();
  };

  private readonly onDragMove = (event: PointerEvent): void => {
    if (!this.draggingHandle) return;
    event.stopImmediatePropagation();
    this.dragMovedPx = Math.max(
      this.dragMovedPx,
      Math.hypot(event.clientX - this.dragStartClient.x, event.clientY - this.dragStartClient.y),
    );
    const ray = this.pointerRay(event);
    if (!ray) return;
    this.updateDrag(this.draggingHandle, ray);
    this.ctx.onDragMove();
    this.ctx.requestRender();
  };

  private readonly onDragUp = (): void => {
    if (!this.draggingHandle) return;
    const moved = this.dragMovedPx > DRAG_CLICK_SUPPRESS_PX;
    this.finishDrag();
    if (moved) this.swallowNextClick();
  };

  private finishDrag(): void {
    this.draggingHandle = null;
    window.removeEventListener('pointermove', this.onDragMove, true);
    window.removeEventListener('pointerup', this.onDragUp, true);
    window.removeEventListener('pointercancel', this.onDragUp, true);
    this.ctx.cameraControls.enabled = this.controlsEnabledBeforeDrag;
    this.setCursor(this.previousCursor);
    this.ctx.onDragEnd();
    this.ctx.requestRender();
  }

  /**
   * A drag ends with the browser synthesizing a `click` on the canvas, which
   * would run the app's click-to-select and clear/steal the selection. Swallow
   * exactly that one click (capture phase); a 150 ms timeout covers pointerups
   * that produce no click at all.
   */
  private swallowNextClick(): void {
    const swallow = (event: MouseEvent): void => {
      event.stopImmediatePropagation();
      event.preventDefault();
      cleanup();
    };
    const cleanup = (): void => {
      window.removeEventListener('click', swallow, true);
      clearTimeout(timer);
    };
    window.addEventListener('click', swallow, true);
    const timer = setTimeout(cleanup, 150);
  }

  private readonly onCameraUpdate = (): void => {
    if (this.disposed) return;
    this.updateScale();
  };
}

// ---------------------------------------------------------------------------
// Section plane
// ---------------------------------------------------------------------------

export interface SectionPlaneGizmoOptions {
  normal: THREE.Vector3;
  origin: THREE.Vector3;
  /** Model bounds — sizes the quad and clamps the travel. */
  bounds: THREE.Box3;
  /** Fired on every user mutation; the orchestrator drives the clipper plane. */
  onChange: (normal: THREE.Vector3, origin: THREE.Vector3, kind: 'translate' | 'rotate') => void;
}

interface PlaneDragState {
  mode: 'translate' | 'rotate';
  startOrigin: THREE.Vector3;
  startT: number;
  travel: { lo: number; hi: number };
  axis: THREE.Vector3;
  startVec: THREE.Vector3;
  startNormal: THREE.Vector3;
  startU: THREE.Vector3;
  startV: THREE.Vector3;
  pivot: THREE.Vector3;
}

export class SectionPlaneGizmo extends SectionGizmoBase {
  private readonly options: SectionPlaneGizmoOptions;
  private normal: THREE.Vector3;
  private origin: THREE.Vector3;
  private u: THREE.Vector3;
  private v: THREE.Vector3;
  private readonly bounds: THREE.Box3;

  private readonly planeHolder = new THREE.Group();
  private readonly quad: THREE.Mesh;
  private readonly border: THREE.LineLoop;
  private readonly centerHolder = new THREE.Group();
  private readonly kit: HandleKit;
  private drag: PlaneDragState | null = null;

  constructor(ctx: SectionGizmoContext, options: SectionPlaneGizmoOptions) {
    super(ctx);
    this.options = options;
    this.normal = options.normal.clone().normalize();
    this.origin = options.origin.clone();
    this.bounds = options.bounds.clone();
    const basis = planeBasis(this.normal);
    this.u = basis.u;
    this.v = basis.v;

    this.quad = new THREE.Mesh(new THREE.PlaneGeometry(1, 1), gizmoMaterial(this.accent, 0.08));
    this.quad.renderOrder = 997;
    const borderGeometry = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(-0.5, -0.5, 0),
      new THREE.Vector3(0.5, -0.5, 0),
      new THREE.Vector3(0.5, 0.5, 0),
      new THREE.Vector3(-0.5, 0.5, 0),
    ]);
    this.border = new THREE.LineLoop(borderGeometry, gizmoMaterial(this.accent, 0.9, false));
    this.border.renderOrder = 998;
    this.planeHolder.add(this.quad, this.border);

    this.kit = buildHandleKit(this.accent);
    this.centerHolder.add(this.kit.group);
    this.group.add(this.planeHolder, this.centerHolder);

    this.handles = [
      { id: 'arrow', hit: this.kit.arrowHit, cursor: 'grab' },
      { id: 'ring-u', hit: this.kit.ringAHit, cursor: 'grab' },
      { id: 'ring-v', hit: this.kit.ringBHit, cursor: 'grab' },
      { id: 'quad', hit: this.quad, cursor: 'grab' },
    ];

    this.refresh();
    this.updateScale();
    this.ctx.requestRender();
  }

  getNormal(): THREE.Vector3 {
    return this.normal.clone();
  }

  getOrigin(): THREE.Vector3 {
    return this.origin.clone();
  }

  /** External sync (glass slider / restore): reposition without firing onChange. */
  setFromPlane(normal: THREE.Vector3, origin: THREE.Vector3): void {
    const newNormal = normal.clone().normalize();
    if (!newNormal.equals(this.normal)) {
      this.normal = newNormal;
      const basis = planeBasis(this.normal);
      this.u = basis.u;
      this.v = basis.v;
    }
    this.origin = origin.clone();
    this.refresh();
    this.updateScale();
    this.ctx.requestRender();
  }

  handleWorldPoint(id: string): THREE.Vector3 | null {
    const scale = this.scaleAt(this.origin);
    switch (id) {
      case 'plane-arrow':
        return this.origin.clone().addScaledVector(this.normal, 0.6 * scale);
      case 'plane-arrow-tip':
        return this.origin.clone().addScaledVector(this.normal, 1.2 * scale);
      case 'plane-ring-u':
        return this.origin.clone().addScaledVector(this.v, HANDLE_RADIUS * scale);
      case 'plane-ring-u-swept':
        // The ring-u grab point rotated ~20° about u — a screen-space drag
        // target that produces a measurable rotation in e2e.
        return this.origin
          .clone()
          .addScaledVector(this.v, HANDLE_RADIUS * scale * Math.cos(0.35))
          .addScaledVector(this.normal, HANDLE_RADIUS * scale * Math.sin(0.35));
      case 'plane-quad': {
        const rect = planeQuadRect(this.bounds, this.origin, this.u, this.v);
        if (!rect) return null;
        return rect.center.clone().addScaledVector(this.u, rect.halfU * 0.55);
      }
      default:
        return null;
    }
  }

  protected beginDrag(handle: GizmoHandle, _hitPoint: THREE.Vector3, ray: SectionRay): boolean {
    if (handle.id === 'arrow' || handle.id === 'quad') {
      const startT = axisDragOffset(this.origin, this.normal, ray);
      if (startT === null) return false;
      this.drag = {
        mode: 'translate',
        startOrigin: this.origin.clone(),
        startT,
        travel: axisRangeOfBox(this.bounds, this.normal),
        axis: this.normal.clone(),
        startVec: new THREE.Vector3(),
        startNormal: this.normal.clone(),
        startU: this.u.clone(),
        startV: this.v.clone(),
        pivot: this.origin.clone(),
      };
      return true;
    }
    const axis = handle.id === 'ring-u' ? this.u.clone() : this.v.clone();
    const pivot = this.origin.clone();
    const hit = intersectRayPlane(ray, pivot, axis);
    if (!hit) return false;
    const startVec = hit.sub(pivot);
    if (startVec.lengthSq() < 1e-10) return false;
    this.drag = {
      mode: 'rotate',
      startOrigin: this.origin.clone(),
      startT: 0,
      travel: { lo: 0, hi: 0 },
      axis,
      startVec,
      startNormal: this.normal.clone(),
      startU: this.u.clone(),
      startV: this.v.clone(),
      pivot,
    };
    return true;
  }

  protected updateDrag(_handle: GizmoHandle, ray: SectionRay): void {
    const drag = this.drag;
    if (!drag) return;
    if (drag.mode === 'translate') {
      const t = axisDragOffset(drag.startOrigin, drag.axis, ray);
      if (t === null) return;
      const startDot = drag.axis.dot(drag.startOrigin);
      const dot = Math.min(drag.travel.hi, Math.max(drag.travel.lo, startDot + (t - drag.startT)));
      this.origin = drag.startOrigin.clone().addScaledVector(drag.axis, dot - startDot);
      this.refresh();
      this.options.onChange(this.normal.clone(), this.origin.clone(), 'translate');
      return;
    }
    const hit = intersectRayPlane(ray, drag.pivot, drag.axis);
    if (!hit) return;
    const current = hit.sub(drag.pivot);
    if (current.lengthSq() < 1e-10) return;
    const angle = signedAngleAroundAxis(drag.axis, drag.startVec, current);
    this.normal = drag.startNormal.clone().applyAxisAngle(drag.axis, angle).normalize();
    this.u = drag.startU.clone().applyAxisAngle(drag.axis, angle).normalize();
    this.v = drag.startV.clone().applyAxisAngle(drag.axis, angle).normalize();
    this.refresh();
    this.options.onChange(this.normal.clone(), this.origin.clone(), 'rotate');
  }

  protected applyHighlight(handle: GizmoHandle | null): void {
    setMaterial(this.quad, this.accent, handle?.id === 'quad' ? 0.16 : 0.08);
    const highlight = new THREE.Color(0xffffff).lerp(this.accent, 0.35);
    for (const part of this.kit.arrow.children) {
      setMaterial(part as THREE.Mesh, handle?.id === 'arrow' ? highlight : this.accent, 0.95);
    }
    setMaterial(this.kit.ringA, handle?.id === 'ring-u' ? highlight : this.accent, handle?.id === 'ring-u' ? 1 : 0.85);
    setMaterial(this.kit.ringB, handle?.id === 'ring-v' ? highlight : this.accent, handle?.id === 'ring-v' ? 1 : 0.85);
  }

  protected updateScale(): void {
    const scale = this.scaleAt(this.origin);
    this.centerHolder.scale.setScalar(scale);
    this.updateWorldMatrices();
  }

  private refresh(): void {
    const rect = planeQuadRect(this.bounds, this.origin, this.u, this.v);
    const quaternion = new THREE.Quaternion().setFromRotationMatrix(
      new THREE.Matrix4().makeBasis(this.u, this.v, this.normal),
    );
    if (rect) {
      // Keep the gizmo origin at the quad center: it stays on the plane (any
      // coplanar point defines the same plane) and keeps the handles mid-model.
      this.origin = rect.center;
      this.planeHolder.position.copy(rect.center);
      this.planeHolder.quaternion.copy(quaternion);
      this.quad.scale.set(rect.halfU * 2, rect.halfV * 2, 1);
      this.border.scale.set(rect.halfU * 2, rect.halfV * 2, 1);
    }
    this.centerHolder.position.copy(this.origin);
    this.centerHolder.quaternion.copy(quaternion);
    this.updateWorldMatrices();
  }
}

// ---------------------------------------------------------------------------
// Section box
// ---------------------------------------------------------------------------

export interface SectionBoxGizmoOptions {
  box: OrientedSectionBox;
  /**
   * Fired on every user mutation with the six clip planes (face order — see
   * OrientedSectionBox) and the new box state.
   */
  onChange: (planes: Array<{ normal: THREE.Vector3; point: THREE.Vector3 }>, box: OrientedSectionBox) => void;
}

interface BoxDragState {
  mode: 'face' | 'rotate';
  face: number;
  axisIndex: number;
  sign: 1 | -1;
  axis: THREE.Vector3;
  startT: number;
  startCenter: THREE.Vector3;
  startHalf: number;
  startVec: THREE.Vector3;
  startAxes: [THREE.Vector3, THREE.Vector3, THREE.Vector3];
  pivot: THREE.Vector3;
}

export class SectionBoxGizmo extends SectionGizmoBase {
  private readonly options: SectionBoxGizmoOptions;
  private box: OrientedSectionBox;
  private readonly minThickness: number;

  private readonly boxHolder = new THREE.Group();
  private readonly faces: THREE.Mesh[] = [];
  private readonly edges: THREE.LineSegments;
  private readonly kit: HandleKit;
  private readonly kitHolder = new THREE.Group();
  private activeFace = 0;
  private drag: BoxDragState | null = null;

  constructor(ctx: SectionGizmoContext, options: SectionBoxGizmoOptions) {
    super(ctx);
    this.options = options;
    this.box = {
      center: options.box.center.clone(),
      halfSizes: options.box.halfSizes.clone(),
      axes: [options.box.axes[0].clone(), options.box.axes[1].clone(), options.box.axes[2].clone()],
    };
    this.minThickness = Math.max(this.box.halfSizes.length() * 2 * 0.01, 1e-3);

    for (let face = 0; face < 6; face += 1) {
      const mesh = new THREE.Mesh(new THREE.PlaneGeometry(1, 1), gizmoMaterial(this.accent, 0.05));
      mesh.renderOrder = 997;
      this.faces.push(mesh);
      this.boxHolder.add(mesh);
    }
    this.edges = new THREE.LineSegments(
      new THREE.EdgesGeometry(new THREE.BoxGeometry(1, 1, 1)),
      gizmoMaterial(this.accent, 0.65, false),
    );
    this.edges.renderOrder = 998;
    this.boxHolder.add(this.edges);

    this.kit = buildHandleKit(this.accent);
    this.kitHolder.add(this.kit.group);
    this.group.add(this.boxHolder, this.kitHolder);

    this.handles = [
      ...this.faces.map((mesh, face) => ({ id: `face-${face}`, hit: mesh, cursor: 'grab' })),
      { id: 'arrow', hit: this.kit.arrowHit, cursor: 'grab' },
      { id: 'ring-a', hit: this.kit.ringAHit, cursor: 'grab' },
      { id: 'ring-b', hit: this.kit.ringBHit, cursor: 'grab' },
    ];

    this.refresh();
    this.updateScale();
    this.ctx.requestRender();
  }

  getBox(): OrientedSectionBox {
    return {
      center: this.box.center.clone(),
      halfSizes: this.box.halfSizes.clone(),
      axes: [this.box.axes[0].clone(), this.box.axes[1].clone(), this.box.axes[2].clone()],
    };
  }

  handleWorldPoint(id: string): THREE.Vector3 | null {
    const faceMatch = id.match(FACE_POINT_PATTERN);
    if (faceMatch) {
      const face = Number(faceMatch[1]);
      const center = sectionBoxFaceCenter(this.box, face);
      if (!faceMatch[2]) return center;
      const scale = this.scaleAt(center);
      return center.addScaledVector(sectionBoxFaceNormal(this.box, face), 1.2 * scale);
    }
    if (id === 'box-arrow') {
      const center = sectionBoxFaceCenter(this.box, this.activeFace);
      const scale = this.scaleAt(center);
      return center.addScaledVector(sectionBoxFaceNormal(this.box, this.activeFace), 0.6 * scale);
    }
    return null;
  }

  /** In-plane axes of a face (for the rotation rings + kit orientation). */
  private faceTangents(face: number): { a: THREE.Vector3; b: THREE.Vector3 } {
    const i = Math.floor(face / 2);
    return { a: this.box.axes[(i + 1) % 3].clone(), b: this.box.axes[(i + 2) % 3].clone() };
  }

  protected beginDrag(handle: GizmoHandle, _hitPoint: THREE.Vector3, ray: SectionRay): boolean {
    const faceMatch = handle.id.match(FACE_HANDLE_PATTERN);
    const face = faceMatch ? Number(faceMatch[1]) : this.activeFace;
    const axisIndex = Math.floor(face / 2);
    const sign: 1 | -1 = face % 2 === 0 ? 1 : -1;
    const axis = this.box.axes[axisIndex].clone();

    if (faceMatch || handle.id === 'arrow') {
      if (faceMatch) this.setActiveFace(face);
      const faceCenter = sectionBoxFaceCenter(this.box, face);
      const startT = axisDragOffset(faceCenter, axis, ray);
      if (startT === null) return false;
      const half =
        axisIndex === 0 ? this.box.halfSizes.x : axisIndex === 1 ? this.box.halfSizes.y : this.box.halfSizes.z;
      this.drag = {
        mode: 'face',
        face,
        axisIndex,
        sign,
        axis,
        startT,
        startCenter: this.box.center.clone(),
        startHalf: half,
        startVec: new THREE.Vector3(),
        startAxes: [this.box.axes[0].clone(), this.box.axes[1].clone(), this.box.axes[2].clone()],
        pivot: this.box.center.clone(),
      };
      return true;
    }

    // Ring drag → rotate the whole box about the face's in-plane axis through
    // the box center (the Autodesk section-box rotation).
    const tangents = this.faceTangents(this.activeFace);
    const rotationAxis = handle.id === 'ring-a' ? tangents.a : tangents.b;
    const gizmoAnchor = sectionBoxFaceCenter(this.box, this.activeFace);
    const hit = intersectRayPlane(ray, gizmoAnchor, rotationAxis);
    if (!hit) return false;
    const startVec = hit.sub(gizmoAnchor);
    if (startVec.lengthSq() < 1e-10) return false;
    this.drag = {
      mode: 'rotate',
      face: this.activeFace,
      axisIndex: 0,
      sign: 1,
      axis: rotationAxis,
      startT: 0,
      startCenter: this.box.center.clone(),
      startHalf: 0,
      startVec,
      startAxes: [this.box.axes[0].clone(), this.box.axes[1].clone(), this.box.axes[2].clone()],
      pivot: gizmoAnchor.clone(),
    };
    return true;
  }

  protected updateDrag(_handle: GizmoHandle, ray: SectionRay): void {
    const drag = this.drag;
    if (!drag) return;
    if (drag.mode === 'face') {
      const startCenterA = drag.axis.dot(drag.startCenter);
      const startFaceA = startCenterA + drag.sign * drag.startHalf;
      const faceCenterStart = drag.startCenter.clone().addScaledVector(drag.axis, drag.sign * drag.startHalf);
      const t = axisDragOffset(faceCenterStart, drag.axis, ray);
      if (t === null) return;
      const targetA = startFaceA + (t - drag.startT);
      const moved = moveBoxFace(startCenterA, drag.startHalf, drag.sign, targetA, this.minThickness);
      this.box.center = drag.startCenter.clone().addScaledVector(drag.axis, moved.centerA - startCenterA);
      if (drag.axisIndex === 0) this.box.halfSizes.x = moved.half;
      else if (drag.axisIndex === 1) this.box.halfSizes.y = moved.half;
      else this.box.halfSizes.z = moved.half;
      this.emitChange();
      return;
    }
    const hit = intersectRayPlane(ray, drag.pivot, drag.axis);
    if (!hit) return;
    const current = hit.sub(drag.pivot);
    if (current.lengthSq() < 1e-10) return;
    const angle = signedAngleAroundAxis(drag.axis, drag.startVec, current);
    this.box.axes = [
      drag.startAxes[0].clone().applyAxisAngle(drag.axis, angle).normalize(),
      drag.startAxes[1].clone().applyAxisAngle(drag.axis, angle).normalize(),
      drag.startAxes[2].clone().applyAxisAngle(drag.axis, angle).normalize(),
    ];
    this.emitChange();
  }

  protected applyHighlight(handle: GizmoHandle | null): void {
    const faceMatch = handle ? handle.id.match(FACE_HANDLE_PATTERN) : null;
    const hoveredFace = faceMatch ? Number(faceMatch[1]) : null;
    this.faces.forEach((mesh, face) => {
      setMaterial(mesh, this.accent, face === hoveredFace ? 0.16 : 0.05);
    });
    const highlight = new THREE.Color(0xffffff).lerp(this.accent, 0.35);
    for (const part of this.kit.arrow.children) {
      setMaterial(part as THREE.Mesh, handle?.id === 'arrow' ? highlight : this.accent, 0.95);
    }
    setMaterial(this.kit.ringA, handle?.id === 'ring-a' ? highlight : this.accent, handle?.id === 'ring-a' ? 1 : 0.85);
    setMaterial(this.kit.ringB, handle?.id === 'ring-b' ? highlight : this.accent, handle?.id === 'ring-b' ? 1 : 0.85);
  }

  protected onHandleHovered(handle: GizmoHandle | null): void {
    const faceMatch = handle ? handle.id.match(FACE_HANDLE_PATTERN) : null;
    if (faceMatch) this.setActiveFace(Number(faceMatch[1]));
  }

  protected updateScale(): void {
    const anchor = sectionBoxFaceCenter(this.box, this.activeFace);
    this.kitHolder.scale.setScalar(this.scaleAt(anchor));
    this.updateWorldMatrices();
  }

  private setActiveFace(face: number): void {
    this.activeFace = face;
    this.positionKit();
  }

  private positionKit(): void {
    const face = this.activeFace;
    const normal = sectionBoxFaceNormal(this.box, face);
    const tangents = this.faceTangents(face);
    // Kit local basis: Z = outward normal (arrow), X/Y = the ring axes. Keep
    // the basis right-handed so the quaternion is well-formed.
    const x = tangents.a;
    const y = new THREE.Vector3().crossVectors(normal, x);
    const quaternion = new THREE.Quaternion().setFromRotationMatrix(new THREE.Matrix4().makeBasis(x, y, normal));
    this.kitHolder.position.copy(sectionBoxFaceCenter(this.box, face));
    this.kitHolder.quaternion.copy(quaternion);
    this.updateScale();
  }

  private emitChange(): void {
    this.refresh();
    this.options.onChange(orientedSectionBoxPlanes(this.box), this.getBox());
  }

  private refresh(): void {
    const { center, halfSizes, axes } = this.box;
    this.boxHolder.position.copy(center);
    this.boxHolder.quaternion.setFromRotationMatrix(new THREE.Matrix4().makeBasis(axes[0], axes[1], axes[2]));
    const half = [halfSizes.x, halfSizes.y, halfSizes.z];
    for (let face = 0; face < 6; face += 1) {
      const i = Math.floor(face / 2);
      const s = face % 2 === 0 ? 1 : -1;
      const j = (i + 1) % 3;
      const k = (i + 2) % 3;
      const mesh = this.faces[face];
      const x = new THREE.Vector3();
      const y = new THREE.Vector3();
      const z = new THREE.Vector3();
      z.setComponent(i, s);
      if (s === 1) {
        x.setComponent(j, 1);
        y.setComponent(k, 1);
        mesh.scale.set(half[j] * 2, half[k] * 2, 1);
      } else {
        x.setComponent(k, 1);
        y.setComponent(j, 1);
        mesh.scale.set(half[k] * 2, half[j] * 2, 1);
      }
      mesh.quaternion.setFromRotationMatrix(new THREE.Matrix4().makeBasis(x, y, z));
      mesh.position.set(0, 0, 0).setComponent(i, s * half[i]);
    }
    this.edges.scale.set(halfSizes.x * 2, halfSizes.y * 2, halfSizes.z * 2);
    this.positionKit();
    this.updateWorldMatrices();
  }
}
