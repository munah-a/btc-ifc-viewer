/**
 * Keyboard router (the W3.5-deferred keyboard-router extraction, folded into
 * W5.3). Owns the two pieces of `onKeyDown` that are pure and reusable: the
 * form-target guard (ignore keys typed into inputs/textareas/selects/
 * contenteditable) and the key → action dispatch table. The action bodies stay
 * in the orchestrator (they touch engine + DOM state) and are injected as
 * `KeyboardActions`.
 *
 * DOM-free apart from reading `event.target`/`event.key`; unit-tested with a
 * stub actions object and synthetic events.
 */

/** The named actions the router dispatches to (implemented by the viewer). */
export interface KeyboardActions {
  /** Escape: cancel the active tool / gizmo / open sheet or panel. */
  cancel(): void;
  fitToModel(): void;
  setNavigationMode(mode: 'Orbit' | 'Plan' | 'FirstPerson'): void;
  toggleSelectionMode(): void;
  toggleMeasure(mode: 'length' | 'area'): void;
  toggleGrid(): void;
  toggleXray(): void;
  /** 'e': gizmo rotate when a gizmo is active, else toggle edges. */
  edgesOrGizmoRotate(): void;
  /** 'w': gizmo translate (no-op when no gizmo active). */
  gizmoTranslate(): void;
  /** 'r': reset the active model transform (no-op when no gizmo active). */
  gizmoReset(): void;
  toggleIssuePin(): void;
  deleteSelectedIssue(): void;
  /** Enter: finish an in-progress area measurement. */
  finishAreaMeasurement(): void;
}

/** True when the key event originated in a text-entry control (ignore shortcuts). */
export function isFormFieldTarget(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  return (
    !!el &&
    (el.tagName === 'INPUT' ||
      el.tagName === 'TEXTAREA' ||
      el.tagName === 'SELECT' ||
      el.isContentEditable)
  );
}

/**
 * Dispatches a keydown to the matching action. No-op for keys typed in form
 * fields and for unmapped keys. Behaviour matches the original inline onKeyDown
 * exactly (same keys, same guards).
 */
export function routeKeyboardEvent(event: KeyboardEvent, actions: KeyboardActions): void {
  if (isFormFieldTarget(event.target)) return;

  if (event.key === 'Escape') {
    actions.cancel();
    return;
  }

  switch (event.key.toLowerCase()) {
    case 'f':
      actions.fitToModel();
      break;
    case '1':
      actions.setNavigationMode('Orbit');
      break;
    case '2':
      actions.setNavigationMode('Plan');
      break;
    case '3':
      actions.setNavigationMode('FirstPerson');
      break;
    case 'm':
      actions.toggleSelectionMode();
      break;
    case 'l':
      actions.toggleMeasure('length');
      break;
    case 'a':
      actions.toggleMeasure('area');
      break;
    case 'g':
      actions.toggleGrid();
      break;
    case 'x':
      actions.toggleXray();
      break;
    case 'e':
      actions.edgesOrGizmoRotate();
      break;
    case 'w':
      actions.gizmoTranslate();
      break;
    case 'r':
      actions.gizmoReset();
      break;
    case 'i':
      actions.toggleIssuePin();
      break;
    case 'delete':
      actions.deleteSelectedIssue();
      break;
    case 'enter':
      actions.finishAreaMeasurement();
      break;
    default:
      break;
  }
}
