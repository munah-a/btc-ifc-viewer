import { describe, expect, it, vi } from 'vitest';

import {
  isFormFieldTarget,
  routeKeyboardEvent,
  type KeyboardActions,
} from '../../src/input/keyboard-router';

const makeActions = (): KeyboardActions & { calls: string[] } => {
  const calls: string[] = [];
  const rec = (name: string) => (...args: unknown[]) => {
    calls.push(args.length ? `${name}:${args.join(',')}` : name);
  };
  return {
    calls,
    cancel: rec('cancel'),
    fitToModel: rec('fitToModel'),
    setNavigationMode: rec('setNavigationMode'),
    toggleSelectionMode: rec('toggleSelectionMode'),
    toggleMeasure: rec('toggleMeasure'),
    toggleGrid: rec('toggleGrid'),
    toggleXray: rec('toggleXray'),
    edgesOrGizmoRotate: rec('edgesOrGizmoRotate'),
    gizmoTranslate: rec('gizmoTranslate'),
    gizmoReset: rec('gizmoReset'),
    toggleIssuePin: rec('toggleIssuePin'),
    deleteSelectedIssue: rec('deleteSelectedIssue'),
    finishAreaMeasurement: rec('finishAreaMeasurement'),
  };
};

const evt = (key: string, target: EventTarget | null = null): KeyboardEvent =>
  ({ key, target } as unknown as KeyboardEvent);

describe('isFormFieldTarget (keyboard-router — W5.3)', () => {
  it('detects text-entry controls', () => {
    expect(isFormFieldTarget({ tagName: 'INPUT', isContentEditable: false } as unknown as EventTarget)).toBe(true);
    expect(isFormFieldTarget({ tagName: 'TEXTAREA', isContentEditable: false } as unknown as EventTarget)).toBe(true);
    expect(isFormFieldTarget({ tagName: 'SELECT', isContentEditable: false } as unknown as EventTarget)).toBe(true);
    expect(isFormFieldTarget({ tagName: 'DIV', isContentEditable: true } as unknown as EventTarget)).toBe(true);
  });

  it('ignores non-entry targets and null', () => {
    expect(isFormFieldTarget({ tagName: 'BUTTON', isContentEditable: false } as unknown as EventTarget)).toBe(false);
    expect(isFormFieldTarget(null)).toBe(false);
  });
});

describe('routeKeyboardEvent (keyboard-router — W5.3)', () => {
  it('does nothing when the target is a form field', () => {
    const actions = makeActions();
    routeKeyboardEvent(evt('f', { tagName: 'INPUT', isContentEditable: false } as unknown as EventTarget), actions);
    expect(actions.calls).toEqual([]);
  });

  it('routes Escape to cancel', () => {
    const actions = makeActions();
    routeKeyboardEvent(evt('Escape'), actions);
    expect(actions.calls).toEqual(['cancel']);
  });

  it('maps navigation number keys', () => {
    const actions = makeActions();
    routeKeyboardEvent(evt('1'), actions);
    routeKeyboardEvent(evt('2'), actions);
    routeKeyboardEvent(evt('3'), actions);
    expect(actions.calls).toEqual([
      'setNavigationMode:Orbit',
      'setNavigationMode:Plan',
      'setNavigationMode:FirstPerson',
    ]);
  });

  it('maps measure/grid/xray/fit and is case-insensitive', () => {
    const actions = makeActions();
    routeKeyboardEvent(evt('F'), actions);
    routeKeyboardEvent(evt('L'), actions);
    routeKeyboardEvent(evt('A'), actions);
    routeKeyboardEvent(evt('G'), actions);
    routeKeyboardEvent(evt('X'), actions);
    routeKeyboardEvent(evt('M'), actions);
    expect(actions.calls).toEqual([
      'fitToModel',
      'toggleMeasure:length',
      'toggleMeasure:area',
      'toggleGrid',
      'toggleXray',
      'toggleSelectionMode',
    ]);
  });

  it('maps gizmo + issue + delete + enter keys', () => {
    const actions = makeActions();
    routeKeyboardEvent(evt('e'), actions);
    routeKeyboardEvent(evt('w'), actions);
    routeKeyboardEvent(evt('r'), actions);
    routeKeyboardEvent(evt('i'), actions);
    routeKeyboardEvent(evt('Delete'), actions);
    routeKeyboardEvent(evt('Enter'), actions);
    expect(actions.calls).toEqual([
      'edgesOrGizmoRotate',
      'gizmoTranslate',
      'gizmoReset',
      'toggleIssuePin',
      'deleteSelectedIssue',
      'finishAreaMeasurement',
    ]);
  });

  it('ignores unmapped keys', () => {
    const actions = makeActions();
    routeKeyboardEvent(evt('q'), actions);
    routeKeyboardEvent(evt('Tab'), actions);
    expect(actions.calls).toEqual([]);
  });

  it('does not call actions beyond the dispatched one', () => {
    const actions = makeActions();
    const spy = vi.spyOn(actions, 'fitToModel');
    routeKeyboardEvent(evt('f'), actions);
    expect(spy).toHaveBeenCalledTimes(1);
  });
});
