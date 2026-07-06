import { describe, expect, it } from 'vitest';

import { normalizePersistedState } from '../../src/core/persistence';

describe('normalizePersistedState (AUDIT A7 — W1.6)', () => {
  it('rejects values that are not a v1 state', () => {
    expect(normalizePersistedState(null)).toBeNull();
    expect(normalizePersistedState(undefined)).toBeNull();
    expect(normalizePersistedState('viewer')).toBeNull();
    expect(normalizePersistedState([])).toBeNull();
    expect(normalizePersistedState({})).toBeNull();
    expect(normalizePersistedState({ version: 2 })).toBeNull();
  });

  it('accepts the minimal `{"version":1}` import crash-free with full defaults', () => {
    // The old importViewerState crashed on parsed.viewpoints being undefined.
    const state = normalizePersistedState(JSON.parse('{"version":1}'));
    expect(state).not.toBeNull();
    expect(state).toMatchObject({
      version: 1,
      selectionMode: 'single',
      navigationMode: 'Orbit',
      visualStyle: 'color-pen-shadows',
      xray: false,
      edges: false,
      gridVisible: false,
      theme: 'dark',
      viewpoints: [],
      issues: [],
    });
  });

  it('coerces malformed field types to defaults instead of crashing', () => {
    const state = normalizePersistedState({
      version: 1,
      selectionMode: 'sneaky',
      navigationMode: 42,
      visualStyle: 'xray', // dead legacy value (pre-A7 fallback) — not a style
      xray: 'yes',
      theme: 'hotdog',
      backgroundColor: 'javascript:alert(1)',
      backgroundByTheme: { dark: '#0b1220', light: 'red' },
      viewpoints: 'nope',
      issues: { 0: {} },
    });
    expect(state).not.toBeNull();
    expect(state?.selectionMode).toBe('single');
    expect(state?.navigationMode).toBe('Orbit');
    expect(state?.visualStyle).toBe('color-pen-shadows');
    expect(state?.xray).toBe(false);
    expect(state?.theme).toBe('dark');
    expect(state?.backgroundColor).toBeUndefined();
    expect(state?.backgroundByTheme).toBeUndefined(); // partial pair dropped
    expect(state?.viewpoints).toEqual([]);
    expect(state?.issues).toEqual([]);
  });

  it('drops malformed viewpoints and keeps valid ones', () => {
    const valid = {
      id: 'vp1',
      name: 'North wing',
      createdAt: '2026-07-06T00:00:00.000Z',
      camera: {
        position: { x: 1, y: 2, z: 3 },
        target: { x: 0, y: 0, z: 0 },
        projection: 'Orthographic',
        mode: 'Plan',
      },
      clippingPlanes: [
        { normal: { x: 0, y: -1, z: 0 }, origin: { x: 0, y: 3, z: 0 } },
        { normal: null, origin: { x: 0, y: 0, z: 0 } }, // dropped
      ],
      hiddenItems: { 'a.ifc': [1, 2, 'x'], 'b.ifc': [] },
      xray: true,
      edges: false,
      snapshot: 'data:image/jpeg;base64,AAAA',
    };
    const state = normalizePersistedState({
      version: 1,
      viewpoints: [valid, { id: 'no-camera', name: 'broken' }, 'garbage', null],
      issues: [],
    });
    expect(state?.viewpoints).toHaveLength(1);
    const viewpoint = state?.viewpoints[0];
    expect(viewpoint?.camera.projection).toBe('Orthographic');
    expect(viewpoint?.camera.mode).toBe('Plan');
    expect(viewpoint?.clippingPlanes).toHaveLength(1);
    expect(viewpoint?.hiddenItems).toEqual({ 'a.ifc': [1, 2] });
    expect(viewpoint?.xray).toBe(true);
    expect(viewpoint?.snapshot).toBe('data:image/jpeg;base64,AAAA');
  });

  it('rejects non-image snapshots (no javascript: smuggling)', () => {
    const state = normalizePersistedState({
      version: 1,
      viewpoints: [{
        id: 'vp1',
        name: 'v',
        camera: { position: { x: 0, y: 0, z: 0 }, target: { x: 0, y: 0, z: 0 } },
        snapshot: 'javascript:alert(1)',
      }],
      issues: [],
    });
    expect(state?.viewpoints[0]?.snapshot).toBeUndefined();
  });

  it('normalizes issues: enums coerced, ids filtered, comments validated', () => {
    const state = normalizePersistedState({
      version: 1,
      viewpoints: [],
      issues: [
        {
          id: 'i1',
          title: 'Clash at grid B2',
          priority: 'Urgent', // invalid → Medium
          status: 'Weird', // invalid → Open
          modelId: 'a.ifc',
          localIds: [1, 'x', 2.5, Infinity],
          elementsByModel: { 'a.ifc': [1], 'b.ifc': 'bad' },
          point: { x: 1, y: 2, z: 'three' }, // invalid → null
          comments: [{ id: 'c1', text: 'check', author: 'QA' }, { id: '', text: 'no-id' }, 42],
        },
        { title: 'missing id' },
      ],
    });
    expect(state?.issues).toHaveLength(1);
    const issue = state?.issues[0];
    expect(issue?.priority).toBe('Medium');
    expect(issue?.status).toBe('Open');
    expect(issue?.localIds).toEqual([1, 2.5]);
    expect(issue?.elementsByModel).toEqual({ 'a.ifc': [1] });
    expect(issue?.point).toBeNull();
    expect(issue?.comments).toHaveLength(1);
    expect(issue?.comments[0]?.author).toBe('QA');
  });

  it('carries a valid language and drops an invalid/absent one (C7)', () => {
    expect(normalizePersistedState({ version: 1 })?.language).toBeUndefined();
    expect(normalizePersistedState({ version: 1, language: 'de' })?.language).toBe('de');
    expect(normalizePersistedState({ version: 1, language: 'en' })?.language).toBe('en');
    // Invalid string coerces to the 'en' default; non-string stays undefined.
    expect(normalizePersistedState({ version: 1, language: 'fr' })?.language).toBe('en');
    expect(normalizePersistedState({ version: 1, language: 42 })?.language).toBeUndefined();
  });

  it('round-trips a well-formed state unchanged in the fields that matter', () => {
    const input = {
      version: 1,
      selectionMode: 'multi',
      navigationMode: 'FirstPerson',
      visualStyle: 'pen',
      xray: true,
      edges: true,
      gridVisible: true,
      backgroundColor: '#8EC0F4',
      backgroundByTheme: { dark: '#1b1f24', light: '#e9ecef' },
      theme: 'light',
      viewpoints: [],
      issues: [],
    };
    const state = normalizePersistedState(input);
    expect(state).toMatchObject({
      selectionMode: 'multi',
      navigationMode: 'FirstPerson',
      visualStyle: 'pen',
      xray: true,
      edges: true,
      gridVisible: true,
      backgroundColor: '#8ec0f4',
      backgroundByTheme: { dark: '#1b1f24', light: '#e9ecef' },
      theme: 'light',
    });
  });
});
