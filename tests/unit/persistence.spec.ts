import { describe, expect, it } from 'vitest';

import {
  buildPersistedState,
  normalizePersistedState,
  type PersistedStateInput,
} from '../../src/core/persistence';

describe('normalizePersistedState (AUDIT A7 — W1.6; C8 v2 — W5.2)', () => {
  it('rejects values that are not a recognizable state', () => {
    expect(normalizePersistedState(null)).toBeNull();
    expect(normalizePersistedState(undefined)).toBeNull();
    expect(normalizePersistedState('viewer')).toBeNull();
    expect(normalizePersistedState([])).toBeNull();
    expect(normalizePersistedState({})).toBeNull();
    expect(normalizePersistedState({ version: 3 })).toBeNull();
    expect(normalizePersistedState({ version: 0 })).toBeNull();
  });

  it('accepts the minimal `{"version":2}` import crash-free with full defaults', () => {
    const state = normalizePersistedState(JSON.parse('{"version":2}'));
    expect(state).not.toBeNull();
    expect(state).toMatchObject({
      version: 2,
      selectionMode: 'single',
      navigationMode: 'Orbit',
      visualStyle: 'color-pen-shadows',
      xray: false,
      edges: false,
      gridVisible: false,
      theme: 'dark',
      viewpoints: [],
      issues: [],
      models: [],
      sectionPlanes: [],
      selection: {},
    });
  });

  it('migrates a legacy v1 blob to v2 with empty C8 defaults (W5.2)', () => {
    // v1 had no models/camera/section/selection — those default to empty and
    // the version is bumped to 2, so an existing user's saved state still loads.
    const state = normalizePersistedState({ version: 1, xray: true, theme: 'light' });
    expect(state).not.toBeNull();
    expect(state?.version).toBe(2);
    expect(state?.xray).toBe(true);
    expect(state?.theme).toBe('light');
    expect(state?.models).toEqual([]);
    expect(state?.camera).toBeUndefined();
    expect(state?.sectionPlanes).toEqual([]);
    expect(state?.selection).toEqual({});
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

describe('normalizePersistedState — C8 full-session fields (W5.2)', () => {
  it('normalizes model records; drops records without an identity or fragKey', () => {
    const state = normalizePersistedState({
      version: 2,
      models: [
        {
          modelId: 'a.ifc',
          fileName: 'A.ifc',
          sizeBytes: 1024,
          fragKey: 'deadbeef-400',
          offsetPosition: { x: 1, y: 2, z: 3 },
          offsetRotation: { x: 0, y: 90, z: 0 },
          opacity: 0.5,
          visible: false,
          hiddenIds: [10, 20, 'x'],
        },
        { modelId: 'no-key.ifc' }, // dropped: no fragKey
        { fragKey: 'orphan' }, // dropped: no modelId
        'garbage',
      ],
    });
    expect(state?.models).toHaveLength(1);
    const model = state?.models[0];
    expect(model?.modelId).toBe('a.ifc');
    expect(model?.fragKey).toBe('deadbeef-400');
    expect(model?.offsetPosition).toEqual({ x: 1, y: 2, z: 3 });
    expect(model?.offsetRotation).toEqual({ x: 0, y: 90, z: 0 });
    expect(model?.opacity).toBe(0.5);
    expect(model?.visible).toBe(false);
    expect(model?.hiddenIds).toEqual([10, 20]); // non-number dropped
  });

  it('clamps opacity and defaults missing per-model fields', () => {
    const state = normalizePersistedState({
      version: 2,
      models: [{ modelId: 'm', fragKey: 'k', opacity: 5 }],
    });
    const model = state?.models[0];
    expect(model?.opacity).toBe(1); // clamped
    expect(model?.visible).toBe(true); // default
    expect(model?.offsetPosition).toEqual({ x: 0, y: 0, z: 0 });
    expect(model?.hiddenIds).toEqual([]);
  });

  it('restores a valid camera and drops an incomplete one', () => {
    const good = normalizePersistedState({
      version: 2,
      camera: { position: { x: 1, y: 2, z: 3 }, target: { x: 0, y: 0, z: 0 }, projection: 'Orthographic' },
    });
    expect(good?.camera).toEqual({
      position: { x: 1, y: 2, z: 3 },
      target: { x: 0, y: 0, z: 0 },
      projection: 'Orthographic',
    });
    const bad = normalizePersistedState({ version: 2, camera: { position: { x: 1, y: 2, z: 3 } } });
    expect(bad?.camera).toBeUndefined();
  });

  it('normalizes section planes and selection, dropping malformed entries', () => {
    const state = normalizePersistedState({
      version: 2,
      sectionPlanes: [
        { normal: { x: 0, y: 1, z: 0 }, origin: { x: 0, y: 5, z: 0 } },
        { normal: null, origin: { x: 0, y: 0, z: 0 } }, // dropped
      ],
      selection: { 'a.ifc': [1, 2, 'x'], 'b.ifc': [], 'c.ifc': 'nope' },
    });
    expect(state?.sectionPlanes).toHaveLength(1);
    expect(state?.selection).toEqual({ 'a.ifc': [1, 2] }); // empty + non-array dropped
  });

  it('normalizes search sets, dropping malformed entries (2026-07-12)', () => {
    const state = normalizePersistedState({
      version: 2,
      searchSets: [
        {
          id: 's1',
          name: 'walls',
          query: 'wall',
          color: '#123456',
          colorActive: true,
          visible: false,
          elementsByModel: { 'a.ifc': [1, 2, 'x'], 'b.ifc': 'nope' },
        },
        { id: 42, name: 'bad' }, // dropped: non-string id
        'not-a-record', // dropped
        { id: 's2', name: 'defaults' }, // kept, defaults filled in
      ],
    });
    expect(state?.searchSets).toHaveLength(2);
    expect(state?.searchSets?.[0]).toMatchObject({
      id: 's1',
      colorActive: true,
      visible: false,
      elementsByModel: { 'a.ifc': [1, 2] },
    });
    expect(state?.searchSets?.[1]).toMatchObject({
      id: 's2',
      query: '',
      colorActive: false,
      visible: true,
      elementsByModel: {},
    });
    // Absent (v1 / pre-feature states) → empty list, never crash (A7).
    expect(normalizePersistedState({ version: 2 })?.searchSets).toEqual([]);
  });
});

describe('buildPersistedState — pure serializer (W3.5 extraction, W5.2)', () => {
  const baseInput = (): PersistedStateInput => ({
    selectionMode: 'multi',
    navigationMode: 'Plan',
    visualStyle: 'color-pen',
    xray: true,
    edges: false,
    gridVisible: true,
    backgroundColor: '#0b1220',
    backgroundByTheme: { dark: '#0b1220', light: '#c6d5e8' },
    theme: 'dark',
    language: 'de',
    viewpoints: [],
    issues: [],
    models: [
      {
        modelId: 'a.ifc',
        fileName: 'A.ifc',
        sizeBytes: 2048,
        fragKey: 'k1-800',
        offsetPosition: { x: 5, y: 0, z: 0 },
        offsetRotation: { x: 0, y: 0, z: 0 },
        opacity: 0.7,
        visible: true,
        hiddenIds: [3, 4],
      },
    ],
    camera: { position: { x: 10, y: 10, z: 10 }, target: { x: 0, y: 0, z: 0 }, projection: 'Perspective' },
    sectionPlanes: [{ normal: { x: 1, y: 0, z: 0 }, origin: { x: 2, y: 0, z: 0 } }],
    selection: { 'a.ifc': [7, 8] },
    activeTab: 'models',
    searchSets: [
      {
        id: 'set-1',
        name: 'wall',
        query: 'wall',
        color: '#e4572e',
        colorActive: true,
        visible: false,
        elementsByModel: { 'a.ifc': [7, 8, 9] },
      },
    ],
    maxSnapshotChars: 150_000,
  });

  it('projects the input into a valid v2 state that survives normalize()', () => {
    const built = buildPersistedState(baseInput());
    expect(built.version).toBe(2);
    // A build → serialize → parse → normalize round-trip is lossless in fields.
    const roundTripped = normalizePersistedState(JSON.parse(JSON.stringify(built)));
    expect(roundTripped).toEqual(built);
  });

  it('deep-copies models/selection so the live objects are not aliased', () => {
    const input = baseInput();
    const built = buildPersistedState(input);
    input.models[0].offsetPosition.x = 999;
    input.selection['a.ifc'].push(99);
    expect(built.models[0].offsetPosition.x).toBe(5);
    expect(built.selection['a.ifc']).toEqual([7, 8]);
  });

  it('drops viewpoint snapshots above the size cap (F2 guard)', () => {
    const input = baseInput();
    const big = 'data:image/jpeg;base64,' + 'A'.repeat(200_000);
    input.viewpoints = [
      {
        id: 'v1', name: 'big', createdAt: '2026-07-07T00:00:00.000Z',
        camera: { position: { x: 0, y: 0, z: 0 }, target: { x: 0, y: 0, z: 0 }, projection: 'Perspective', mode: 'Orbit' },
        clippingPlanes: [], hiddenItems: {}, xray: false, edges: false, snapshot: big,
      },
      {
        id: 'v2', name: 'small', createdAt: '2026-07-07T00:00:00.000Z',
        camera: { position: { x: 0, y: 0, z: 0 }, target: { x: 0, y: 0, z: 0 }, projection: 'Perspective', mode: 'Orbit' },
        clippingPlanes: [], hiddenItems: {}, xray: false, edges: false, snapshot: 'data:image/jpeg;base64,AAAA',
      },
    ];
    const built = buildPersistedState(input);
    expect(built.viewpoints[0].snapshot).toBeUndefined();
    expect(built.viewpoints[1].snapshot).toBe('data:image/jpeg;base64,AAAA');
  });
});
