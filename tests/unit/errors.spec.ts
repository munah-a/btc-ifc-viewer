import { describe, expect, it } from 'vitest';

import { isModelNotFoundError, serializeError } from '../../src/core/errors';

describe('serializeError', () => {
  it('unwraps Error messages and stringifies everything else', () => {
    expect(serializeError(new Error('boom'))).toBe('boom');
    expect(serializeError('plain')).toBe('plain');
    expect(serializeError(42)).toBe('42');
    expect(serializeError(undefined)).toBe('undefined');
  });
});

describe('isModelNotFoundError (AUDIT A9 — W1.8)', () => {
  it('matches the engine race in any casing, Error or string', () => {
    expect(isModelNotFoundError(new Error('Model not found: school_str.ifc'))).toBe(true);
    expect(isModelNotFoundError(new Error('MODEL NOT FOUND'))).toBe(true);
    expect(isModelNotFoundError('model not found while updating')).toBe(true);
  });

  it('rejects unrelated errors', () => {
    expect(isModelNotFoundError(new Error('WebGL context lost'))).toBe(false);
    expect(isModelNotFoundError(new Error('model found'))).toBe(false);
    expect(isModelNotFoundError(null)).toBe(false);
    expect(isModelNotFoundError(undefined)).toBe(false);
  });
});
