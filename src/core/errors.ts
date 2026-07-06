/**
 * Error helpers (AUDIT A9). The fragments engine throws plain Errors when an
 * operation races a model unload/dispose ("Model not found") — a benign,
 * self-healing condition. The predicate names that classification instead of
 * scattering message-substring checks, and callers log suppressed instances
 * via console.debug so the condition stays observable.
 *
 * NOTE: message matching is the only classification the library offers today
 * (no typed error classes on the ^3.3 range) — revisit if W6 upgrades expose
 * typed errors.
 */

export const serializeError = (error: unknown): string => {
  if (error instanceof Error) return error.message;
  return String(error);
};

/** True for the engine's benign "model not found" race (A9). */
export const isModelNotFoundError = (error: unknown): boolean =>
  serializeError(error).toLowerCase().includes('model not found');
