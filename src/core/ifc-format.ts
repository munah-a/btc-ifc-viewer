/**
 * Fast, deterministic IFC sniff. Real IFC files are STEP (ISO-10303-21) text
 * that begins with the `ISO-10303-21;` header (optionally after a UTF-8/UTF-16
 * BOM and whitespace). We check this BEFORE handing bytes to web-ifc because
 * web-ifc's behavior on non-IFC input is undefined — it may throw, hang, or
 * silently yield an empty model, and which one happens varies by environment
 * (AUDIT T13: a corrupt-file test passed locally but the error state never
 * appeared on the CI runner). Sniffing here gives the user an instant, clear
 * "not a valid IFC" error on every platform.
 */
export function isProbablyIfc(bytes: Uint8Array): boolean {
  // Header always lands well within the first block; decode a small prefix.
  const prefix = new TextDecoder('utf-8', { fatal: false }).decode(bytes.subarray(0, 512));
  return prefix.includes('ISO-10303-21');
}
