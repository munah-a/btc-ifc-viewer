/**
 * Engine-adjacent helpers that carry NO heavy engine dependencies (no three,
 * @thatopen or web-ifc). Extracted from viewer-core.ts (W5.1 / P1) so the app
 * shell can import them statically without pulling the ~1MB-gzip engine into the
 * initial download — the engine itself is dynamically imported on demand.
 *
 * DOM-free except for the FPS monitor's single text output element.
 */

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
