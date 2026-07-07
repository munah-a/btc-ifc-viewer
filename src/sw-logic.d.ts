/** Type declarations for the pure SW decision helpers (src/sw-logic.js). */
export function pickShellName(pathname: string): 'embed.html' | 'index.html';
export function shouldCacheNavigation(
  response: { ok?: boolean; type?: string } | null | undefined,
): boolean;
