/**
 * Inline SVG icon set (AUDIT U5, C1/A2).
 *
 * The design prototype pulled Material Symbols from Google Fonts and rendered
 * glyphs via ligature text. That has two problems: it is a runtime CDN
 * dependency (C1/A2), and screen readers announce the ligature text
 * ("upload_file") as the accessible name (U5). Inline SVGs fix both: no font
 * fetch, and every icon is `aria-hidden` with a real `aria-label` on its button.
 *
 * Paths are 24x24 Material-Symbols-equivalent outlines. `icon(name)` returns an
 * inline `<svg aria-hidden="true">` string for use in trusted template markup
 * (icon names are compile-time constants — no user data ever reaches these
 * strings, so this is not an A1 injection surface); `setIcon(el, name)` swaps an
 * existing element's icon by building the node via the DOM.
 */

export type IconName =
  | 'upload_file'
  | 'photo_camera'
  | 'download'
  | 'upload'
  | 'light_mode'
  | 'dark_mode'
  | 'view_in_ar'
  | 'near_me'
  | 'select_all'
  | 'center_focus_strong'
  | 'visibility'
  | 'visibility_off'
  | 'straighten'
  | 'square_foot'
  | 'delete_sweep'
  | 'vertical_split'
  | 'horizontal_split'
  | 'layers'
  | 'crop_free'
  | 'blur_on'
  | 'border_style'
  | 'grid_on'
  | 'fit_screen'
  | '3d_rotation'
  | 'crop_portrait'
  | 'crop_landscape'
  | 'map'
  | 'directions_walk'
  | 'account_tree'
  | 'deployed_code'
  | 'info'
  | 'flag'
  | 'help'
  | 'search'
  | 'chevron_right'
  | 'close'
  | 'right_panel_close'
  | 'right_panel_open'
  | 'touch_app'
  | 'delete'
  | 'place'
  | 'content_cut'
  | 'more_horiz'
  | 'home'
  | 'error_outline'
  | 'restart_alt'
  | 'save'
  | 'history'
  | 'filter_list'
  | 'add_comment'
  | 'menu'
  | 'play_circle'
  | 'fullscreen'
  | 'open_in_new';

/** Raw inner SVG (path/shape) markup per icon, on a 0 0 24 24 viewBox. */
const PATHS: Record<IconName, string> = {
  upload_file:
    '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm4 18H6V4h7v5h5v11zM8 15.01l1.41 1.41L11 14.84V19h2v-4.16l1.59 1.59L16 15.01 12.01 11 8 15.01z"/>',
  photo_camera:
    '<circle cx="12" cy="12" r="3.2"/><path d="M9 2 7.17 4H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2h-3.17L15 2H9zm3 15a5 5 0 1 1 0-10 5 5 0 0 1 0 10z"/>',
  download: '<path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>',
  upload: '<path d="M5 20h14v-2H5v2zM5 10h4v6h6v-6h4l-7-7-7 7z"/>',
  light_mode:
    '<circle cx="12" cy="12" r="4.5"/><path d="M12 1v3M12 20v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M1 12h3M20 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1" stroke="currentColor" stroke-width="2" stroke-linecap="round" fill="none"/>',
  dark_mode:
    '<path d="M12 3a9 9 0 1 0 9 9c0-.46-.04-.92-.1-1.36a5.39 5.39 0 0 1-4.4 2.26 5.4 5.4 0 0 1-2.4-10.24C13.64 3.09 12.83 3 12 3z"/>',
  view_in_ar:
    '<path d="M3 4c0-.55.45-1 1-1h3v2H5v2H3V4zm14-1h3c.55 0 1 .45 1 1v3h-2V5h-2V3zM3 17v3c0 .55.45 1 1 1h3v-2H5v-2H3zm18 0v3c0 .55-.45 1-1 1h-3v-2h2v-2h2zM12 7.5 7 10v4l5 2.5 5-2.5v-4L12 7.5zm0 2.2 2.8 1.4L12 12.5 9.2 11.1 12 9.7z"/>',
  // ── Tool-rail set (user directive 2026-07-12): standard BIM-viewer glyphs in
  // a consistent 24px line style (Autodesk-Viewer-like) — stroke 1.6–1.8, round
  // caps/joins. Filled shapes only where the convention is filled (cursors).
  near_me: '<path d="M7 3v14.6l3.7-3.5 2.2 5.3 2.3-1-2.2-5.2 5.1-.5z"/>',
  select_all:
    '<path d="M5.5 5v11.4l2.9-2.7 1.7 4.1 2.1-.9-1.7-4 4-.4z"/><path d="M17.5 4.5v6M14.5 7.5h6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" fill="none"/>',
  center_focus_strong:
    '<path d="M4 8V5.5A1.5 1.5 0 0 1 5.5 4H8M16 4h2.5A1.5 1.5 0 0 1 20 5.5V8M20 16v2.5a1.5 1.5 0 0 1-1.5 1.5H16M8 20H5.5A1.5 1.5 0 0 1 4 18.5V16" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" fill="none"/><path d="M12 8.2 15.3 10.1v3.8L12 15.8 8.7 13.9v-3.8z"/>',
  visibility:
    '<path d="M2.8 12c2-4.2 5.5-6.4 9.2-6.4s7.2 2.2 9.2 6.4c-2 4.2-5.5 6.4-9.2 6.4S4.8 16.2 2.8 12z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><circle cx="12" cy="12" r="2.7" fill="none" stroke="currentColor" stroke-width="1.7"/>',
  visibility_off:
    '<path d="M2.8 12c2-4.2 5.5-6.4 9.2-6.4s7.2 2.2 9.2 6.4c-2 4.2-5.5 6.4-9.2 6.4S4.8 16.2 2.8 12z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><circle cx="12" cy="12" r="2.7" fill="none" stroke="currentColor" stroke-width="1.7"/><path d="M4.5 3.5l15 17" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" fill="none"/>',
  straighten:
    '<path d="M2.8 16.4 16.4 2.8l4.8 4.8L7.6 21.2z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M6.2 13l1.6 1.6M9.6 9.6l1.6 1.6M13 6.2l1.6 1.6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" fill="none"/>',
  square_foot:
    '<path d="M5 7 18 4.6l1.4 12L6.4 19.4z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><circle cx="5" cy="7" r="1.6"/><circle cx="18" cy="4.6" r="1.6"/><circle cx="19.4" cy="16.6" r="1.6"/><circle cx="6.4" cy="19.4" r="1.6"/>',
  delete_sweep:
    '<path d="M13.6 3.9a1.7 1.7 0 0 1 2.4 0l4.1 4.1a1.7 1.7 0 0 1 0 2.4L10 20.5H5.9l-2-2a1.7 1.7 0 0 1 0-2.4z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M9.3 8.2l6.5 6.5M12.5 20.5h8" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" fill="none"/>',
  vertical_split:
    '<path d="M9.3 12.5 14.6 5.5H21l-5.3 7z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.6 14.8l4.6 5.6M9.2 14.8l-4.6 5.6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" fill="none"/>',
  horizontal_split:
    '<path d="M9.3 12.5 14.6 5.5H21l-5.3 7z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.6 14.8l2.3 3m2.3-3-2.3 3m0 0v2.8" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" fill="none"/>',
  layers:
    '<path d="M9.3 12.5 14.6 5.5H21l-5.3 7z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.8 14.8h4.4l-4.4 5.6h4.4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" fill="none"/>',
  crop_free:
    '<path d="M3.2 7.4V4.7c0-.8.7-1.5 1.5-1.5h2.7M16.6 3.2h2.7c.8 0 1.5.7 1.5 1.5v2.7M20.8 16.6v2.7c0 .8-.7 1.5-1.5 1.5h-2.7M7.4 20.8H4.7c-.8 0-1.5-.7-1.5-1.5v-2.7" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" fill="none"/><path d="M12 6.9 16.4 9.4v5.2L12 17.1 7.6 14.6V9.4z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M7.6 9.4 12 11.9l4.4-2.5M12 11.9v5.2" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/>',
  blur_on:
    '<path d="M12 3.6 19.3 7.8v8.4L12 20.4 4.7 16.2V7.8z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M4.7 7.8 12 12l7.3-4.2M12 12v8.4" fill="none" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2.2 2"/>',
  border_style:
    '<path d="M12 3.6 19.3 7.8v8.4L12 20.4 4.7 16.2V7.8z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M4.7 7.8 12 12l7.3-4.2M12 12v8.4" fill="none" stroke="currentColor" stroke-width="1.7"/><circle cx="12" cy="3.6" r="1.4"/><circle cx="19.3" cy="7.8" r="1.4"/><circle cx="19.3" cy="16.2" r="1.4"/><circle cx="12" cy="20.4" r="1.4"/><circle cx="4.7" cy="16.2" r="1.4"/><circle cx="4.7" cy="7.8" r="1.4"/><circle cx="12" cy="12" r="1.4"/>',
  grid_on:
    '<path d="M4.5 4.5h15v15h-15z M9.5 4.5v15M14.5 4.5v15M4.5 9.5h15M4.5 14.5h15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/>',
  fit_screen:
    '<path d="M17 4h3a2 2 0 0 1 2 2v3h-2V6h-3V4zM4 9V6h3V4H4a2 2 0 0 0-2 2v3h2zm16 6v3h-3v2h3a2 2 0 0 0 2-2v-3h-2zM7 18H4v-3H2v3a2 2 0 0 0 2 2h3v-2zM6 8h12v8H6V8z"/>',
  '3d_rotation':
    '<path d="M7.52 21.48A11 11 0 0 1 2.05 13H.03A12 12 0 0 0 12 24l.68-.03-3.44-3.44-1.72.95zM12 0l-.68.03 3.44 3.44 1.72-.94A11 11 0 0 1 21.95 11h2.02A12 12 0 0 0 12 0zm0 6a4 4 0 0 0-2 7.46V16h4v-2.54A4 4 0 0 0 12 6z"/>',
  crop_portrait:
    '<path d="M4.5 9 9.9 3.9H19.5V13.4L14 18.5H4.5z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.5 9H14m0 0v9.5M14 9l5.5-5.1" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.5 9H14v9.5H4.5z" opacity="0.35"/>',
  crop_landscape:
    '<path d="M4.5 9 9.9 3.9H19.5V13.4L14 18.5H4.5z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.5 9H14m0 0v9.5M14 9l5.5-5.1" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/><path d="M4.5 9 9.9 3.9H19.5L14 9z" opacity="0.35"/>',
  map: '<path d="M20.5 3l-.16.03L15 5.1 9 3 3.36 4.9c-.21.07-.36.25-.36.48V20.5c0 .28.22.5.5.5l.16-.03L9 18.9l6 2.1 5.64-1.9c.21-.07.36-.25.36-.48V3.5c0-.28-.22-.5-.5-.5zM15 19l-6-2.11V5l6 2.11V19z"/>',
  directions_walk:
    '<path d="M13.5 5.5a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM9.8 8.9L7 23h2.1l1.8-8 2.1 2v6h2v-7.5l-2.1-2 .6-3A7.3 7.3 0 0 0 18 13v-2c-1.7 0-3.2-.9-4-2.2l-1-1.6c-.4-.6-1-1-1.7-1-.3 0-.5.1-.8.1L6 8.3V13h2V9.6l1.8-.7z"/>',
  account_tree:
    '<path d="M22 11V3h-7v3H9V3H2v8h7V8h2v10h4v3h7v-8h-7v3h-2V8h2v3h7zM7 9H4V5h3v4zm10 6h3v4h-3v-4zm0-10h3v4h-3V5z"/>',
  deployed_code:
    '<path d="M20 8.5 12 4 4 8.5v7L12 20l8-4.5v-7zM12 6.28 17.5 9.5 12 12.72 6.5 9.5 12 6.28zM6 11.24l5 2.87v5.05l-5-2.81v-5.11zm7 7.92v-5.05l5-2.87v5.11l-5 2.81z"/>',
  info: '<path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/>',
  flag: '<path d="M14.4 6 14 4H5v17h2v-7h5.6l.4 2h7V6z"/>',
  help: '<path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm1 17h-2v-2h2v2zm2.07-7.75-.9.92C13.45 12.9 13 13.5 13 15h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26A2 2 0 1 0 10 9H8a4 4 0 1 1 7.07 2.25z"/>',
  search:
    '<path d="M15.5 14h-.79l-.28-.27a6.5 6.5 0 1 0-.7.7l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0A4.5 4.5 0 1 1 9.5 5a4.5 4.5 0 0 1 0 9z"/>',
  chevron_right: '<path d="M10 6 8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/>',
  close: '<path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>',
  right_panel_close:
    '<path d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zm0 16h-5V5h5v14zM8.5 8.5 10 10l-2 2 2 2-1.5 1.5L5 12l3.5-3.5z"/>',
  right_panel_open:
    '<path d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zm0 16h-5V5h5v14zm-9-1.5L6.5 14 9 12 6.5 10 8 8.5 12 12l-2 2 1.5 1.5z"/>',
  touch_app:
    '<path d="M9 11.24V7.5a2.5 2.5 0 0 1 5 0v3.74a5 5 0 1 0-5 0zM17 12h-1v-1.5a1.5 1.5 0 0 0-3 0V6.5a1.5 1.5 0 0 0-3 0V15l-2.4-.5a1 1 0 0 0-.98 1.66l3.2 4.13c.38.49.96.71 1.55.71H16a2 2 0 0 0 2-2v-5a1.5 1.5 0 0 0-1-1.5z"/>',
  delete:
    '<path d="M6 19a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/>',
  place:
    '<path d="M12 2a7 7 0 0 0-7 7c0 5.25 7 13 7 13s7-7.75 7-13a7 7 0 0 0-7-7zm0 9.5a2.5 2.5 0 1 1 0-5 2.5 2.5 0 0 1 0 5z"/>',
  content_cut:
    '<path d="M9.64 7.64c.23-.5.36-1.05.36-1.64a4 4 0 1 0-4 4c.59 0 1.14-.13 1.64-.36L10 12l-2.36 2.36A3.9 3.9 0 0 0 6 14a4 4 0 1 0 4 4c0-.59-.13-1.14-.36-1.64L12 14l7 7h3v-1L9.64 7.64zM6 8a2 2 0 1 1 0-4 2 2 0 0 1 0 4zm0 12a2 2 0 1 1 0-4 2 2 0 0 1 0 4zm6-7.5a.5.5 0 1 1 0-1 .5.5 0 0 1 0 1zM19 3l-6 6 2 2 7-7V3z"/>',
  more_horiz:
    '<path d="M6 10a2 2 0 1 0 0 4 2 2 0 0 0 0-4zm12 0a2 2 0 1 0 0 4 2 2 0 0 0 0-4zm-6 0a2 2 0 1 0 0 4 2 2 0 0 0 0-4z"/>',
  home: '<path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/>',
  error_outline:
    '<path d="M11 15h2v2h-2v-2zm0-8h2v6h-2V7zm1-5a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 18a8 8 0 1 1 0-16 8 8 0 0 1 0 16z"/>',
  restart_alt:
    '<path d="M12 5V2L8 6l4 4V7a5 5 0 1 1-5 5H5a7 7 0 1 0 7-7z"/>',
  save:
    '<path d="M17 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V7l-4-4zm-5 16a3 3 0 1 1 0-6 3 3 0 0 1 0 6zm3-10H5V5h10v4z"/>',
  history:
    '<path d="M13 3a9 9 0 0 0-9 9H1l3.9 3.9.1.1L9 12H6a7 7 0 1 1 7 7 6.97 6.97 0 0 1-4.95-2.05l-1.42 1.42A9 9 0 1 0 13 3zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/>',
  filter_list: '<path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/>',
  add_comment:
    '<path d="M20 2H4a2 2 0 0 0-2 2v18l4-4h14a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2zm-3 9h-4v4h-2v-4H7V9h4V5h2v4h4v2z"/>',
  menu: '<path d="M3 18h18v-2H3v2zm0-5h18v-2H3v2zm0-7v2h18V6H3z"/>',
  play_circle:
    '<path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm-2 14.5v-9l7 4.5-7 4.5z"/>',
  fullscreen:
    '<path d="M7 14H5v5h5v-2H7v-3zm-2-4h2V7h3V5H5v5zm12 7h-3v2h5v-5h-2v3zM14 5v2h3v3h2V5h-5z"/>',
  open_in_new:
    '<path d="M19 19H5V5h7V3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7h-2v7zM14 3v2h3.59l-9.83 9.83 1.41 1.41L19 6.41V10h2V3h-7z"/>',
};

/** Returns an inline `<svg>` string for the icon, decorative (aria-hidden). */
export function icon(name: IconName, sizePx = 20): string {
  const inner = PATHS[name];
  if (inner === undefined) return '';
  return `<svg class="btc-icon" aria-hidden="true" focusable="false" width="${sizePx}" height="${sizePx}" viewBox="0 0 24 24" fill="currentColor">${inner}</svg>`;
}

/**
 * Replaces the icon inside `el` (used for stateful toggles like theme/eye).
 *
 * Uses innerHTML with the trusted static icon string: the HTML parser places
 * the <svg> in the correct SVG namespace so its width/height/fill presentation
 * attributes render. (The prior DOMParser('image/svg+xml') + node-adoption path
 * produced a mis-namespaced/adopted <svg> whose attributes were ignored — every
 * icon rendered at 0×0 with black fill; see AUDIT A18.) `icon()` interpolates
 * only static template constants (no user data), so this is not an A1 surface.
 */
export function setIcon(el: Element, name: IconName, sizePx = 20): void {
  el.innerHTML = icon(name, sizePx);
}

/**
 * Hydrates every `[data-icon]` element in `root` with its inline SVG. Called
 * once at boot so index.html stays icon-declarative (the icon set is the single
 * source of truth). `data-icon` values are static template constants — not a
 * user-data path, so this is not an A1 injection surface.
 */
export function hydrateIcons(root: ParentNode): void {
  root.querySelectorAll<HTMLElement>('[data-icon]').forEach((el) => {
    const name = el.dataset.icon as IconName | undefined;
    if (!name || PATHS[name] === undefined) return;
    const size = el.dataset.iconSize ? Number(el.dataset.iconSize) : 20;
    setIcon(el, name, Number.isFinite(size) ? size : 20);
  });
}
