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
  | 'filter_list'
  | 'add_comment'
  | 'menu';

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
  near_me: '<path d="M21 3 3 10.53v.98l6.84 2.65L12.48 21h.98L21 3z"/>',
  select_all:
    '<path d="M3 5h2V3c-1.1 0-2 .9-2 2zm0 8h2v-2H3v2zm4 8h2v-2H7v2zM3 9h2V7H3v2zm10-6h-2v2h2V3zm6 0v2h2c0-1.1-.9-2-2-2zM5 21v-2H3c0 1.1.9 2 2 2zm-2-4h2v-2H3v2zM9 3H7v2h2V3zm2 18h2v-2h-2v2zm8-8h2v-2h-2v2zm0 8c1.1 0 2-.9 2-2h-2v2zm0-12h2V7h-2v2zm0 8h2v-2h-2v2zm-4 4h2v-2h-2v2zm0-16h2V3h-2v2zM7 17h10V7H7v10zm2-8h6v6H9V9z"/>',
  center_focus_strong:
    '<path d="M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8zm0 6a2 2 0 1 1 0-4 2 2 0 0 1 0 4zM3 5v4h2V5h4V3H5a2 2 0 0 0-2 2zm2 10H3v4a2 2 0 0 0 2 2h4v-2H5v-4zm14 4h-4v2h4a2 2 0 0 0 2-2v-4h-2v4zm0-16h-4v2h4v4h2V5a2 2 0 0 0-2-2z"/>',
  visibility:
    '<path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5C21.27 7.61 17 4.5 12 4.5zm0 12a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9zm0-7a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z"/>',
  visibility_off:
    '<path d="M12 6.5a5.5 5.5 0 0 1 5.5 5.5c0 .73-.16 1.42-.42 2.05l3.21 3.21A11.8 11.8 0 0 0 23 12c-1.73-4.39-6-7.5-11-7.5-1.2 0-2.36.18-3.45.5l2.32 2.32c.36-.2.75-.32 1.13-.32zM2.7 3.3 1.3 4.7l2.5 2.5A11.8 11.8 0 0 0 1 12c1.73 4.39 6 7.5 11 7.5 1.55 0 3.03-.3 4.38-.84l3.2 3.2 1.4-1.42L2.7 3.3zM7.5 12a4.5 4.5 0 0 0 6.28 4.14l-1.5-1.5A2.5 2.5 0 0 1 9.36 11.2l-1.5-1.5c-.23.7-.36 1.45-.36 2.3z"/>',
  straighten:
    '<path d="M21 6H3a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h18a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2zm0 10H3V8h2v4h2V8h2v4h2V8h2v4h2V8h2v4h2V8h2v8z"/>',
  square_foot:
    '<path d="M17.66 17.66l-1.06 1.06-.71-.71 1.06-1.06-1.94-1.94-1.06 1.06-.71-.71 1.06-1.06-1.94-1.94-1.06 1.06-.71-.71 1.06-1.06L9 8.71V21h12v-3l-3.34-.34zM19 19h-8v-7.59L19 19zM6.99 2.99 2 8v11h2V9l4.99-4.99-2-1.02z"/>',
  delete_sweep:
    '<path d="M15 16h4v2h-4v-2zm0-8h7v2h-7V8zm0 4h6v2h-6v-2zM3 18c0 1.1.9 2 2 2h6c1.1 0 2-.9 2-2V8H3v10zM14 5h-3l-1-1H6L5 5H2v2h12V5z"/>',
  vertical_split:
    '<path d="M3 5v14h18V5H3zm8 12H5V7h6v10zm8 0h-6v-2h6v2zm0-4h-6v-2h6v2zm0-4h-6V7h6v2z"/>',
  horizontal_split:
    '<path d="M3 5v14h18V5H3zm16 6H5V7h14v4zm0 6h-6v-2h6v2zm-8 0H5v-2h6v2z"/>',
  layers:
    '<path d="M11.99 18.54 4.62 12.8 3 14.07l9 7 9-7-1.63-1.27-7.38 5.74zM12 16l7.36-5.73L21 9l-9-7-9 7 1.63 1.27L12 16z"/>',
  crop_free:
    '<path d="M3 5v4h2V5h4V3H5a2 2 0 0 0-2 2zm2 10H3v4a2 2 0 0 0 2 2h4v-2H5v-4zm14 4h-4v2h4a2 2 0 0 0 2-2v-4h-2v4zM19 3h-4v2h4v4h2V5a2 2 0 0 0-2-2z"/>',
  blur_on:
    '<path d="M6 13a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm0-4a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm0 8a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm-4-4a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm8 8a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm-4 0a1 1 0 1 0 0 2 1 1 0 0 0 0-2zM10 5a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm10 8a1 1 0 1 0 0 2 1 1 0 0 0 0-2zM10 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6zm0 4a1 1 0 1 1 0-2 1 1 0 0 1 0 2zm0-8a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm4 16a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm0-16a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm4 12a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm0-4a1 1 0 1 0 0 2 1 1 0 0 0 0-2zm0-4a1 1 0 1 0 0 2 1 1 0 0 0 0-2z"/>',
  border_style:
    '<path d="M15 21h2v-2h-2v2zm4 0h2v-2h-2v2zM7 21h2v-2H7v2zm4 0h2v-2h-2v2zm8-4h2v-2h-2v2zm0-4h2v-2h-2v2zM3 3v18h2V5h16V3H3zm16 6h2V7h-2v2z"/>',
  grid_on:
    '<path d="M20 2H4c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zM8 20H4v-4h4v4zm0-6H4v-4h4v4zm0-6H4V4h4v4zm6 12h-4v-4h4v4zm0-6h-4v-4h4v4zm0-6h-4V4h4v4zm6 12h-4v-4h4v4zm0-6h-4v-4h4v4zm0-6h-4V4h4v4z"/>',
  fit_screen:
    '<path d="M17 4h3a2 2 0 0 1 2 2v3h-2V6h-3V4zM4 9V6h3V4H4a2 2 0 0 0-2 2v3h2zm16 6v3h-3v2h3a2 2 0 0 0 2-2v-3h-2zM7 18H4v-3H2v3a2 2 0 0 0 2 2h3v-2zM6 8h12v8H6V8z"/>',
  '3d_rotation':
    '<path d="M7.52 21.48A11 11 0 0 1 2.05 13H.03A12 12 0 0 0 12 24l.68-.03-3.44-3.44-1.72.95zM12 0l-.68.03 3.44 3.44 1.72-.94A11 11 0 0 1 21.95 11h2.02A12 12 0 0 0 12 0zm0 6a4 4 0 0 0-2 7.46V16h4v-2.54A4 4 0 0 0 12 6z"/>',
  crop_portrait:
    '<path d="M17 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zm0 16H7V5h10v14z"/>',
  crop_landscape:
    '<path d="M19 5H5a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2zm0 12H5V7h14v10z"/>',
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
  filter_list: '<path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/>',
  add_comment:
    '<path d="M20 2H4a2 2 0 0 0-2 2v18l4-4h14a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2zm-3 9h-4v4h-2v-4H7V9h4V5h2v4h4v2z"/>',
  menu: '<path d="M3 18h18v-2H3v2zm0-5h18v-2H3v2zm0-7v2h18V6H3z"/>',
};

/** Returns an inline `<svg>` string for the icon, decorative (aria-hidden). */
export function icon(name: IconName, sizePx = 20): string {
  const inner = PATHS[name];
  if (inner === undefined) return '';
  return `<svg class="btc-icon" aria-hidden="true" focusable="false" width="${sizePx}" height="${sizePx}" viewBox="0 0 24 24" fill="currentColor">${inner}</svg>`;
}

/**
 * Replaces the icon inside `el` (used for stateful toggles like theme/eye).
 * Parses the trusted static icon string once and swaps the node — avoids
 * assigning to innerHTML directly.
 */
const iconParser = new DOMParser();
export function setIcon(el: Element, name: IconName, sizePx = 20): void {
  const doc = iconParser.parseFromString(icon(name, sizePx), 'image/svg+xml');
  const svg = doc.documentElement;
  el.replaceChildren(svg);
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
