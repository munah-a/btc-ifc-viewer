/**
 * Mobile "More" sheet (U2) — builds the bottom-sheet body with the visual
 * toggles (x-ray/edges/grid/theme) and the visual-style selector. Extracted
 * from viewer.ts: it constructs the DOM into the provided sheet body and calls
 * back into the orchestrator for each action. State + callbacks are passed
 * explicitly so this holds no app state.
 */
import { setIcon, type IconName } from './icons';
import type { VisualStyle } from '../core/persistence';

export interface MobileSheetState {
  xrayEnabled: boolean;
  edgesEnabled: boolean;
  gridVisible: boolean;
  themeMode: 'dark' | 'light';
  visualStyle: VisualStyle;
}

export interface MobileSheetActions {
  toggleXray: () => void;
  toggleEdges: () => void;
  toggleGrid: () => void;
  toggleTheme: () => void;
  setVisualStyle: (value: string) => void;
}

const STYLE_OPTIONS: ReadonlyArray<readonly [string, string]> = [
  ['basic', 'Basic'],
  ['pen', 'Pen'],
  ['color-pen', 'Color pen'],
  ['color-shadows', 'Color shadows'],
  ['color-pen-shadows', 'Color pen shadows'],
];

export function buildMobileSheet(
  sheetBody: HTMLElement,
  state: MobileSheetState,
  actions: MobileSheetActions,
): void {
  const toggles: Array<{ icon: IconName; label: string; on: boolean; onClick: () => void }> = [
    { icon: 'blur_on', label: 'X-ray', on: state.xrayEnabled, onClick: actions.toggleXray },
    { icon: 'border_style', label: 'Edges', on: state.edgesEnabled, onClick: actions.toggleEdges },
    { icon: 'grid_on', label: 'Grid', on: state.gridVisible, onClick: actions.toggleGrid },
    { icon: state.themeMode === 'dark' ? 'dark_mode' : 'light_mode', label: 'Light theme', on: state.themeMode === 'light', onClick: actions.toggleTheme },
  ];
  sheetBody.replaceChildren();
  for (const t of toggles) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `sheet-toggle${t.on ? ' is-active' : ''}`;
    const iconSpan = document.createElement('span');
    setIcon(iconSpan, t.icon);
    const label = document.createElement('span');
    label.textContent = t.label;
    label.style.flex = '1';
    const track = document.createElement('span');
    track.className = 'sheet-toggle-track';
    const knob = document.createElement('span');
    knob.className = 'sheet-toggle-knob';
    track.append(knob);
    button.append(iconSpan, label, track);
    button.addEventListener('click', t.onClick);
    sheetBody.append(button);
  }
  // Visual style selector
  const field = document.createElement('label');
  field.className = 'sheet-toggle';
  field.style.gap = '10px';
  const styleLabel = document.createElement('span');
  styleLabel.textContent = 'Style';
  styleLabel.style.flex = '1';
  const select = document.createElement('select');
  select.className = 'text-input';
  select.style.width = 'auto';
  for (const [value, label] of STYLE_OPTIONS) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    if (value === state.visualStyle) option.selected = true;
    select.append(option);
  }
  select.addEventListener('change', () => actions.setVisualStyle(select.value));
  field.append(styleLabel, select);
  sheetBody.append(field);
}
