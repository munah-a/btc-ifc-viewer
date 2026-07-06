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

/** Already-translated labels for the sheet (C7); passed in by the orchestrator. */
export interface MobileSheetLabels {
  xray: string;
  edges: string;
  grid: string;
  lightTheme: string;
  style: string;
  /** Visual-style option label for a given style value. */
  styleOption(value: string): string;
}

const STYLE_VALUES: readonly string[] = ['basic', 'pen', 'color-pen', 'color-shadows', 'color-pen-shadows'];

export function buildMobileSheet(
  sheetBody: HTMLElement,
  state: MobileSheetState,
  actions: MobileSheetActions,
  labels: MobileSheetLabels,
): void {
  const toggles: Array<{ icon: IconName; label: string; on: boolean; onClick: () => void }> = [
    { icon: 'blur_on', label: labels.xray, on: state.xrayEnabled, onClick: actions.toggleXray },
    { icon: 'border_style', label: labels.edges, on: state.edgesEnabled, onClick: actions.toggleEdges },
    { icon: 'grid_on', label: labels.grid, on: state.gridVisible, onClick: actions.toggleGrid },
    { icon: state.themeMode === 'dark' ? 'dark_mode' : 'light_mode', label: labels.lightTheme, on: state.themeMode === 'light', onClick: actions.toggleTheme },
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
  styleLabel.textContent = labels.style;
  styleLabel.style.flex = '1';
  const select = document.createElement('select');
  select.className = 'text-input';
  select.style.width = 'auto';
  for (const value of STYLE_VALUES) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = labels.styleOption(value);
    if (value === state.visualStyle) option.selected = true;
    select.append(option);
  }
  select.addEventListener('change', () => actions.setVisualStyle(select.value));
  field.append(styleLabel, select);
  sheetBody.append(field);
}
