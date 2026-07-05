/**
 * Static shell behaviors (splitters, menubar dropdowns, tab scroll, mobile
 * FAB). Moved out of an inline <script> in index.html so the CSP can enforce
 * `script-src 'self'` without 'unsafe-inline' (AUDIT A1). Replaced wholesale
 * by the W3 rebuild.
 */

// Splitter drag behavior
document.querySelectorAll<HTMLElement>('.panel-splitter').forEach((splitter) => {
  let dragging = false;
  let startX = 0;
  let startW = 0;
  const side = splitter.dataset.splitter;
  const prop = side === 'left' ? '--model-panel-w' : '--panel-w';
  const min = side === 'left' ? 180 : 200;
  const max = side === 'left' ? 400 : 420;

  splitter.addEventListener('mousedown', (event) => {
    dragging = true;
    startX = event.clientX;
    startW = parseInt(getComputedStyle(document.documentElement).getPropertyValue(prop), 10);
    splitter.classList.add('dragging');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    event.preventDefault();
  });

  document.addEventListener('mousemove', (event) => {
    if (!dragging) return;
    const delta = side === 'left' ? event.clientX - startX : startX - event.clientX;
    const width = Math.min(max, Math.max(min, startW + delta));
    document.documentElement.style.setProperty(prop, `${width}px`);
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    splitter.classList.remove('dragging');
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  });

  let savedW = 0;
  splitter.addEventListener('dblclick', () => {
    const current = parseInt(getComputedStyle(document.documentElement).getPropertyValue(prop), 10);
    if (current > 0) {
      savedW = current;
      document.documentElement.style.setProperty(prop, '0px');
    } else {
      document.documentElement.style.setProperty(prop, `${savedW || (side === 'left' ? 260 : 280)}px`);
    }
  });
});

// Mobile FAB triggers file upload
const fab = document.getElementById('mobileFab');
if (fab) {
  fab.addEventListener('click', () => document.getElementById('btnUpload')?.click());
}

// Menu dropdown functionality
document.querySelectorAll<HTMLButtonElement>('.menu-dropdown > .menu-item').forEach((button) => {
  button.addEventListener('click', (event) => {
    const dropdown = button.parentElement;
    if (!dropdown) return;
    const wasOpen = dropdown.classList.contains('open');
    document.querySelectorAll('.menu-dropdown.open').forEach((d) => d.classList.remove('open'));
    if (!wasOpen) dropdown.classList.add('open');
    event.stopPropagation();
  });
});

document.querySelectorAll<HTMLButtonElement>('.menu-dropdown-content button[data-menu-action]').forEach((button) => {
  button.addEventListener('click', () => {
    const action = button.dataset.menuAction;
    if (action === 'tab-help') {
      document.querySelector<HTMLButtonElement>('[data-tab="help"]')?.click();
    } else if (action) {
      document.getElementById(action)?.click();
    }
    document.querySelectorAll('.menu-dropdown.open').forEach((d) => d.classList.remove('open'));
  });
});

// Prevent menu close when interacting with inline form controls
document.querySelectorAll<HTMLElement>('.menu-inline-control').forEach((control) => {
  control.addEventListener('click', (event) => event.stopPropagation());
  control.addEventListener('mousedown', (event) => event.stopPropagation());
});

document.addEventListener('click', () => {
  document.querySelectorAll('.menu-dropdown.open').forEach((d) => d.classList.remove('open'));
});

// Grid off by default — dispatch change after viewer init.
// (AUDIT F4: this 500 ms hack races init and is deleted in W1.5 — kept
// verbatim in this commit so the CSP move stays behavior-neutral.)
setTimeout(() => {
  const grid = document.getElementById('toggleGrid') as HTMLInputElement | null;
  if (grid && grid.checked) {
    grid.checked = false;
    grid.dispatchEvent(new Event('change'));
  }
}, 500);

// Tab scroll buttons
const tabsWrap = document.querySelector<HTMLElement>('.panel-tabs-wrap');
const tabsContainer = document.querySelector<HTMLElement>('.panel-tabs');
if (tabsWrap && tabsContainer) {
  const checkOverflow = (): void => {
    tabsWrap.classList.toggle('can-scroll', tabsContainer.scrollWidth > tabsContainer.clientWidth);
  };
  checkOverflow();
  new ResizeObserver(checkOverflow).observe(tabsContainer);
  tabsWrap.querySelectorAll<HTMLButtonElement>('.tab-scroll-btn').forEach((button) => {
    button.addEventListener('click', () => {
      const direction = button.dataset.scrollDir === 'left' ? -80 : 80;
      tabsContainer.scrollBy({ left: direction, behavior: 'smooth' });
    });
  });
}
