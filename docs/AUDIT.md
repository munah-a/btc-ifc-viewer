# BTC IFC Viewer — Audit Findings Registry

> Source: full multi-agent implementation audit, 2026-07-05 (55 agents, all factual claims adversarially
> verified against the code; one claim refuted by live-running the app). This file is the **single source
> of truth** for defect specs. Implementation tasks in `IMPLEMENTATION_PLAN.md` reference these IDs.
> **Do not re-audit these areas — trust the file:line references and verify only after fixing.**
>
> Line numbers refer to the repo state at commit `498c9c1`. If the file has since changed, locate the
> pattern described rather than the exact line.

Verdicts: ✅ confirmed by adversarial verifier · 🔬 confirmed by live-running the app · ⚠️ unverified
(session limit) but corroborated by ≥2 independent auditors · ❌ refuted.

## A. Architecture (viewer.ts unless noted)

| ID | Sev | Verdict | Location | Issue → Fix |
|----|-----|---------|----------|-------------|
| A1 | High | ✅ | `viewer.ts:1443-1445, 1480-1487, 1500-1507` | XSS: IFC-derived strings (storey/class names, modelId) interpolated into innerHTML unescaped in renderSpatialTree / renderClassFilters / renderLevelFilters. Sibling renderers use `escapeHtml` (line 312) correctly. → Apply `escapeHtml` in these paths; add CSP meta tag. |
| A2 | High | ✅ | `viewer.ts:1153-1159, 1163-1167` | Runtime hard-dependency on unpkg (web-ifc WASM) + thatopen.github.io (fragments worker, unpinned). Local `public/web-ifc.wasm` unused (referenced only by dead code at 2935). → Self-host: WASM from `node_modules/web-ifc/web-ifc.wasm`, worker from `node_modules/@thatopen/fragments/dist/Worker/worker.mjs` (v3.3.5), copied to `public/` at build. |
| A3 | Med | ✅ | `viewer.ts:2932-3296` | 365 lines dead code (`_unused_extractAndApplyMaterialColors`) keeps whole web-ifc JS API in bundle. Also dead: `setRightView/LeftView/BackView/BottomView` (3485-3499), `listen()` (357-363). → Delete all; git history preserves. |
| A4 | High | ✅ | `viewer.ts:365-459, 5275` | Untestable by construction: DOM cache of ~90 required IDs throws at construction; bootstrap at module scope outside try/catch → blank page on any renamed ID; pure logic unexported. → Extract pure modules (see Wave 2); lazy/injected DOM; guarded bootstrap. |
| A5 | Med | ✅ | `viewer.ts:575-616, 545-569` | `destroy()` never called; window/document listeners (915, 918, 1021) escape the abort-signal patch. `console.warn` permanently monkey-patched (641-655); patch clobbers caller `signal` option. → Wire destroy to HMR dispose or delete the machinery; route global listeners through `listen()`. |
| A6 | Med | ✅ | `viewer.ts:1263, 2854, 3298-3378` | Model identity: FIFO `pendingModelMetaQueue` can mis-attribute file names; polling loops (40ms/8s/30s, isBusy 10s); ~130 lines alias-resolution (2326-2414) papering over it. → Key metadata by the name passed to `ifcLoader.load` (becomes `model.modelId`); collapse alias layer. |
| A7 | Med | ✅ | `viewer.ts:4955-4999 vs 5041-5079` | `importViewerState` duplicates `restoreLocalState`; unguarded `parsed.viewpoints` crashes on import without viewpoints array (4969); dead legacy fallback `'xray'/'shaded'` (4963). → Single `applyPersistedState()` with schema validation/defaults. |
| A8 | Med | ✅ | `viewer.ts:1245, 2351, 3649, 3760, 3822` | ThatOpen boundary fully untyped (~20 `any` sites) incl. private `_controls` access (3649) on ^-ranged deps. → Define `FragmentsModelLike` interface for the ~10 methods used; isolate `_controls` behind one function. |
| A9 | Med | ✅ | `viewer.ts:2420-2433` | Errors classified by message substring `'model not found'`; suppressed branch logs nothing. → Named predicate + console.debug; typed errors if library exposes them. |
| A10 | Med | — | `viewer.ts:2876-2882, 2904-2907` | Load timeout doesn't cancel: late model registers anyway (ghost), steals next file's metadata from FIFO. → `.catch` before race; record stale id; ignore/dispose late arrivals. (Same root as A6.) |
| A11 | Med | — | `viewer.ts:1292-1298, 2482+` | Full innerHTML re-render collapses `<details>` state, rebuilds slider mid-drag; mixed wiring (delegation vs per-item listeners). → Standardize on delegation; preserve expanded state. |
| A12 | Low | — | `viewer.ts:3468/3501, 770-796, 3979/4044, 1461` | Duplicated blocks: fitToModel≡setHomeView; 3 section handlers; array-preview ×2; spatial-tree isolate reimplemented. → Collapse each. |
| A13 | Low | — | `viewer.ts:492, 937, 3822 + styles.css` | Brand magenta `#c8145c` in 4 places; magic numbers (chunkSize 360, timeout 120000, xray 0.28…). → Promote to constants block; brand color from CSS custom property. |
| A14 | Low | — | `tsconfig.json:12-17` | Inert emit options under noEmit; missing noUnusedLocals/noUnusedParameters/noUncheckedIndexedAccess. → Clean + enable. |
| A15 | Low | — | `viewer.ts:5259-5272` | FPS counter measures rAF cadence not render work; overwrites load metrics within 1s. → Hook renderer events; separate status slot. |

## F. Features (broken behavior)

| ID | Sev | Verdict | Location | Issue → Fix |
|----|-----|---------|----------|-------------|
| F1 | High | ⚠️ (API shape confirmed in node_modules) | `viewer.ts:2761-2776` | Search crashes on any hit: fragments v3.3 `getItemsData` returns `{value,type}` ItemAttribute objects; `escapeHtml(obj)` throws TypeError; results never render. Every other path unwraps via `readPrimitiveValue`. → Unwrap attributes in `searchElements`. |
| F2 | High | 🔬 | `viewer.ts:4561, 4934` | Screenshots + viewpoint snapshots are **blank transparent PNGs**: `toDataURL()` on non-preserved WebGL buffer with no render-before-capture (verified by running the app: 16KB transparent PNG; control render-then-capture produced real 532KB image). → Render immediately before capture (or preserveDrawingBuffer); then downscale viewpoint thumbnails (~320px JPEG) and show them in the list. Supersedes the refuted localStorage-quota claim (❌ A-quota): quota risk only becomes real once capture is fixed — cap/downscale then. |
| F3 | High | ⚠️ | `viewer.ts:2301-2307, 1300` | Loading any model resets X-ray/edges: `setVisualStyle` clears toggles; `onModelAdded` calls it per registration; persisted `xray:true` can never survive. → Move toggle-reset to explicit user style change only. |
| F4 | Med | ⚠️ | `index.html:644-647` vs `viewer.ts:2104` | 500ms grid-reset hack races init; default grid state nondeterministic; can clobber persisted preference. → Delete hack; one default in viewer.ts. |
| F5 | Med | ⚠️ | `viewer.ts:3908+, 3838+` | Property units fabricated from label keywords, not IfcUnitAssignment. → Read the model's unit assignment; keep keyword inference as fallback only. |
| F6 | Med | ⚠️ | — | No model unload; fragments/indices/worker memory accumulates until reload. → Add per-model dispose in federation panel (fragments.core disposal API). |
| F7 | Low | — | `index.html:101, 194` vs `viewer.ts:748-759` | "Show All" menu runs show-selection, errors with no selection; tooltip disagrees. → Rename or repoint to Reset Visibility. |
| F8 | Low | — | `viewer.ts:849` | Theme toggle force-overwrites custom background. → Per-theme background memory; only swap defaults. |
| F9 | Low | — | `viewer.ts:4693+` | Issues capture only first model's elements in multi-model selection; orphan pins render with no model loaded. → Iterate all models; hide pins with no resolvable model. |
| F10 | Low | — | `viewer.ts:4633-4638` | Viewpoint apply restores clipping without gizmo-visibility fix; section buttons out of sync. → Reuse section-plane creation path. |
| F11 | Low | — | filters | Disjoint class∩level filter silently hides entire model. → Warn/empty-state when intersection is empty. |

## U. UI/UX

| ID | Sev | Verdict | Location | Issue → Fix |
|----|-----|---------|----------|-------------|
| U1 | **Critical** | ✅ | `styles.css:1960-2110`, `index.html:519-538` | **All panels unreachable ≤1023px**: `.drawer-open` never added by any JS; 5 `[data-mobile-nav]` buttons have no handlers; phone hides dock+status bar too (2075, 2112) → canvas + upload FAB only, no error channel. → Wire drawers/backdrop/bottom-nav; move `#viewerDock` out of `.app-toolbar`; route errors through toasts. |
| U2 | High | ✅ | `styles.css:1952, 2074`, `index.html:34-77` | Theme/style/grid/background live only in menubar, hidden ≤1023px — no light theme on tablets (sunlight!). → Relocate to responsive settings surface; default from `prefers-color-scheme`. |
| U3 | High | ✅ | `styles.css:1643, 1683, 2033, 1516, 1373, 1463` | Hardcoded dark rgba breaks light theme (flyouts ~1.5:1 contrast, white-on-white hover, invisible separators). → Replace with theme tokens + light overrides. |
| U4 | High | ✅ | `viewer.ts:2916, 4997, 5037, 2430` | Errors only as 11px status text (hidden on phones); `showToast` used only for successful deletes. → All catch paths → error toast; loading overlay gets error state + Retry. |
| U5 | Med | ✅ | `styles.css:348, 430, 1669, 1724` | Icon-only buttons announce as `upload_file`/`3d_rotation` (ligature text; labels display:none'd; zero aria-hidden in repo). Google Fonts failure shows raw words. → aria-hidden on icons + aria-label per button; prefer inline SVG (rebrand covers this). |
| U6 | Med | ✅ | `viewer.ts:2771+, 1451+, 4681+, 4806+` | Search/tree/viewpoint/issue rows are clickable divs — no keyboard/SR path (model browser does it right with buttons/details). → `<button>` rows + list semantics. |
| U7 | Med | ✅ | `index.html:348, 14, 594-624` | Broken ARIA: tablist without tabs; menus without aria-expanded/Escape; `role=application` on whole shell. → Complete tabs pattern; menu semantics; drop role=application. |
| U8 | Med | ✅ | `viewer.ts:5216-5257, 5117` | Confirm dialog: Escape leaks to global shortcuts (cancels tools behind dialog), no focus trap, focus not restored, destructive button pre-focused. → `<dialog>.showModal()`; focus Cancel; restore invoker. |
| U9 | Med | — | `index.html:552-587` | Splitters mouse-only (no touch/keyboard), 4px hit target. → Pointer Events + setPointerCapture; ~12px hit area; role=separator + arrows. |
| U10 | Med | — | `styles.css:436-484, 657-715, 752, 2114` | ~120 lines dead CSS; dock styled 3× (specificity war). → Delete; restructure dock mobile-first. |
| U11 | Low | — | `styles.css:1041, 1813, various` | 10px accent titles ~3.2:1 contrast; toasts overlap view cube; 'E' double-bound; no keyboard camera; no busy feedback on search/props; pervasive 9-11px type. → Batch with rebrand. |

## P. Performance & delivery (measured)

| ID | Sev | Verdict | Measurement | Issue → Fix |
|----|-----|---------|-------------|-------------|
| P1 | High | ✅ | `index-*.js` = 5,725.61 kB (999.32 kB gzip), 1 chunk | No code-splitting/dynamic imports; everything blocks first paint. → `manualChunks` (three / thatopen / web-ifc) + dynamic-import the IFC load path on first file-open. |
| P2 | High | ✅ | see A2 | Init awaits two external CDNs before usable. → Self-host (A2). |
| P3 | High | ✅ | `viewer.ts:3724-3749, 1220-1227, 1072-1083` | EdgesGeometry rebuilt for every mesh per pointer-move during gizmo drag & per opacity-slider tick — O(triangles)/frame. → Cache per geometry UUID; update matrices only; debounce slider. |
| P4 | Med | ✅ | `viewer.ts:2861-2874` | IFC parse on main thread (worker only does culling) — UI freezes on big models. → Move web-ifc into a worker (Wave 5); mitigate first with progress + chunked yields. |
| P5 | Med | ✅ | dist = 28.20 MB | ~78% dead weight: 21.4MB sample IFCs never fetched by deployed app (e2e reads from disk), 1.26MB unused WASM (used after A2), favicon 0 bytes. → Move samples to `e2e/fixtures/`; real favicon. |
| P6 | Med | ✅ | continuous rAF + heaviest preset default | No on-demand rendering; outlines+gloss+SMAA default; battery drain. → On-demand render mode (esp. embeds); lighter default style. |
| P7 | Med | ✅ | `viewer.ts:1348` | Indexing: hundreds of sequential worker round-trips; browser HTML fully rebuilt multiple times per load. → Parallelize chunks; render once at end. |
| P8 | Med | ✅ | `vite.config.ts:18` | Hardcoded base `/btc-ifc-viewer/`; GH Pages fixed 10-min cache; no vercel.json despite Vercel link — the two targets can't both work. → Env-driven base + vercel.json headers. |
| P9 | Low | — | `viewer.ts:3563+, styles.css` | 26 hotspot style writes per camera frame; 6 backdrop-filter surfaces over live WebGL. → Batch transforms; reduce blur layers (rebrand). |

## T. Testing & CI

| ID | Sev | Verdict | Location | Issue → Fix |
|----|-----|---------|----------|-------------|
| T1 | **Critical** | ✅ | `.github/workflows/deploy.yml:28-29` | Every push to main → build → production. No tests/typecheck/lint/audit. Playwright CI config is dead code. → CI gate (typecheck, lint, e2e) before deploy; PR workflow. |
| T2 | High | ✅ | `tsconfig.json`, `viewer.ts:1139` | `tsc --noEmit` FAILS today (`import.meta.env` needs `"types":["vite/client"]`); nothing anywhere runs tsc. → Fix types entry; add typecheck script; run in CI. |
| T3 | High | ✅ | `package.json` | 7 vulns: critical fast-xml-parser (CVSS 9.3) in RUNTIME @thatopen chain (fix non-major), high rollup + vite (vite fix = major v8), 4 moderate. No Dependabot. → `npm audit fix` now; plan vite major; add Dependabot + CI audit gate. |
| T4 | High | ✅ | `playwright.config.ts:35`, `viewer.ts:1139-1144` | E2E can only test dev server: hooks need `window.__viewer` which is DEV-gated; production artifact untestable; base-path/minification failures invisible. → Gate handles behind explicit `VITE_E2E` define; test `vite build && vite preview`. |
| T5 | Med | ✅ | `e2e/viewer.spec.ts:197-557` | One 557-line test; seven 750ms sleeps; ~20 exact status-copy assertions; effective 36-min timeout; screenshots never compared. → Split into describe blocks w/ fixtures; wait on state not sleeps; assert aria/data attrs. |
| T6 | Med | ✅ | `e2e/viewer.spec.ts:184-189+` | Selection tested via private internals; canvas clicks, keyboard shortcuts, drag-drop, federation multi-model: 0% coverage. → Frozen `window.__viewerTestApi` contract; add real-interaction tests. |
| T7 | Med | ✅ | — | Zero unit tests / no runner. → Vitest (pairs with Wave 2 extraction). |
| T8 | Low | — | `tsconfig.json:19`, `playwright.config.ts:26` | e2e + configs type-checked by nothing; chromium-only. → tsconfig.node.json; add firefox/webkit happy path later. |
| T9 | Low | — | `deploy.yml` | Actions pinned to mutable major tags. → Pin to SHAs. |

## Verified-good (don't "fix")

- `escapeHtml` correctly used in renderModelBrowser / renderFederatedTree / property sections / search result *template* (the F1 crash is pre-escape).
- Strict TypeScript ON; typed state schemas; discriminated unions for modes.
- AbortController listener strategy (for `dom`-cached elements); request-id guard on load path; `fireAndForget` routing.
- ThatOpen v3.3 API usage in Hider/Clipper/Marker/measurements/highlight verified correct against installed lib.
- deploy.yml secrets/OIDC handling is correct; concurrency group sane.
- GLB-relevant: all meshes reachable via three scene graph (export feasible).
