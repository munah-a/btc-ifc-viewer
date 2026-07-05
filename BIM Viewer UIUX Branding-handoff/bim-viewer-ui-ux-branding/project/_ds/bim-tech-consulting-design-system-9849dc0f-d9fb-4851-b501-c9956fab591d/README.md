# BIM Tech Consulting — Design System

**Codename:** "Precision Architect"
**Version:** 1.0 · distilled from the [btc-intranet](https://github.com/munahahmed-btc/btc-intranet) production codebase
**Author:** BTC + Claude, November 2026

---

## 1 · Who BTC is

**BIM Tech Consulting GmbH** is a Swiss/German engineering-consultancy that helps AEC firms adopt **Building Information Modelling (BIM)** and digital workflows. The company runs its own internal intranet — HR, timesheets, projects, finance, CRM, helpdesk — which is the canonical surface this design system was extracted from.

The brand voice lives at the intersection of **Swiss engineering rigour** and **approachable software tooling**:

- *Precise* — numbers are tabular, hierarchies are exact, every pixel earns its place.
- *Calm* — deep navy + warm-lilac neutrals. No decoration for decoration's sake.
- *Trustworthy* — the same lockup an AEC firm expects on a stamped drawing should feel at home on a timesheet approval.
- *Bilingual-ready* — UI is EN/DE, compliance nods to **Swiss nDSG** + **GDPR**.

> The audience is architects, BIM managers, engineers, and HR/Finance staff. They are precise, they are busy, and they do not suffer ornament.

---

## 2 · Content fundamentals

### Voice
Write like a Swiss engineer who has been asked to explain something to a colleague. **Direct, unhedged, no filler.** Avoid marketing adjectives. Prefer nouns and verbs over adjectives and adverbs. Never apologise for the product.

- ✅ "Your timesheet is due Friday."
- ❌ "Just a friendly reminder that you might want to submit your timesheet soon!"

### Tone registers

| Register | Where | Example |
|---|---|---|
| **Operational** | Tables, forms, status. Factual. | `Invoice INV-2026-0142 · Due in 3 days` |
| **Conversational** | Onboarding, empty states, help. Warmer, 2nd person. | "No leave requests yet. Submit one to get started." |
| **Authoritative** | Policy, errors, compliance banners. Terse, unambiguous. | "Access denied. Only @bimtechconsulting.com addresses are allowed." |

### Language rules
- **English is primary** for internal product UI; **German strings ship alongside** via `next-intl`. Never rely on wordplay or idiom.
- **Dates** — default `DD.MM.YYYY` (Swiss). Show relative ("in 3 days") next to absolute, never replace it.
- **Currency** — `CHF 4'240.50` (apostrophe separator, 2 decimals). Never abbreviate to "K" in finance surfaces.
- **Numbers** — tabular-nums always; negative values red, never parenthesised.
- **Sentence case everywhere** except the wordmark. No Title Case Buttons.

### Data slop — avoid
Do not invent KPIs, icons, or decorative stats to fill a layout. An empty widget is better than a lying one. Placeholders must read "—" (em-dash), not "0" or "N/A".

---

## 3 · Visual foundations

### Logo

The BTC mark is a **geometric monogram** built from three flanking plates and a central obelisk — a stylised cross-section of a building core. It reads at 16×16 and at billboard scale.

| Variant | When |
|---|---|
| `logo-primary.svg` | Default, on white/light surfaces |
| `logo-blue.png` | Display on light surfaces — **single-colour brand blue** |
| `logo-white.svg` / `logo-white-v2.svg` | Reversed, on navy or photography |
| `logo-black.svg` | Single-colour print, dark ink |
| `logo-gray.svg` | Disabled / watermark contexts |
| `iconmark.svg` | Favicon, app icon, any square slot < 48px |

**Clearspace:** minimum 1× the obelisk stem on every side. Never recolour the mark outside the palette above. Never place on gradients or busy imagery.

### Palette philosophy

BTC uses a **single chromatic anchor** (the brand blue `#002D7B`) against a **tinted-neutral** ramp that drifts slightly warm-lilac. This is intentional — it distances the product from the ubiquitous "cool grey SaaS" look and references the faint violet you get mixing CAD blueprint inks with paper.

- **Brand Blue** — hierarchy, primary action, links. Never decorative.
- **Magenta tertiary** (`#C70063`) — marks *people and moments* (recruitment, team, welcomes, celebration, marketing). Give it a job, not a sprinkle: it never touches *data and actions* (invoices, dashboards, forms). Scale: `#8C0048` deep → `#C70063` primary → `#FF5C9D` bright → `#FFD3E6` pale.
- **Warm-tinted neutrals** — the entire surface ramp. Warmer than pure grey by 2–3° on the hue wheel.
- **Semantic** — Green/Red/Amber. Standard traffic-light, never decorative.

Full tokens live in `colors_and_type.css`. Do not introduce new colours without mapping them to a `--surface-container-*` or `--primary-*` token first.

### Typography

Two families, no exceptions:

- **Outfit** (300–800) — display, headlines, numeric KPIs. Slightly geometric, technical feel without going into mono.
- **Inter** (300–700) — all UI chrome, body, labels.

**Rules:**
1. Outfit only above 18px, and for tabular numbers of any size.
2. Eyebrow labels are Inter 11px, 500 weight, `letter-spacing: 0.05em`, uppercase.
3. Never stack more than 4 type sizes in a single screen (display · headline · body · label).
4. Line-height `1.6` for body, `1.25–1.35` for headlines, `1` for numeric tiles.

### Spacing & rhythm

Grid is a **4-pixel base unit**. Tokens: `--spacing-1` (4px) through `--spacing-20` (80px). Layouts use multiples of 8 by default; 4 is reserved for tight icon gaps and badge padding.

**Radius ramp:** 8 / 12 / 16 / 24 / full-pill. Avoid arbitrary radii. Cards and widgets settle on `--radius-xl` (24px); inputs on `--radius-md` (12px).

**Elevation is blue-tinted** — shadows carry a 6–10% opacity of brand blue instead of black. This keeps the depth consistent with the tinted-neutral surfaces.

```css
--shadow-ambient: 0 8px 32px  rgba(0, 45, 123, 0.06);
--shadow-float:   0 16px 48px rgba(0, 45, 123, 0.10);
```

### Motion

Three durations, one curve: `150ms`, `250ms`, `350ms` on `cubic-bezier(0.4, 0, 0.2, 1)`. Never animate layout. Never loop animations on idle content.

---

## 4 · Component stance

The intranet is built from **Material-3-flavoured primitives** (tonal surfaces, state layers, pill shapes) re-tinted with BTC's palette. Key patterns:

- **Data tiles** — numeric KPI on top, label below, optional trend sparkline. Always `--radius-lg`, `--surface-container-low` background.
- **Status badges** — pill shape, `label-sm` text, colour-coded against the semantic palette. Variants: `pill` (filled tonal), `outline`, `dot`.
- **Sidebar nav** — fixed 260px, glass background, two-level expansion, chevron rotation as the only motion.
- **Cards over borders** — containment is expressed with surface tone + shadow, never with a 1px line, except in data tables.
- **Forms** — labels above inputs, always. Input height 48px, `--radius-md`. Helper text below, `label-sm`, `on-surface-variant`.

The UI kit pages (`ui-kit-*.html`) document the full set.

---

## 5 · Using this system

### For designers
1. Start from `colors_and_type.css`. It is the source of truth for tokens.
2. Reference the preview cards (`type-preview.html`, `colors-preview.html`, `spacing-preview.html`, `components-preview.html`) for the canonical rendering of each token.
3. For whole-screen mockups, use `ui-kit-intranet.html` as a composition reference.

### For engineers
The production code is at `munahahmed-btc/btc-intranet` on GitHub. Tokens in this repo are a 1:1 mirror of `src/app/globals.css` on `master`. Component class names in the kit match the production CSS modules.

### Do / don't

| Do | Don't |
|---|---|
| Use tokens. Always. | Introduce hex values inline. |
| Use Outfit for numbers, Inter for everything else. | Mix in a third typeface. |
| Use the magenta tertiary for people &amp; moments only (give it a job, not a sprinkle). | Use magenta on data &amp; actions — invoices, dashboards, forms. |
| Keep surfaces tonal (use `--surface-container-*`). | Put a 1px border on every card. |
| Use `—` for empty values. | Pad a dashboard with fake or rounded-to-zero KPIs. |

---

## 6 · File map

```
README.md                      · this document
SKILL.md                       · how to use the system when designing
colors_and_type.css            · all design tokens
assets/                        · logos + LinkedIn banner (SVG + PNG)
type-preview.html              · type scale
colors-preview.html            · palette swatches
spacing-preview.html           · spacing, radius, shadow, motion
components-preview.html        · buttons, inputs, badges, tiles, cards
ui-kit-intranet.html           · whole-screen dashboard composition
brand-external.html            · marketing-surface example (LinkedIn-style)
```
