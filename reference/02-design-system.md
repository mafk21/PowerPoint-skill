# 02 — The PlacePlate Design System

How their slides look, codified into rules you (or the generator) can apply to any
topic. Distilled from their design-tips articles, product pages, site copy
("clean, modern, consistent… structured storytelling") and preview imagery.

## The master formula: 1 motif + 1 color story

Every PlacePlate SKU is **one geometric motif × one dominant color family** —
that's literally their naming system:

`[Color/Material] + [Motif]` → Teal **Orbit** (rings), Emerald **Ring**, Crimson
**Diamond**, Silver **Blade**, Amber **Radial**, Azure **Bloom**, Golden **Lotus**,
Obsidian **Star**, Aqua **Neumorphic** (soft embossed cards)…

**Generation rule:** before building any deck, pick (a) a motif shape that will
recur as decoration, section markers and the Morph "travelling element"
(circle/orbit, diamond, blade/diagonal, radial rays, bloom/petals), and (b) a
single color story. Never mix two motifs or two accents in one deck.

## Canvas & layout law

- **16:9 widescreen only** (13.333 × 7.5 in). Every PlacePlate file follows it.
- **Margins ≥ 5% of width** (≈0.67 in) on all four sides. Nothing critical in the
  **bottom 10%** (projector bezels / video-call overlays crop it).
- **Z-pattern reading path**: key element top-left → scan right → diagonal →
  takeaway/CTA bottom-right. Headline top-left, always.
- **Whitespace is a feature**: generous gaps between headline and body; if an
  element isn't earning its place, delete it. Premium = empty, not full.
- **One idea per slide.** If a slide can't be summarized in one sentence, split it.
- **Rounded-corner cards everywhere** — content lives on soft rectangles floating
  on the background, with the motif bleeding off the edges. (Confirmed across
  their preview videos: white/soft cards, circular diagram hubs, big % badges.)
- **Visual hierarchy test ("squint test")**: blur/squint at the slide — the main
  message must still be identifiable. Biggest + highest-contrast = most important.

## Color discipline (3 roles, nothing else)

| Role | Rule |
|---|---|
| Background | White/off-white, very dark navy, charcoal, or a deep tint of the accent |
| Text (ink) | Near-black on light, near-white on dark — check contrast (WebAIM) |
| Accent | **One** saturated color: headings' details, highlights, CTAs, motif, the one chart bar that matters |

6+ colors = amateur. Need categories? Use **shades/tints of the accent**, not new
hues. Red only for negative/warning data, green only for positive/growth.

### Built-in palettes (generator themes)

Reconstructions in the spirit of the named SKUs (see honesty note in
`08-source-research.md`). `bg / card / ink / accent`:

| Theme | Mode | bg | card | ink | accent | Best for |
|---|---|---|---|---|---|---|
| `teal-orbit` (PP-23) | dark | `04333F` | `0B4A5A` | `F2FBFA` | `14C8B4` | startup pitch |
| `aqua-neumorphic` (PP-20) | light | `E9F5F5` | `FFFFFF` | `0E3737` | `0AA9A0` | thesis defense |
| `royal-blue` (PP-43) | light | `FFFFFF` | `EEF3FB` | `10295E` | `1D4ED8` | formal academic/corporate |
| `obsidian-star` (PP-49) | dark | `0B0B10` | `17171F` | `F5F2EA` | `D4A537` gold | premium pitch/keynote |
| `steel-blue` (PP-46, free) | dark | `0F2233` | `17344D` | `F2F7FC` | `4FA3E3` | SaaS/investor |
| `crimson-diamond` (PP-39) | dark | `1A0A10` | `2A1219` | `FDF2F4` | `E11D48` | bold business |
| `emerald-ring` (PP-38) | dark | `06231A` | `0B3527` | `ECFCF5` | `10B981` | growth/sustainability |
| `lilac-soft` (PP-44) | light | `F6F2FC` | `FFFFFF` | `35245E` | `8B5CF6` | soft creative/education |

**Smart color system:** PlacePlate's better templates let you re-skin the whole
deck from the Slide Master (Design → Variants → Colors). When hand-building,
define palette as Master/theme colors, never hard-coded one-offs — same reason
our generator themes are single dictionaries.

Tone matching (from their template-selection guide):

- **Academic/thesis** → muted, structured, hierarchical. Nothing "flashy".
- **Business/client** → polished neutral dark or light + one strong accent.
- **Investor pitch** → minimal, credibility-first; every slide intentional.
- **Creative/portfolio** → the only place to experiment harder with color/layout.

## Typography law

- **Max 2 fonts**: one hero font for headings, one supporting font for body.
  Both **sans-serif** (serifs render poorly on projectors).
- PlacePlate ships **free fonts**: PP-1 uses **Onest / Inter**; their guides also
  endorse Inter, Montserrat, Poppins, Calibri. Generator defaults:
  Montserrat (headings) + Inter (body). Use Calibri only if the deck must open
  perfectly on machines without fonts installed — and say so.
- **Size scale (hard floors):**

| Element | Size |
|---|---|
| Hero stat | 80–120 pt |
| Slide title / headline | 36–44 pt (title slide up to ~54) |
| Section heading | 28–42 pt |
| Body text | 20–24 pt preferred; **never below 16 pt** visible (academic: 24 pt min body, 18 pt absolute floor) |
| Captions / labels | 16–18 pt (generator uses 11–13 pt only for footer furniture & chart labels on-screen; raise for lecture halls) |

- **Left-align body text.** Centered body copy scans badly and reads informal.
- **Bold for key terms** — not italics, not ALL-CAPS paragraphs.
- Avoid thin weights (100–300): they vanish on projectors. Design for the back row.
- **Headlines are assertions, not labels.** "Q3 revenue hit a 3-year high" beats
  "Results". A deck should tell its story through headlines alone.

## Text budget law

- **≤ 3 bullets per slide**, each **5–7 words** (max ~30 words/slide on pitch decks).
- Academic **6×6 rule** as an outer limit: ≤6 bullets × ≤6 words.
- Slides are **visual anchors, not transcripts**. Overflow goes in **speaker notes**
  (presenter view), never shrunk onto the slide.
- Replace bullets with structure wherever possible:

| Content type | Use instead |
|---|---|
| Sequence of steps | process/flow diagram |
| 3–6 features | icon-and-label card grid |
| Comparison | side-by-side cards/table |
| Single key number | hero-stat slide |
| Data with a finding | chart with assertion title + insight rail |

## Images & iconography

- **Full-bleed** images on section dividers for impact; elsewhere inside rounded cards.
- Unify photos with a **semi-transparent brand-color overlay**.
- **Tight crops** on the subject beat wide generic scenes. No clipart, no blurry
  images (judge at 100% zoom in edit view).
- **One icon style** per deck, inside identical chips/circles (motif carriers).
- Their infographic vocabulary (from previews): circular hub/gear diagrams,
  donut % badges, orbit rings, radial segments, chevron processes, big-number
  stat blocks. Prefer these native-shape infographics over pasted screenshots.

## Animation law (beyond Morph)

- Animation exists to **direct attention**, not entertain. Every move must answer
  "what does this help them understand?"
- **Fade/Appear** for progressive disclosure; skip Fly-In/Bounce/Spin relics.
- Duration **0.3–0.5 s** for object animations; **0.8–1.2 s** Morph between slides.
- **One transition, deck-wide**: Morph. (See `01-morph-transition-mastery.md`.)

## Consistency machinery

- Build on **Slide Master + layouts**: logo once on the master, fonts/colors once
  in theme. Inconsistency signals carelessness → kills credibility (their core
  selling point is that templates enforce this automatically).
- Deck furniture (brand, page number, motif) sits in the **same place on every
  slide** — and under the **same `!!` names** so Morph treats it as continuous.
- File rules: `.pptx`, 16:9, PowerPoint 2019+ (Win/Mac), fonts free/standard,
  everything editable (no flattened text-as-image), speaker notes welcome.
