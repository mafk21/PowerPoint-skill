# 03 — Slide Recipes (layout blueprints)

The reusable slide types that make up PlacePlate decks. A standard PlacePlate
template = **12 unique layouts** picked from this set; the flagship PP-1 has 20;
free samplers have 5. Each recipe lists purpose, layout anatomy (16:9 canvas,
inches), content limits, the Morph behavior to design for, and the matching
`type` in `generator/generate_pptx.py`.

Coordinate system: origin top-left, canvas 13.333 × 7.5 in, margins 0.67 in,
content width ≈ 12.0 in. Headline zone y≈0.9–1.5, accent bar y≈1.52, body zone
y≈2.0–6.6, footer furniture y≈7.0.

---

### R1 · Title / Hero cover — `type: "title"`
- **Purpose:** set visual tone in 3 seconds; bold and memorable.
- **Anatomy:** giant motif orb bleeding off top-right; title 44–54 pt left,
  max 2 lines (x0.67, y1.5, w8.6); one-sentence subtitle 20 pt (what + for whom);
  meta line (event/round/date) in accent caps at y≈5.1; brand bottom-left.
- **Content limits:** title ≤ 8 words; subtitle ≤ 25 words. Nothing else.
- **Morph:** orb starts huge here, then shrinks/tucks into corners on later
  slides (zoom-out). Title text named `!!hero-title` can shrink into slide 2's
  headline position.

### R2 · Agenda / Contents — `type: "agenda"`
- **Purpose:** roadmap; lowers audience cognitive load (academic: mandatory).
- **Anatomy:** headline; 4–6 rows of numbered motif dots (first = accent-filled,
  rest outlined) + section names 19 pt, rows start y2.0, pitch 0.78.
- **Morph:** dots individually `!!`-named (`!!nav-dot-0…`) so they can expand from
  the orb; on later section dividers, highlight the current dot (progress feel).
- **Limits:** ≤ 6 items; each ≤ 6 words.
- **Progress variant:** pass `active: n` to reuse this slide mid-deck as a
  section tracker — completed dots flip to ✓, current dot is accent-filled,
  upcoming items mute. The highlight *walks* across via Morph.

### R3 · Section divider — `type: "section"`
- **Purpose:** chapter break; reset attention. Often paired with a hero stat.
- **Anatomy:** oversized ghost number (80–90 pt accent) y2.3; section title
  40–44 pt y4.0; optional one-line subtitle 18 pt muted. Full-bleed image variant:
  photo edge-to-edge + brand-color overlay + title on top.
- **Morph:** zoom-in hero reveal — previous slide's visual panel grows to fill
  this slide. Use one of your 3–5 "big Morph moments" here.

### R4 · Content split (text cards + visual) — `type: "content"`
- **Purpose:** the workhorse. Problem points, findings, arguments.
- **Anatomy:** headline (assertion). Left column (w7.15): ≤4 stacked cards
  (h1.02, pitch 1.18) each with accent spine, bold head 17 pt + one-line support
  13.5 pt. Right column (x8.25, w4.4, y2.05, h4.55): photo or motif panel with a
  large figure/glyph + caption.
- **Limits:** card head ≤ 8 words; body ≤ 18 words; total slide ≤ ~60 words.
- **Morph:** cards share names across consecutive content slides → they slide/
  settle rather than pop; visual panel persists as `!!visual-panel`.

### R5 · Icon / feature grid — `type: "icon_grid"`
- **Purpose:** 3–6 parallel features/benefits (replaces the bullet list).
- **Anatomy:** headline; grid of cards 3.86 × 2.02 (3 cols, gaps 0.27/0.3),
  each: accent circle chip with glyph top-left, head 16.5 pt bold, body 12.5 pt
  muted ≤ 2 lines.
- **Limits:** 3, 4 or 6 items (never 5 in a 3-col grid — orphan). Head ≤ 4 words.
- **Morph:** cards `!!card-i` + dots `!!card-dot-i` pairwise named, so a following
  detail slide can grow one card full-screen (card-to-detail recipe).

### R6 · Process / Timeline — `type: "process"` (alias `"timeline"`)
- **Purpose:** sequences, methodology, roadmap, how-it-works.
- **Anatomy:** headline; chevron header row (STEP n) at y2.25; a spine line
  (`!!spine`) y3.02; step cards below (h2.2): head 16 pt + body 12.5 pt.
  3–5 steps, equal widths filling 12.0.
- **Morph:** spine stretches between slides; step highlight (accent fill) can
  "walk" across duplicates to show progress. Vertical variant for academic
  methodology flowcharts: same logic rotated.
- **Limits:** ≤ 5 steps; body ≤ 15 words each.

### R7 · Comparison columns — `type: "comparison"`
- **Purpose:** us-vs-them, pricing tiers, before/after, typical-vs-PlacePlate
  (their service page uses exactly this: ✕ list vs ✓ list).
- **Anatomy:** headline; 2–3 cards (h4.5, y2.05); **hero column filled with
  accent**, others neutral card; rows prefixed ✓ / ✕ ( `+ ` / `- ` in spec).
- **Limits:** ≤ 7 rows per column, parallel wording so rows compare cleanly.
- **Morph:** columns `!!cmp-i` — hero card can grow from a prior grid card.

### R8 · Data chart + insight rail — `type: "chart"`
- **Purpose:** any chart whose job is one finding.
- **Anatomy:** headline **states the finding** (never "Revenue by Quarter").
  Chart zone x0.67–8.97 (native PowerPoint chart, decluttered per
  `05-data-visualization.md`); right rail card (x9.35, w3.3): "KEY INSIGHT"
  label + 1–2 sentence so-what.
- **Morph:** duplicate with new values → bars animate; or morph from a hero stat
  into the full chart (number → its own bar highlighted).
- **Limits:** one series focus; ≤ 6 categories visible; highlight exactly one bar.

### R9 · Hero stat — `type: "hero_stat"`
- **Purpose:** the keynote moment; one number says everything. Use at section
  breaks after detail-heavy slides.
- **Anatomy:** number 80–120 pt accent (y1.9); label 24 pt directly under;
  one context line 16–18 pt gray (the comparison that makes it meaningful).
- **Rule:** readable & understood within 3 seconds, before you say a word.
- **Morph:** number scales up from its chart/table cell on the previous slide
  (`!!hero-stat`), or counts visually by stepping through duplicates.

### R10 · Market rings (TAM/SAM/SOM) — `type: "market"`
- **Purpose:** market sizing for pitches (also: nested audiences/scopes).
- **Anatomy:** 3 concentric motif circles left (d 4.6 → 2.04; outer gray, mid
  tint, core accent); right rail rows: value 30 pt + label + note with source.
- **Credibility rule:** bottom-up math + named source; "$100B market" alone
  convinces nobody. Show the path from total market → your slice → revenue.
- **Morph:** rings `!!ring-i` grow outward-in across duplicates (best: SOM first
  full-screen, then zoom out to reveal SAM/TAM around it).

### R11 · Team — `type: "team"`
- **Purpose:** "why us" — people, advisors, partners.
- **Anatomy:** ≤4 cards; photo or initials in accent circle (d1.24); name 17 pt,
  role 13 pt accent caps; 2-line proof bio (domain wins, exits, credentials).
- **Limits:** bio = evidence, not adjectives. Skip anyone not decision-relevant.

### R12 · Quote / testimonial — `type: "quote"`
- **Purpose:** social proof, customer voice, key citation.
- **Anatomy:** oversized accent quote mark; quote 28–30 pt heading font; small
  attribution (name, role/company/country — PlacePlate's own site uses exactly
  this pattern). Optional 5-star row for product contexts.

### R13 · Closing / The Ask — `type: "closing"`
- **Purpose:** end on a specific, confident next step.
- **Anatomy:** big closing line 48–54 pt (mirror the title slide's energy);
  the ask in accent 22 pt (amount + use of funds + milestone, or the decision
  you want); contact/next-step line muted.
- **Never** end on "Questions?" as the visual last impression — keep the ask up
  during Q&A. (Academic variant: conclusion takeaways → references slide after.)
- **Morph:** orb returns to its title-slide hero size — bookend the deck.

### Bonus recipes (build from the above)
- **Roadmap/milestones:** R6 with dates as step heads.
- **Logo wall / partners:** R5 grid with logos instead of glyphs.
- **Pricing:** R7 with price as hero-stat inside hero column.
- **References (academic):** R4 with cards replaced by a clean hanging-indent
  list; smallest allowed text, still ≥ 16 pt; institution citation style.
- **Backup/Q&A slides:** extra R8/R4 slides parked *after* R13 for deep dives
  during questions (academic panels: prepare 2–3).
- **"Parking lot" slide:** one blank-ish slide titled "Questions" to log
  off-scope questions live without derailing timing.

### R14 · Stat strip — `type: "stat_strip"`
- **Purpose:** executive-summary opener / proof wall (QBRs, "year in numbers").
- **Anatomy:** headline; row of 3–4 cards (y2.25, h3.3), each: value 42 pt
  accent, label 15.5 pt bold, one context line muted. Hero-stat logic ×4.
- **Limits:** ≤4 stats; every value needs its label+comparison to mean anything.

### R15 · Before / After — `type: "before_after"`
- **Purpose:** transformations — redesigns, migrations, Duarte's what-is →
  what-could-be in one slide (PlacePlate's service page literally uses this).
- **Anatomy:** left neutral panel with ✕ rows; right **accent-filled** panel
  with ✓ rows; motif arrow between (`!!ba-arrow` travels well).
- **Limits:** ≤6 rows/side, rows parallel so contrast is apples-to-apples.

### R16 · Logo / proof wall — `type: "logo_wall"`
- **Purpose:** social proof — clients, partners, press, certifications.
- **Anatomy:** headline; 4×2 grid of calm cards (wordmark text, or swap in
  real logo images in PowerPoint); optional caption strip beneath.
- **Limits:** ≤8 marks; equal visual weight; permission for real logos.

## Recipe → generator map

| Recipe | spec `type` | Key fields |
|---|---|---|
| R1 Title | `title` | title, subtitle, meta |
| R2 Agenda | `agenda` | title, items[≤6] |
| R3 Section | `section` | number, title, subtitle |
| R4 Content split | `content` | title, bullets[{head,body}≤4], visual_big, visual_caption *or* image |
| R5 Icon grid | `icon_grid` | title, items[{glyph,head,body}≤6] |
| R6 Process | `process`/`timeline` | title, steps[{head,body}≤5] |
| R7 Comparison | `comparison` | title, columns[{head,highlight,rows["+ …","- …"]}] |
| R8 Chart | `chart` | title(finding), categories, values, highlight, insight |
| R9 Hero stat | `hero_stat` | stat, label, context |
| R10 Market | `market` | title, rings[{label,value,note}×3] |
| R11 Team | `team` | title, people[{name,role,bio}≤4] |
| R12 Quote | `quote` | quote, attribution |
| R13 Closing | `closing` | title, ask, contact |
| R14 Stat strip | `stat_strip` | title, stats[{value,label,context}≤4] |
| R15 Before/After | `before_after` | title, before{head,rows}, after{head,rows} |
| R16 Logo wall | `logo_wall` | title, names[≤8], caption |

Every recipe also accepts `kicker` (small section label), `notes` (speaker
notes), `morph_ms` and `morph_type` (`byObject`/`byWord`/`byChar`) overrides.


## The Chart Lab (generator types, reference/05 §6)

### R17 — Trend + insight rail (`line_chart`)
Native line chart, one series in the accent, all others gray; the verdict sits in a persistent insight card right of the plot — the audience reads data → verdict in one eye-sweep (spatial contiguity). Duplicate the slide and change `highlight` to re-tell the same data from the gray series' angle.

### R18 — Donut with center stat (`donut`)
Max four segments, one accent, the takeaway *inside* the ring as a 34pt stat (`!!donut-center` — it can fly in from a big_number slide). Custom legend with values + % — never the default legend.

### R19 — Shape-built waterfall (`waterfall`)
Bars are shapes (`!!wf-N`), not a chart object, because charts can't Morph bar-by-bar: duplicate the slide, delete the last bar, and rebuild order = reveal order. Down-steps in Okabe vermilion (semantic + colorblind-safe), totals in accent, connectors thin gray.

### R20 — 2×2 matrix (`matrix`)
Quadrant *names* do the arguing ("Leaders", "Loud & shallow"); one hero dot, rivals gray; `!!mx-N` dots let you Morph the market shifting across quarters by duplicating with new coordinates.

### R21 — Funnel with conversion callouts (`funnel`)
Bar width ∝ value; the persuasive payload is the % between stages (the leak), set in the accent. Bar color grades accent→accent2 so depth reads as progression, not category.

> Auto-timings: set `timings: true` on the spec (or `seconds` per slide) — the generator estimates advance times from speaker notes (~130 wpm) and writes `advTm` into the transitions, so the same deck exports as a correctly-paced video (reference/13 §6).


### R22 — Bilingual twin panels (`bilingual`)
For Saudi/Gulf enterprise rooms where the deck must serve Arabic-first and English-reference readers at once: Arabic panel on the right (reading order starts there), English twin left, same bullet order so eyes can cross-check line by line. Keep both sides terse — a bilingual slide carries 2× words by definition (QA warns "dense"; that's the one warn you're allowed to nod at). Full mechanics in reference/14 (Arabic Pro section).
