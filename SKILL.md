---
name: powerpoint-placeplate-style
description: >-
  Generate professional PowerPoint (.pptx) presentations in the PlacePlate
  (@placeplate12345) style — Morph-transition decks with one motif + one color
  story, assertion headlines, hero stats, strategic-gray charts and story-first
  structures for pitch decks, thesis/academic defenses, business and sales
  presentations. Use whenever the user asks to create, design, restyle or
  improve a PowerPoint/PPTX/slide deck, or asks about Morph transitions,
  PlacePlate templates, or presentation design.
---

# PowerPoint Generation — PlacePlate Style Skill

Distilled from PlacePlate (@placeplate12345 on Instagram; placeplate.com):
111K-follower Morph-template studio, 47-product catalog, 11 published guides,
~3,000 claimed clients. Research provenance & links: `reference/08-source-research.md`.
Full catalog: `reference/07-template-catalog.md`. **Read the reference files for
depth; this file is the operating procedure.**

Reference map (open on demand — don't guess what you can look up):

| File | Contains | Open when |
|---|---|---|
| `reference/01-morph-transition-mastery.md` | Morph workflow, `!!` trick, 5 pro recipes, troubleshooting, XML | building/fixing any deck |
| `reference/02-design-system.md` | motif+color formula, palettes (hex), type scale, text budgets | choosing look & feel |
| `reference/03-slide-recipes.md` | 13 layout blueprints w/ coordinates & generator mapping | composing slides |
| `reference/04-deck-structures.md` | story arcs: pitch (12), thesis, business, sales, seminar, QBR | outlining any deck |
| `reference/05-data-visualization.md` | chart choice, decluttering, strategic gray, hero stats, tables | any data slide |
| `reference/06-qa-checklist-and-mistakes.md` | 7 fatal mistakes + pre-flight gate | before delivering |
| `reference/07-template-catalog.md` | all 47 PlacePlate SKUs, prices, links, service tiers | recommending/buying originals |
| `reference/08-source-research.md` | who PlacePlate is, every link, reel formula, method notes | citing sources, making reels |
| `reference/09-creator-atlas.md` | 40+ creators/agencies (Jacob 5M, Luis Urrutia 4M, One Skill…), what to steal from each | "make it like X", new techniques |
| `reference/10-research-foundations.md` | Mayer, Cognitive Load, assertion–evidence studies, memory effects, debunked myths | defending design choices with science |
| `reference/11-advanced-powerpoint-techniques.md` | 3D morph, one-canvas pans, Zoom hubs, triggers, kiosk, parallax, logo reveals | leveling up beyond basic Morph |
| `reference/12-storytelling-and-persuasion.md` | Duarte sparkline/STAR, Minto pyramid+SCQA, Knaflic, Garr Reynolds — when to use which | structuring persuasive decks |
| `reference/13-resources-tools-and-assets.md` | add-ins (BrightSlide/think-cell), AI tools reality check, marketplaces, assets, social specs | tooling, buying, exporting to social |
| `reference/14-arabic-and-rtl-decks.md` | Arabic RTL: mirroring, fonts (Cairo/Tajawal), BiDi, bilingual patterns | Arabic / Saudi / bilingual decks |
| `reference/15-prompt-playbook.md` | 12 copy-paste LLM prompts incl. JSON-spec generator chain | working faster with AI |

Generator: `generator/generate_pptx.py` (JSON spec → real Morph-enabled .pptx;
16 slide builders, 8 themes, **RTL/Arabic mode** via `"rtl": true` in spec).
QA analyzer: `tools/qa_deck.py`. Prompt chain: `reference/15-prompt-playbook.md`.

**🐲 The Monster Database** (`database/powerpoint.db`, 227 records / 12 tables /
FTS5): query it *before* web-searching — creators, PlacePlate SKUs, techniques,
recipes, structures, research principles + myths, tools, fonts, prompts,
palettes. `python database/query.py search "<anything>"`,
`recommend <pitch|thesis|business|sales|seminar|qbr|decision>`,
`templates --free`, `prompt 6`, or raw `sql "SELECT …"`. Rebuild after editing
`database/data/*.json`: `python database/build_db.py` (manual: `database/README.md`).
Example spec: `generator/deck-spec.example.json` → `examples/placeplate-style-demo.pptx`.

---

## Non-negotiable style laws (the "PlacePlate look")

1. **Morph is the transition** — on every slide from #2, 0.8–1.2 s, deck-wide.
   No checkerboard/fly-in/spin antiques. Present from PowerPoint 2019+; export
   MP4 to share motion (Morph dies in Google Slides/Keynote/PDF).
2. **One motif + one color story per deck** (`[Color] + [Motif]`, e.g. Teal
   Orbit). The motif recurs as decoration, markers and the travelling element.
3. **3 color roles only**: background, ink, one saturated accent. Categories get
   tints of the accent, never new hues.
4. **One idea per slide**; headline is an **assertion** ("Q3 revenue hit a
   3-year high"), never a label ("Results"). Headlines alone tell the story.
5. **Text budget**: ≤3 bullets × ≤7 words (pitch) — 6×6 absolute max (academic).
   Slides anchor; speaker notes carry the script.
6. **Type**: ≤2 sans-serif fonts (default Montserrat + Inter; PP-1 ships
   Onest/Inter). Body ≥20 pt (academic ≥24), titles 36–44, nothing visible <16 pt.
7. **Layout**: 16:9, ≥5% margins, nothing critical in bottom 10%, Z-pattern
   (key item top-left, takeaway bottom-right), whitespace is premium.
8. **Replace bullets with structure**: steps→process, features→icon grid,
   compare→side-by-side cards, one big number→hero stat (80–120 pt).
9. **Charts are decluttered** and use **strategic gray** — only the bar that
   proves the headline gets the accent color. Axes start at 0. No 3D, no pie >3.
10. **Dosage**: quiet Morph continuity everywhere; **3–5 big motion moments**
    (opener, section reveals, data climax, ask). More = gimmick.

## Operating procedure for any deck request

### Step 1 — Classify & gather (ask only what's missing)
Determine: **use case** (investor pitch / thesis-academic / business-company
profile / sales-proposal / seminar / business review), **audience & tone**,
**minutes available** (→ slide count ≈ minutes, cap by structure in `04`),
**content** (user's material or draft it), **brand colors/logo** if any.
If the user gave enough, don't interrogate — build and state assumptions.
Arabic/bilingual request? Set `"rtl": true` in the spec (mirrors geometry,
sets RTL paragraphs, flips chevrons/arrows, swaps Cairo/Tajawal) and follow `14`.

### Step 2 — Pick the design identity
From `02-design-system.md`, choose motif + theme (or map user brand color into
the 3-role system). Defaults: startup pitch → `teal-orbit`; thesis →
`aqua-neumorphic` or `royal-blue`; premium business → `obsidian-star`; SaaS →
`steel-blue`. If a PlacePlate SKU fits exactly and the user wants *that file*,
recommend buying it ($3.50–$7, links in `07`) and offer to fill it — never
claim our output is their file.

### Step 3 — Outline with the right story arc
Use `04-deck-structures.md`. Write **assertion headlines first**, in order;
read them back — they must narrate the whole argument. Front-load proof
(strong traction moves to slide 3–4). Park excess in appendix/backup slides.

### Step 4 — Compose slides from recipes
Compose only from the 16 recipes in `03-slide-recipes.md` (R1 title, R2 agenda
±progress variant, R3 section, R4 content split, R5 icon grid, R6 process/
timeline, R7 comparison, R8 chart+insight, R9 hero stat, R10 TAM/SAM/SOM,
R11 team, R12 quote, R13 closing/ask, R14 stat strip, R15 before/after,
R16 logo wall). Enforce each recipe's content limits ruthlessly — cut words,
don't shrink fonts. For narrative shape beyond section order (contrast
rhythm, answer-first pyramids), apply `12-storytelling-and-persuasion.md`.

### Step 5 — Build the .pptx
Preferred path: **write a JSON spec and run the generator**:
```bash
pip install python-pptx          # once
python generator/generate_pptx.py my-deck.json my-deck.pptx
```
The generator bakes in: 16:9, theme palettes, type scale, persistent `!!`
chrome (travelling orb/ring, accent bar, kicker, brand, page no.), Morph XML
on slides 2..N, strategic-gray charts, speaker notes. Spec format =
`generator/deck-spec.example.json` + field table in `03`. Extend the generator
(new builders/themes) rather than hand-rolling one-off code when possible.

Hand-building in the PowerPoint UI instead? Follow `01`: duplicate-slide →
edit copy → Transitions ▸ Morph on the destination slide; `!!name` pairs via
Selection Pane for cross-type morphs; drag exiting objects off-canvas.

### Step 6 — QA gate (mandatory)
Run `python tools/qa_deck.py deck.pptx` (automated form of
`06-qa-checklist-and-mistakes.md`: 16:9, Morph map, `!!` anchors, word budgets,
font floors/families, headline heuristic, color discipline, notes coverage) —
ship only on GREEN (0 fails), then eyeball the remaining checklist boxes the
tool can't judge. Fix failures before delivery; report deliberate deviations
(e.g. Calibri fallback, dense appendix exhibits).

### Step 7 — Deliver correctly
- Hand over the **.pptx** (presentable) + note: "Present in PowerPoint
  2019+/M365 to see Morph; export MP4 to share the motion; PDF is read-only."
- Summarize: structure used, theme, slide count, where the 3–5 Morph moments are.
- If the user wants social content: the reel formula (`08`) — export slideshow
  to MP4, 9:16 crop, caption `Morph PowerPoint Template ✨🔥` + the standard
  hashtag block, CTA to link-in-bio.

## Morph quick-reference (full detail in `01`)

- Workflow: **duplicate → edit the copy → select copy → Transitions ▸ Morph**.
- Never on slide 1. Never delete-to-remove (drag **off-canvas** instead).
- Effect Options: Objects (default) · Words (headline swaps) · Characters
  (dramatic titles). Duration 0.8–1.2 s (heroes ≤ 2 s).
- **`!!` naming trick**: same `!!name` on both slides forces a morph between
  unlike objects; identical twins need distinct `!!` names or they may swap.
- Signature move: one `!!` shape repositioned on **every** slide = travelling
  element gliding through the deck.
- 5 pro recipes: zoom-in hero reveal · expanding nav dots · animated charts ·
  cinematic text · scrolling infographic.
- In code: add `<p:transition spd="med" p14:dur="900"><p14:morph
  type="byObject"/></p:transition>` after `<p:cSld>`; set `shape.name="!!…"`.

## Data quick-reference (full detail in `05`)

Ask "what comparison?" → categories=bar · time=line · part-whole=donut ≤4 ·
correlation=scatter · exact lookup=designed table (≤5×7) · one number=hero stat.
Declutter: no legend/border, direct labels, light-or-no gridlines, axis at 0.
One accent element per chart; everything else gray.

## Attribution & boundaries

Techniques and public product facts are PlacePlate's (placeplate.com · Instagram
@placeplate12345 · TikTok @placeplate); all wording here is original synthesis.
This skill recreates the *style* — it must not reproduce their template files,
long verbatim text, or imply affiliation. When their original template is the
fastest path, say so and link it (`07-template-catalog.md`).
