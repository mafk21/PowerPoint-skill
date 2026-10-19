# Generator — JSON spec → Morph-enabled .pptx

`generate_pptx.py` builds PlacePlate-style decks (see `../SKILL.md`) with
[python-pptx](https://python-pptx.readthedocs.io/), plus hand-injected Morph
transition XML and `!!` forced-morph shape naming (details in
`../reference/01-morph-transition-mastery.md`).

## Usage

```bash
pip install python-pptx
python generate_pptx.py deck-spec.example.json out.pptx   # build from spec
python generate_pptx.py --themes                          # list themes
python generate_pptx.py --demo out.pptx                   # tiny built-in demo
```

Present the result in **PowerPoint 2019+ / Microsoft 365** — Morph won't play
in Google Slides/Keynote/PDF; export MP4 to share the motion.

After building, grade it: `python ../tools/qa_deck.py out.pptx` (aim: GREEN).
Worked examples: `../examples/specs/*.json` → `../examples/*-demo.pptx`
(pitch, thesis, QBR, Arabic RTL).

## Spec format

Top level:

| Field | Meaning |
|---|---|
| `title`, `author` | Document properties |
| `theme` | One of `teal-orbit`, `aqua-neumorphic`, `royal-blue`, `obsidian-star`, `steel-blue`, `crimson-diamond`, `emerald-ring`, `lilac-soft` (hex table in `../reference/02-design-system.md`) |
| `theme_overrides` | Optional dict to re-skin: any of `bg, card, ink, muted, accent, accent2, graybar, dark` (6-hex, no `#`) |
| `brand` | Footer brand text (appears bottom-left, `!!brand`) |
| `morph_ms` | Default Morph duration (default 900) |
| `timings` | `true` → auto-advance (`advTm`) estimated from speaker notes ~130 wpm (6–45s); slide 1 gets a fade carrying its timing — the deck exports as a correctly-paced video |
| `links` | `true` → clickable deck: agenda items hyperlink to their sections (auto-matched by title/kicker words, or explicit `targets: [slide numbers]` on the agenda slide); any slide may set `link_url` (+ `link_text`); closing `contact` with `@`/http auto-links (mailto/URL) |
| `loop` | `true` → presentation loops honoring timings (presProps `loop`+`useTimings`) — kiosk/trade-show playback; pair with `timings` |
| `rtl` | `true` → Arabic mode: mirrors all geometry, sets RTL paragraph direction + right alignment, flips chevrons/arrows, defaults fonts to Cairo/Tajawal (see `../reference/14-arabic-and-rtl-decks.md`) |
| `numerals` | `"arabic-indic"` → Western digits become ٠١٢٣٤٥٦٧٨٩ in all runs except English (`ltr!`) panels |
| `date` | `YYYY-MM-DD` — `{hijri}` / `{gregorian}` tokens in any string expand from this date (default: today); Hijri is tabular (±1 day vs Umm al-Qura) |
| `heading_font` / `body_font` | Override Montserrat/Inter (runs are remapped post-build) |
| `slides` | Ordered list of slide objects (below) |

Every slide object: `type` plus its fields, and optionally `kicker`
(small top label), `notes` (speaker notes), `morph_ms`, `morph_type`
(`byObject` | `byWord` | `byChar`), `seconds` (per-slide auto-advance override). Slide types and fields — full anatomy &
content limits in `../reference/03-slide-recipes.md`:

| `type` | Fields |
|---|---|
| `title` | `title` (`\n` for line break), `subtitle`, `meta` |
| `agenda` | `title`, `items[]` ≤6, `active` (progress-tracker index, default 0) |
| `section` | `number`, `title`, `subtitle` |
| `content` | `title`, `bullets[]` ≤4 of `{head, body}`, `visual_big` + `visual_caption` **or** `image` (path) |
| `icon_grid` | `title`, `items[]` ≤6 of `{icon or glyph, head, body}` — `icon` keyword (`growth`, `shield`, `bolt`, `check`, `star`…) maps to a glyph automatically; explicit `glyph` wins |
| `process` / `timeline` | `title`, `steps[]` ≤5 of `{head, body}` |
| `comparison` | `title`, `columns[]` ≤3 of `{head, highlight, rows[]}` — prefix rows `"+ "` for ✓, `"- "` for ✕ |
| `chart` | `title` (states the finding!), `categories[]`, `values[]`, `highlight` (index of the accent bar), `series_name`, `insight` |
| `hero_stat` | `stat`, `label`, `context` |
| `market` / `tam_sam_som` | `title`, `rings[]` ×3 of `{label, value, note}` (outer→inner) |
| `team` | `title`, `people[]` ≤4 of `{name, role, bio}` (initials avatar) |
| `quote` | `quote`, `attribution` |
| `closing` | `title`, `ask`, `contact` |
| `stat_strip` | `title`, `stats[]` ≤4 of `{value, label, context}` |
| `before_after` | `title`, `before{head, rows[]}`, `after{head, rows[]}` (✓/✕ added automatically) |
| `logo_wall` | `title`, `names[]` ≤8, optional `caption` |
| `line_chart` | `title`, `categories[]`, `series[]` of `{name, values[]}`, `highlight` (accent series index), `insight` (rail card) |
| `donut` | `title`, `categories[]`/`values[]` ≤4, `highlight`, optional `center` stat (default = % of highlight), `insight` |
| `waterfall` | `title`, `steps[]` ≤7 of `{label, value, kind: start/delta/total}` (deltas signed), optional `format`, `insight` — bars are `!!wf-N` shapes → Morph step-by-step; down-steps vermilion |
| `matrix` | `title`, `quadrants` `[TL,TR,BL,BR]`, `items[]` ≤8 of `{label, x, y (0–1), highlight}`, `x_label`, `insight` (rail, heading via `y_label`) — hero dot `!!mx-N` |
| `funnel` | `title`, `stages[]` ≤5 of `{label, value}`, `insight` — width ∝ value, conversion % between stages, bars `!!fn-N` |
| `bilingual` | `ar_title`/`en_title`, `ar_bullets`/`en_bullets` ≤4 `{head, body}`, optional `ar_body`/`en_body` — Arabic right (rtl=1, Cairo/Tajawal), English left; correct in both RTL and LTR decks (reference/14) |
| `image_full` | `image` (path), `title` (headline on the bottom bar), optional `caption`, `kicker` chip — full-bleed cover-cropped photo; pre-treat with `tools/imgkit.py` (duotone/tint/shade to the theme) |

## What's baked in automatically

- 16:9 (13.333×7.5"), ≥5% margins, Z-pattern, type scale & 3-role palettes
- Persistent chrome on every slide under identical `!!` names — travelling
  orb+ring (giant on title, tucked into corners after), accent bar, kicker,
  brand, page number → continuous Morph motion
- Morph transition XML on slides 2..N (never slide 1)
- Charts: native PPT chart, decluttered, data labels, axis at 0,
  "strategic gray" with one accent bar + insight rail card
- Fonts: Montserrat (headings) / Inter (body) — install these free fonts on
  the presenting machine, or the deck falls back to a default sans

## Verify a build

```bash
python - <<'EOF'
import zipfile
z = zipfile.ZipFile("out.pptx")
for n in sorted(x for x in z.namelist() if x.startswith("ppt/slides/slide")):
    x = z.read(n).decode()
    print(n, "morph:", "morph" in x, "| !! anchors:", x.count("!!"))
EOF
```

Expected: `morph: False` on slide1, `True` everywhere else, `!! anchors` ≥ 6
per slide. Then open in PowerPoint and present (F5).

## Extending

Add a builder function `s_<name>(slide, theme, data, idx, total, brand)`,
register it in `BUILDERS`, name any cross-slide elements with `!!`, and document
the recipe in `../reference/03-slide-recipes.md` (+ row in its mapping table).
New themes: one dict entry in `THEMES` + a row in `../reference/02-design-system.md`.


## Portrait: Reels & carousels

`generator/portrait.py` builds vertical decks (9:16 reel, 4:5 carousel)
with the same themes/Morph/helpers — hook, bigstat, tips, bars, steps,
quote, compare, cta, photo (full-bleed + headline panel). `tools/social.py your-spec.json` converts any
landscape spec into the full pack (both formats + caption + voiceover
shotlist); reels carry `advTm` so PowerPoint's Export → Video is paced
correctly out of the box. QA them like any deck (`tools/qa_deck.py`
accepts 9:16/4:5 canvases).
