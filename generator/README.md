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

## Spec format

Top level:

| Field | Meaning |
|---|---|
| `title`, `author` | Document properties |
| `theme` | One of `teal-orbit`, `aqua-neumorphic`, `royal-blue`, `obsidian-star`, `steel-blue`, `crimson-diamond`, `emerald-ring`, `lilac-soft` (hex table in `../reference/02-design-system.md`) |
| `theme_overrides` | Optional dict to re-skin: any of `bg, card, ink, muted, accent, accent2, graybar, dark` (6-hex, no `#`) |
| `brand` | Footer brand text (appears bottom-left, `!!brand`) |
| `morph_ms` | Default Morph duration (default 900) |
| `slides` | Ordered list of slide objects (below) |

Every slide object: `type` plus its fields, and optionally `kicker`
(small top label), `notes` (speaker notes), `morph_ms`, `morph_type`
(`byObject` | `byWord` | `byChar`). Slide types and fields — full anatomy &
content limits in `../reference/03-slide-recipes.md`:

| `type` | Fields |
|---|---|
| `title` | `title` (`\n` for line break), `subtitle`, `meta` |
| `agenda` | `title`, `items[]` ≤6 |
| `section` | `number`, `title`, `subtitle` |
| `content` | `title`, `bullets[]` ≤4 of `{head, body}`, `visual_big` + `visual_caption` **or** `image` (path) |
| `icon_grid` | `title`, `items[]` ≤6 of `{glyph, head, body}` (glyph = 1–2 chars/emoji shown in the chip) |
| `process` / `timeline` | `title`, `steps[]` ≤5 of `{head, body}` |
| `comparison` | `title`, `columns[]` ≤3 of `{head, highlight, rows[]}` — prefix rows `"+ "` for ✓, `"- "` for ✕ |
| `chart` | `title` (states the finding!), `categories[]`, `values[]`, `highlight` (index of the accent bar), `series_name`, `insight` |
| `hero_stat` | `stat`, `label`, `context` |
| `market` / `tam_sam_som` | `title`, `rings[]` ×3 of `{label, value, note}` (outer→inner) |
| `team` | `title`, `people[]` ≤4 of `{name, role, bio}` (initials avatar) |
| `quote` | `quote`, `attribution` |
| `closing` | `title`, `ask`, `contact` |

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
