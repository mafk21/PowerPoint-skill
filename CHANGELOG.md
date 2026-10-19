# Changelog

## v1.0.0 — 2026-10-19

The system reaches its complete form: research library, generator,
Monster Database, QA gate, creator toolchain, social engine, Arabic Pro,
hardening, showpiece — plus the rescue/media/ship expansion.

### Expansion (Phases 9–11)
- **Phase 11 — Ship & Share:** `datafeed.py` (CSV→chart refresh autopilot),
  `handout.py` (printable leave-behind HTML, RTL-aware), Arabic README
  (`README.ar.md`), GitHub Pages workflow for `site/`, this changelog.
- **Phase 10 — Media & Interaction:** `imgkit.py` (duotone/tint/crop/shade
  photos to the theme); `image_full` + portrait `photo` builders;
  clickable decks (`links: true` — agenda slide-jumps, `mailto:` contacts,
  explicit targets) and `loop: true` kiosk playback; `voiceover.txt`
  export in social packs. *Accent-contrast hardening:* WCAG large-text
  floor corrected; all sub-14pt accent texts raised; 5-theme gauntlet green.
- **Phase 9 — Rescue & Reuse:** `decompile.py` (any .pptx → editable spec;
  ~95% word fidelity on the QBR benchmark) and `merge.py` (cross-deck
  slide transplant; native charts + embedded workbooks survive).
  *Latent bug fixed:* `s_title`/`s_closing` ignored `title_lines` — both
  spellings accepted now, and QA fails on leaked placeholder text.

### Core build (Phases 1–8)
- **Phase 8 — Showpiece:** `htmldeck.py` (spec → playable HTML deck),
  `site.py` → `site/` (live census, theme gallery, playable demo).
- **Phase 7 — Hardening:** `diff.py`, `db_lint.py` + `database/WATCHLIST.md`,
  unittest suite, GitHub Actions CI. Linter caught and repaired two
  built-in themes below the muted-text contrast floor.
- **Phase 6 — Arabic Pro:** `bilingual` slide type (AR right / EN left),
  Arabic-Indic numerals, tabular Hijri dates + `{hijri}` tokens,
  `saudi-green` / `desert-sand` themes, `tools/hijri.py`.
- **Phase 5 — Social Factory:** `generator/portrait.py` (Reel 9:16,
  Carousel 4:5, 8 builders, safe-band layouts) + `social.py`
  (caption + VO shotlist). Example pack in `examples/social/`.
- **Phase 4 — Visual Proof:** QA v2 — text-overflow sentinel, WCAG
  contrast auditor (effective-background), multi-ratio canvas —
  and `render.py` contact-sheet wireframes. Sentinel caught a real
  funnel-label bug before it shipped.
- **Phase 3 — Creator Mode:** `wizard.py` (goal → DB structure → QA'd
  starter deck), `draft.py` (raw notes → deck; prose becomes speaker
  notes), `coach.py` (talk-time, reading-risk, rehearsal cue cards).
- **Phase 2 — Brand & Style:** `palette_lab.py` (seed → WCAG-safe theme),
  `brandkit.py` (logo → theme), `anatomy.py` (.pptx → report + theme);
  external themes in `database/data/themes.json` merge into the generator.
- **Phase 1 — Chart Lab:** `line_chart`, `donut`, `waterfall`, `matrix`,
  `funnel` builders; keyword icon engine; notes-driven auto-timings
  (`advTm`) for video/kiosk pacing.

### Foundation
- **Monster Database:** SQLite + FTS5, 238 records / 12 tables
  (creators, PlacePlate catalog, techniques, structures, principles,
  tools, fonts, prompts, palettes) with `database/query.py` CLI.
- **Generator:** JSON → Morph-first .pptx; 22 landscape builders,
  19 WCAG-checked themes, `!!` travelling anchors, RTL/Arabic engine.
- **Reference library:** 15 documents (~18,700 words) — Morph mastery,
  design system, 22 slide recipes, structures, data-viz, QA, catalog,
  creator atlas, research foundations, advanced techniques,
  storytelling, resources, Arabic/RTL, prompt playbook.
- **Examples:** 7 QA-green decks (pitch, thesis, QBR, Arabic RTL,
  bilingual, social reel + carousel).
