# PowerPoint Skill — PlacePlate Style 🎞️✨

A complete, reusable **AI skill for generating professional PowerPoint decks**,
researched from **[PlacePlate](https://www.instagram.com/placeplate12345/)**
(Instagram `@placeplate12345` · [placeplate.com](https://placeplate.com)) — the
111K-follower studio famous for **Morph-transition PowerPoint templates**.

> Present with Clarity. Design with Impact. — their line, and now this repo's
> operating principle.

## What was learned (and where it lives)

Since Instagram blocks scraping, the research reconstructed the full picture
from their cross-posted TikTok reels, YouTube previews, 47-product web catalog
and 11 published guides (method & every link: `reference/08-source-research.md`) —
then expanded to **40+ peer creators, the peer-reviewed research base, the
classic storytelling frameworks, and the full tool ecosystem**:

| File | What you'll get |
|---|---|
| **[`SKILL.md`](SKILL.md)** | ⭐ The skill itself — 10 style laws + 7-step generation procedure |
| [`reference/01-morph-transition-mastery.md`](reference/01-morph-transition-mastery.md) | Morph workflow, the `!!` forced-morph trick, 5 pro recipes, XML internals |
| [`reference/02-design-system.md`](reference/02-design-system.md) | "1 motif + 1 color story", 8 hex palettes, type scale, text budgets |
| [`reference/03-slide-recipes.md`](reference/03-slide-recipes.md) | 13 layout blueprints with coordinates & content limits |
| [`reference/04-deck-structures.md`](reference/04-deck-structures.md) | Story arcs: 12-slide pitch, thesis defense, business, sales, seminar, QBR |
| [`reference/05-data-visualization.md`](reference/05-data-visualization.md) | Chart choice, decluttering, strategic gray, hero stats |
| [`reference/06-qa-checklist-and-mistakes.md`](reference/06-qa-checklist-and-mistakes.md) | The 7 fatal mistakes + pre-flight QA gate |
| [`reference/07-template-catalog.md`](reference/07-template-catalog.md) | All 47 PlacePlate products, prices & links + service tiers |
| [`reference/08-source-research.md`](reference/08-source-research.md) | Creator profile, every URL, reel formula, honesty notes |
| [`reference/09-creator-atlas.md`](reference/09-creator-atlas.md) | 🗺️ 40+ creators like PlacePlate (Jacob 5M, Luis Urrutia 4M, One Skill…) & what to steal |
| [`reference/10-research-foundations.md`](reference/10-research-foundations.md) | 🔬 The science: Mayer, Cognitive Load, assertion–evidence studies, debunked myths |
| [`reference/11-advanced-powerpoint-techniques.md`](reference/11-advanced-powerpoint-techniques.md) | 🚀 3D morph, one-canvas pans, Zoom hubs, triggers, kiosk, parallax, logo reveals |
| [`reference/12-storytelling-and-persuasion.md`](reference/12-storytelling-and-persuasion.md) | 📖 Duarte sparkline/STAR, Minto pyramid + SCQA, Knaflic, Presentation Zen |
| [`reference/13-resources-tools-and-assets.md`](reference/13-resources-tools-and-assets.md) | 🧰 Add-ins, AI-tool reality check, marketplaces, assets, reel/carousel specs |
| [`reference/14-arabic-and-rtl-decks.md`](reference/14-arabic-and-rtl-decks.md) | 🌙 Arabic & RTL decks: mirroring, Cairo/Tajawal, BiDi, bilingual patterns |
| [`reference/15-prompt-playbook.md`](reference/15-prompt-playbook.md) | 💬 12 copy-paste AI prompts, incl. one that emits generator-ready JSON |

## 🐲 The Monster Database

Everything above is also a **real queryable database** — `database/powerpoint.db`
(SQLite + FTS5, **236 records / 12 tables**: creators, templates, techniques,
recipes, structures, research, tools, fonts, prompts, palettes):

```bash
python database/query.py recommend pitch      # structure + themes + SKUs, joined
python database/query.py search "logo reveal" # full-text across everything
python database/query.py templates --free     # the 3 free PlacePlate samplers
python tools/qa_deck.py examples/arabic-pitch-demo.pptx   # automated QA gate
```

Showcase decks in `examples/` (all QA-GREEN, Morph-verified): investor pitch,
thesis defense, QBR decision deck, and a full **Arabic RTL pitch**
(`"rtl": true` → mirrored geometry, RTL paragraphs, Cairo/Tajawal) — specs in
`examples/specs/`. Database manual: [`database/README.md`](database/README.md).

## Try it in 30 seconds

```bash
pip install python-pptx
python generator/generate_pptx.py generator/deck-spec.example.json demo.pptx
# → 12-slide pitch deck with REAL Morph transitions (slides 2–12)
#    open it in PowerPoint 2019+ / Microsoft 365 and press F5
```

`examples/placeplate-style-demo.pptx` is that exact output, committed as proof.
Verified at the XML level: every slide from #2 carries
`<p14:morph type="byObject"/>`, and the signature travelling elements share
`!!` names across slides so they glide instead of fading.

To make your own deck: copy `generator/deck-spec.example.json`, swap in your
content (16 slide types, 8 themes — see [`generator/README.md`](generator/README.md)),
rerun. Or just ask an AI agent working in this repo — `AGENTS.md` points it at
the skill automatically.

## The 10-second style summary

16:9 · Morph everywhere · one motif + one accent color · headlines that state
findings · ≤3 short bullets or a diagram instead · huge hero stats · charts in
gray with one accent bar · 3–5 big motion moments, quiet continuity between.

## Attribution

Style, techniques and catalog facts: **PlacePlate** —
[Instagram](https://www.instagram.com/placeplate12345/) ·
[TikTok](https://www.tiktok.com/@placeplate) ·
[Website](https://placeplate.com/) ·
[Linktree](https://linktr.ee/PlacePlate).
This repo is an original synthesis for learning/generation; it contains none of
their files. Their templates cost $3.50–$7 (three are free) — if you want an
original, buy it via the links in `reference/07-template-catalog.md`.
