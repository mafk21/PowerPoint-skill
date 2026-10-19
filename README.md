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
and 11 published guides (method & every link: `reference/08-source-research.md`):

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
content (13 slide types, 8 themes — see [`generator/README.md`](generator/README.md)),
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
