# 15 — Prompt Playbook (copy-paste AI prompts for deck work)

Battle-tested prompt patterns for using an LLM (this agent, Copilot, ChatGPT…)
at each stage. Copy, fill the `[brackets]`, paste. Each prompt encodes the
house rules so output lands PlacePlate-shaped the first time. Designed to pair
with `generator/generate_pptx.py` (several prompts emit its JSON spec directly).

## 1 · Big Idea extractor (before anything else)

```
Act as a presentation strategist trained in Nancy Duarte's method.
Topic: [TOPIC]. Audience: [WHO]. I want them to [DECISION/ACTION].
Give me 5 candidate "Big Idea" sentences. Each must: state a unique point of
view, convey what's at stake, be one complete sentence. Then rank them and
flag the strongest. No slide content yet.
```

## 2 · SCQA opener (decision decks — Minto)

```
Write an SCQA opening for my deck. Rules: Situation = 1 sentence the audience
already agrees with. Complication = 1 sentence on what changed/broke.
Question = the implied question. Answer = my recommendation in 1 sentence —
this becomes my governing thought. Maximum 4 sentences total.
Context: [PASTE]. My recommendation: [PASTE].
```

## 3 · Pyramid storyline (ghost deck)

```
Structure this as a Minto pyramid. Line 1: the single-sentence answer.
Then 3–4 supporting arguments — they must be MECE (no overlap, no gaps).
Under each, bullet the evidence that would prove it. Then convert it into a
ghost deck: one complete-sentence action title per slide (≤15 words, active
voice, quantified where possible), in order. Material: [PASTE MESSY DRAFT].
```

## 4 · Assertion-headline converter (fixes any lazy deck)

```
Rewrite these topic slide titles as assertion headlines: each a full sentence
stating the slide's conclusion (max 2 lines), specific and quantified.
Read together they must tell the whole story (horizontal logic).
Titles + their slide content: [PASTE].
```

## 5 · Sparkline storyteller (pitch/keynote)

```
Outline a [X]-minute talk on [TOPIC] using Duarte's sparkline: open on
"what is" (≤10%), call to adventure (the gap), then alternate what-is vs
what-could-be beats, one STAR moment around 70% (suggest 3 options: shocking
stat, evocative visual, soundbite), end with a "new bliss" picture + one
explicit call to action. Output: slide-by-slide list with beat label
[IS]/[COULD BE]/[STAR]/[CTA], assertion headline, and visual idea per slide.
```

## 6 · Full deck-spec generator (feeds our generator directly) ⭐

```
You know the JSON format in generator/deck-spec.example.json and the recipes
in reference/03-slide-recipes.md. Create a complete spec for: [USE CASE —
pitch/thesis/business/sales/…] about [TOPIC], audience [WHO], [N] slides,
theme [THEME]. Enforce: assertion headlines; ≤3 bullets × ≤7 words;
≤~30 words/slide; one idea per slide; chart titles state findings; include
speaker notes on every slide; end with a specific ask. Output ONLY valid
JSON. Source content: [PASTE NOTES]. 
```
Then: `python generator/generate_pptx.py spec.json deck.pptx` → QA with `06`.

## 7 · Data-slide surgeon

```
Given this data [PASTE TABLE]: (1) tell me the single most defensible finding
in one sentence — that becomes the slide headline; (2) pick the chart type
(bar/line/donut/scatter/table/hero-stat) and justify in one line; (3) apply
strategic gray: which ONE bar/series gets the accent and why; (4) write the
1–2 sentence "KEY INSIGHT" for the right rail; (5) list anything on the raw
chart to delete (legend, gridlines, clutter).
```

## 8 · Bullet-to-visual converter

```
Convert these bullets into slide visuals. For each group, choose: process
diagram / icon-card grid / comparison cards / hero stat / chart — and write
the exact short labels (heads ≤4 words, bodies ≤12 words). No bullet lists
may survive. Bullets: [PASTE].
```

## 9 · Speaker-notes writer

```
Write presenter notes for each slide: what to SAY (talk track, conversational,
~60–90 seconds per slide), one transition sentence into the next slide, and
one likely audience question with a crisp answer. Slides carry assertions
only — the script lives here. Slides: [PASTE TITLES+CONTENT].
```

## 10 · Deck critic (the brutal review)

```
Review this deck outline against: Mayer's coherence/signaling/redundancy,
assertion–evidence format, horizontal + vertical logic, reverse storyboard.
For every slide output: headline verdict (assertion or label?), words to cut,
one concrete fix. Then give the top 5 changes ranked by impact. Be ruthless.
Deck: [PASTE OUTLINE OR SLIDE TEXT].
```

## 11 · Reel script (PlacePlate formula — `08`/`13`)

```
Write a 20-second Instagram/TikTok reel script showing off this deck's Morph
transitions (no face-cam). Give: 3 hook options for on-screen text ("Did you
know PowerPoint could do this?" style), shot list (which slide flips, 2–3 s
each, STAR moment last), caption (≤150 chars) ending with a link-in-bio CTA,
and the hashtag block (#powerpoint #morphtransition …). Deck theme: [X].
```

## 12 · Arabic version pass (`14`)

```
Localize this slide text to professional Saudi-business Arabic (not literal
translation). Keep assertion headlines ≤2 lines, bullets ≤7 words, Western
numerals unless I say otherwise. Flag any phrase that needs cultural
adaptation, and give the matching English caption for bilingual slides.
Remind me which layout elements must mirror. Text: [PASTE].
```

## Chaining recipe (the full pipeline)

`1 Big Idea → 2 SCQA (if decision) or 5 Sparkline (if pitch/talk) → 3 Pyramid
titles → 8 visual conversion → 7 data surgery → 6 JSON spec → generator →
9 notes → 10 critic → 06 QA gate → deliver (+ 11 reel if marketing).`

A human/agent following this chain produces a deck that satisfies Duarte,
Minto, Knaflic, Mayer and PlacePlate simultaneously — which is the point.
