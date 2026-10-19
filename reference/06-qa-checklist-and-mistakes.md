# 06 — QA Checklist & Mistake Fixes

Run this before declaring any deck done. Merged from PlacePlate's "7 common
mistakes", Morph-guide pitfalls and design-tips articles — rewritten as a
pass/fail gate. A deck ships when every box passes (or a failure is a deliberate,
documented choice).

## The 7 classic mistakes → fixes

1. **Too many slides.** Attention is a budget; every slide withdraws from it.
   → Keep-10 test; 10–15 slides per 20 min; overflow → appendix.
2. **Walls of text.** Audience reads *or* listens — never both.
   → ≤3 bullets/slide, 5–7 words each; slides = anchors, notes = script.
3. **Reading slides verbatim.** Makes the presenter redundant.
   → Speaker Notes + Presenter View; present the insight, slide holds evidence.
4. **Fonts too small.** <18 pt dies beyond the front row; <14 pt is invisible.
   → 20–28 pt body min (their stricter guides say 24–28), 36 pt+ headings;
     stand-2-meters-from-the-monitor test.
5. **Inconsistent design.** 5 fonts + 3 palettes = "threw this together", and
   audiences read it as inconsistent *thinking*.
   → 2 fonts, 1 accent, Slide Master locks it in. (Why templates exist.)
6. **No visual hierarchy.** Equal weight everywhere = eye has no instructions.
   → Most important element = biggest + highest contrast; squint test.
7. **Distracting transitions.** Fly-ins/spin/checkerboard = 2003 energy.
   → One transition, Morph, 0.8–1.2 s, everywhere from slide 2.

## Pre-flight QA checklist

**Story**
- [ ] Headlines alone (slide-sorter view) tell the complete story in order
- [ ] One idea per slide; anything failing the one-sentence test was split
- [ ] Proof/traction appears early (first third), not buried at the end
- [ ] Deck ends on a specific ask / decision / "so what" — not bare "Q&A/Thanks"
- [ ] Slide count ≤ minutes available × ~1 (appendix excluded)

**Design**
- [ ] 16:9; margins ≥ 5% width; nothing critical in bottom 10%
- [ ] Exactly 3 color roles (bg/ink/one accent); categories via tints only
- [ ] ≤ 2 sans-serif fonts; no thin weights; contrast checked (esp. gray-on-white)
- [ ] Size floors respected (body ≥ 20 pt general / 24 pt academic; never < 16 pt)
- [ ] Same icon style, same chart style, same card style throughout
- [ ] Images full-bleed or in cards, tight crops, unified with overlay; sharp at 100%
- [ ] Squint test passes on every slide (hero element identifiable when blurred)

**Text**
- [ ] ≤ 3 bullets × ≤ 7 words (pitch) / 6×6 outer limit (academic)
- [ ] Headlines are assertions ("Churn fell below 2%"), not labels ("Results")
- [ ] Body text left-aligned; bold used for key terms only
- [ ] Numbers checked; claims sourced on-slide where credibility matters

**Data** (full rules in `05-data-visualization.md`)
- [ ] Right chart for the comparison; no 3D; no pie > 3 slices
- [ ] Decluttered: no legend/border, light-or-no gridlines, direct labels
- [ ] Strategic gray: exactly one accent bar/series per chart
- [ ] Bar axes start at 0; axes labeled with units; insight/`so-what` present

**Morph & motion** (full rules in `01-morph-transition-mastery.md`)
- [ ] Morph on every slide from #2; slide 1 has none (or a pre-state exists)
- [ ] Cross-type morphs forced with matching `!!names`; identical twins `!!`-named distinctly
- [ ] Exiting objects dragged off-canvas on the next slide, not deleted
- [ ] Duration consistent (0.8–1.2 s; heroes ≤ 2 s); big moves only at 3–5 moments
- [ ] Object animations: Fade/Appear only, 0.3–0.5 s, each with a teaching purpose

**Delivery readiness**
- [ ] Will be presented from PowerPoint 2019+/M365 (Morph dies in Slides/Keynote/PDF)
- [ ] Sharing version exported as **MP4** (motion) and/or PDF (reading only)
- [ ] Speaker notes written; Presenter View tested; 2–3 backup slides after closing
- [ ] Full-screen run-through done at presentation distance & timed
- [ ] Fonts are free/standard (or fallback-safe) on the presenting machine
- [ ] File opens clean: `.pptx`, all text live/editable, no missing-media warnings

## Fast triage table (symptom → first fix)

| "It looks amateur because…" | First fix |
|---|---|
| Cluttered, tiring | One idea per slide + delete 30% of elements |
| Boring walls of bullets | Convert: steps→process, features→grid, compare→cards |
| Can't tell what matters | Enlarge + accent the key element; gray the rest |
| Feels inconsistent | Lock Master: 2 fonts, 1 accent, fixed furniture |
| Data doesn't land | Assertion headline + strategic gray + insight rail |
| Motion feels gimmicky | Kill object animations; Morph only, shorter, sparser |
| Too long | Keep-10 test; appendix the rest; merge section+stat slides |


## QA v2 — what the analyzer now measures for you (tools/qa_deck.py)

Two checks that used to require opening PowerPoint are now automated:

- **Overflow sentinel** — estimates rendered text height (glyph-advance
  model ≈0.50em, 0.58em all-caps; line height from paragraph spacing;
  margin-aware — `Length(0)` is falsy, so zeroed margins need an
  `is not None` test) versus each box. `>1.12x` = likely spill (FAIL),
  `1.0–1.12x` = tight (WARN). It caught a real bug on day one: narrow
  funnel bars squeezing their labels until they wrapped.
- **Contrast auditor** — measures every colored run against its
  *effective* background (own fill → nearest filled shape beneath
  containing its center → slide bg) and applies WCAG floors (4.5:1 text,
  3:1 large/bold). Runs ≤2 chars are treated as graphics (dingbat chips).
  FAIL only when the background was confidently determined.
- Canvas check accepts 16:9, 9:16, 4:5, 1:1 and 4:3 (social formats).

And `tools/render.py deck.pptx` emits a standalone HTML **contact sheet**
— real geometry/fills/type as a wireframe proof sheet, charts/images as
labeled summary blocks, Morph map + word counts per slide. Eyeball it
before you ever open PowerPoint; it's also the fastest way to share a
preview or eyeball two versions side by side.
