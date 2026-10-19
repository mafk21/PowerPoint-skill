# 01 — Morph Transition Mastery (the core technique)

PlacePlate's entire catalog and reel library is built on one feature: PowerPoint's
**Morph transition**. Master this file and you can reproduce their look. Distilled
from their published Morph guide + tutorial reels; rewritten as generation rules.

## What Morph is

Morph is a **slide transition** (not an object animation). PowerPoint compares two
consecutive slides, finds matching objects, and automatically animates position,
size, rotation, color, crop and text between the two states — smooth, physics-like
motion with zero keyframing. It replaces dozens of manual entrance/exit animations.

## Compatibility (hard constraints — design around them)

| Platform | Create | Play |
|---|---|---|
| Microsoft 365 (Win/Mac) | ✅ | ✅ |
| PowerPoint 2019 / 2021 / 2024 | ✅ | ✅ |
| PowerPoint for Web, iOS, Android | ✅ | ✅ |
| PowerPoint 2016 | ❌ | ✅ (playback only) |
| PowerPoint 2013 or older | ❌ | ❌ |
| Google Slides / Keynote / PDF | ❌ | ❌ (transitions stripped) |

**Delivery rule:** present from the `.pptx`, or export **MP4 video** to share the
motion. Send PDF only as a read-along, never as the presentable.

## The 5-step workflow (how every PlacePlate slide pair is built)

1. **Duplicate** the slide (right-click → Duplicate Slide, or `Ctrl/Cmd + D`).
2. **Edit the duplicate only** — move, resize, recolor, rotate, re-crop, retype.
   Whatever changed is what Morph will animate.
3. **Select the second slide** in the slide panel.
4. **Transitions tab → Morph.** (The transition belongs to the *destination* slide.)
5. **Preview & tune** — set Effect Options and Duration (see below).

Golden rules that follow from this workflow:

- **Never put Morph on slide 1** — there is no "before" state; it degrades to a fade.
- **Never delete an object to remove it** — drag it *off the slide canvas* instead.
  Morph then flies it out smoothly; deleting makes it pop/fade abruptly. (This is
  the most-repeated tip across Morph tutorial content.)
- **Always duplicate first, then edit the copy** — editing the original destroys
  your starting state.
- Keep object **identity**: Morph matches same-type, similar-looking objects. If two
  shapes look identical, matching can swap them — differentiate slightly or use `!!`.

## Effect Options (Transitions → Effect Options)

- **Objects** (default) — whole elements move/transform. Use for 90% of slides.
- **Words** — text morphs word-by-word. Great for headline swaps where shared
  words travel to their new position.
- **Characters** — letter-by-letter. Dramatic title reveals, anagram-style moves.

**Duration:** PlacePlate-style decks feel "cinematic but snappy": **0.8–1.2 s** for
content transitions, up to **1.5–2.0 s** for hero/section reveals. Set the same
duration deck-wide for rhythm. (Our generator defaults to 900 ms.)

## The `!!` forced-morph naming trick (the pro move)

By default Morph only matches visually similar objects. To morph *unlike* objects
(circle → rectangle, icon → photo, small dot → full menu), force the match:

1. Select the object → **Home → Arrange → Selection Pane**.
2. Double-click its name, rename to `!!` + a shared token, e.g. `!!hero-shape`.
3. Give the partner object on the other slide the **exact same** `!!name`.
4. PowerPoint now morphs one into the other regardless of type.

Without matching `!!` names, unlike objects cross-fade instead of transforming —
the #1 cause of "Morph isn't working" complaints.

### The travelling element (PlacePlate's signature)

Name one shape identically (`!!hero-orb`) on **every** slide, changing only its
position/size. It becomes a continuous "floating element" that glides through the
whole deck — the exact effect their reels show off. Our generator does this
automatically (orb + ring + accent bar + kicker + footer on all slides).

## The 5 professional Morph recipes

1. **Zoom-in hero reveal** — object small on slide N, full-bleed on slide N+1.
   Use for section openers and product reveals.
2. **Expanding navigation** — one dot on slide N becomes a row of `!!`-named
   dots/labels (menu/agenda) on N+1.
3. **Animated charts** — duplicate the chart slide, change values; bars/lines
   grow and shrink. (Works best with shape-built bars or same chart object.)
4. **Cinematic text** — Effect Options → Words/Characters; title travels from
   center stage to its resting corner as content fades in around it.
5. **Scrolling infographic** — design a canvas taller than the slide; each
   duplicate shifts it up; Morph simulates a smooth scroll — no video needed.

Bonus patterns seen in their preview videos: **card-to-detail** (grid card expands
into its own slide), **image un-crop** (tight crop relaxes to full photo),
**color-story shift** (same layout, accent color sweeps to the next section's hue).

## Dosage: how much Morph is right?

Motion is attention. PlacePlate templates morph *structurally* (chrome elements
persist everywhere) but reserve **big moves for 3–5 key moments**: opening reveal,
section changes, the data climax, the closing ask. If every slide does a stunt,
nothing stands out — and presenters fighting their own animation timing look amateur.

Practical pattern used by our generator:

- **Every slide (from #2):** Morph on, `byObject`, ~0.9 s — quiet continuity.
- **Hero moments:** enlarge/reposition the `!!` anchors so the same transition
  reads as a deliberate zoom/travel.
- **Text stunts:** switch single slides to `byWord`/`byChar` sparingly.

## Morph troubleshooting checklist

| Symptom | Cause | Fix |
|---|---|---|
| Object fades instead of moving | Objects don't match (different type/look) | Force with identical `!!name` on both slides |
| Wrong objects swap places | Two identical-looking objects matched wrongly | `!!name` each pair distinctly (`!!card-1`, `!!card-2`…) |
| Slide 1 "Morph" does nothing | No previous slide | Start Morph from slide 2; put a pre-state slide first if needed |
| Transition missing when presenting | Opened in Google Slides/Keynote/PDF export | Present from PowerPoint 2019+ or export MP4 |
| Motion feels sluggish | Duration too long / too many huge moves | 0.8–1.2 s; shorten travel distances |
| Text jumps instead of gliding | Text edited inside different text boxes | Keep the same text box across duplicates; try Words mode |

## Generating Morph decks with code (how this repo does it)

python-pptx has no transition API, so `generator/generate_pptx.py` writes the exact
XML PowerPoint itself uses, as a child of each slide (after `<p:cSld>`):

```xml
<p:transition spd="med" p14:dur="900">
  <p14:morph type="byObject"/>   <!-- byObject | byWord | byChar -->
</p:transition>
```

(`p14` = `http://schemas.microsoft.com/office/powerpoint/2010/main`.)
Shape identity for forced morphs is just the shape's `name` attribute
(`<p:cNvPr name="!!hero-orb">`) — set `shape.name = "!!…"` in code, identically
on both slides. Verify any generated file by unzipping it and grepping
`ppt/slides/slide*.xml` for `morph` and `!!`.

*Attribution: technique as taught publicly by PlacePlate —
https://placeplate.com/blog/powerpoint-morph-tutorial/*
