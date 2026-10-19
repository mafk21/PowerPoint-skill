# 11 — Advanced PowerPoint Techniques (beyond basic Morph)

The creator-ecosystem's move library (`09`), distilled into buildable recipes.
Ordered by ROI for PlacePlate-style decks. Each: effect → build steps →
watch-outs. Assumes `01` (Morph + `!!` naming) is mastered.

## A. Morph, leveled up

**A1 · One-canvas "scrolling world" (mega-slide pan)**
Design ONE giant scene (e.g. 4 slides wide worth of content parked off every
edge). Each duplicate shifts the whole group left/up; Morph pans the "camera".
The Prezi-killer look in many viral reels. *Steps:* build scene on slide 1 with
all panels beside the canvas → duplicate → drag group so panel 2 is framed →
repeat → Morph on all. *Watch:* group objects; mind file size with photos.

**A2 · 3D model rotation morph** (yerra.dzn / One Skill signature)
Insert → 3D Models → same model on 2 slides, rotate/reposition on slide 2 →
Morph spins it live (M365). Product reveals, anatomy, engineering. *Watch:*
playback needs modern PowerPoint; fallback = static frame on old versions.

**A3 · Photo zoom / un-crop & Ken Burns**
Slide 1: photo cropped tight on a detail. Slide 2: crop relaxed to full frame
(or vice versa). Morph animates the crop box — cinematic push-in without video.
Also works on **video objects** (enlarge video across duplicate = push-in).

**A4 · Color/blur state shift** (superslideppt pattern)
Duplicate → on copy, change picture brightness/blur (artistic filter) or shape
fills → Morph cross-grades the scene. Great for "focus" moments: background
blurs as the hero stat rises.

**A5 · Logo reveal** (the 60-second viral build)
Photo background → black rectangle at ~40% transparency over it → logo as
**cut-out mask** (SVG logo + Merge Shapes ▸ Subtract from a rect, or logo shape
filled with slide-background) → duplicate → scale mask/logo massively →
Morph. The logo "opens" onto the photo. Variant: text-as-window (giant word
filled with image, morph-scales to reveal full-bleed photo).

**A6 · Rotation/spin motifs** (deskzenzone pattern)
Wheel/pinwheel shapes rotated 30–45° per duplicate + Morph = continuous spin
illusion. Use for cycles, processes, "360°" concepts. Keep center pinned
(rotate around center exactly, same size/pos).

**A7 · Image slider / carousel strip**
Row of photo cards spanning off-canvas; each duplicate shifts strip by one
card; center card scales up, neighbors shrink + desaturate. Portfolio/product
galleries. (Our `!!card-i` naming makes card identities track perfectly.)

**A8 · Counter & progress walks**
Progress bar/dots: duplicate repeatedly, advance fill one step each (slide.sage
dynamic-layout pattern — duplicate elements, change color/number, Morph all).
Numbers count via Characters-mode text morph steps (10→20→…→final).

## B. Navigation & interactivity (One Skill's territory)

**B1 · Zoom hub menu** — Insert ▸ Zoom ▸ **Summary/Section/Slide Zoom** drops
live thumbnails on a hub slide; click = animated dive into that section,
auto-return after. Build PlacePlate-style: hub = R2 agenda with Zoom
thumbnails inside motif frames instead of plain dots. For non-linear talks,
Q&A hopping, training menus. *Watch:* Zoom objects need "Return to Zoom"
format checked to bounce back; test in Slide Show, not edit view.

**B2 · Trigger pop-ups** — Animations ▸ Trigger ▸ On Click Of [object]: a
button/dot reveals hidden detail *on the same slide* (map hotspots, FAQs,
quiz answers). Pair entrance+exit triggers on the same button to toggle.

**B3 · Kiosk mode** — Slide Show ▸ Set Up Show ▸ *Browsed at a kiosk*: advance
keys dead, only your hyperlinks/Zooms navigate; loops forever. Trade-show
screens, self-serve explainer decks, sent interactive modules.

**B4 · Action buttons & mouse-over previews** — Insert ▸ Action: hover a menu
card → triggered preview pops before committing to click. Menu UX from B1+B2.

## C. Animation craft (when object animation beats Morph)

- **Animation Pane discipline**: order, `With/After Previous`, delays in 0.1 s
  steps; stagger fades 0.08–0.15 s apart for a "deal-in" cascade.
- **Motion paths + Ink Replay**: draw with pen → replay stroke-by-stroke for
  hand-drawn arrows/underlines (teacher aesthetic).
- **Masked animations**: object slides *behind* a foreground panel (shape with
  slide-background fill) so it appears to emerge from an edge/card.
- **Parallax**: 2–3 background layers moving at different speeds/distances
  across Morph duplicates = depth. (One Skill parallax tutorial pattern.)
- **Book flip**: page shape + 3D rotation key states + Morph between them.
- **GIF & cinemagraph backgrounds**: subtle looped motion under low-opacity
  overlay for title slides; keep < ~10 MB and kill if it fights the text.
- **Animation Painter / BrightSlide stagger**: clone a tuned animation stack
  across objects instead of rebuilding (consistency machine — see `13`).

## D. Shape & image surgery (how creators fake "Illustrator")

- **Merge Shapes** (Union / Combine / Fragment / Intersect / **Subtract**) —
  custom motifs, cut-out text windows, gear/ring diagrams, logo masks (A5).
- **Edit Points + freeform** for organic blobs; **Convert to Shape** on
  SmartArt/icons to recolor parts separately.
- **SVG icons**: Insert ▸ Icons (native library) or paste SVG → right-click
  *Convert to Shape* → ungroup & recolor to theme accent. One style per deck.
- **Image treatment stack**: crop-to-shape (motif-shaped photos!), recolor to
  duotone (Picture Format ▸ Color), transparency for overlay blends, soft-edge
  fades at panel boundaries.
- **Smart guides & distribution**: Alt-drag duplicates; Align/Distribute;
  `Ctrl+D` duplicate-and-repeat transform (PowerPoint remembers the offset —
  hit Ctrl+D repeatedly for instant grids/rows).
- **Eyedropper** (fill ▸ Eyedropper, incl. long-press outside the window)
  to sample brand colors from client logos/photos into the 3-role system.

## E. Speed workflow (pros' invisible advantage)

- **Quick Access Toolbar** stuffed with: Align, Distribute, Merge Shapes,
  Selection Pane, Format Painter, Morph (QAT + shortcuts = minutes saved).
- **Selection Pane as project manager**: rename everything (`!!` discipline),
  toggle visibility to edit buried layers, lock background art.
- **Slide Master surgery once**: logo, furniture, theme colors/fonts,
  placeholder styles — then never hand-fix again (their "Smart Color System").
- **Rehearse → Record**: Slide Show ▸ Record for timed narration + ink;
  Export ▸ Create Video (1080p/4K) *is* the reel pipeline (`08` formula).
- **Designer (AI) & Copilot**: use for ugly-first-draft layouts/content, then
  re-skin into the design system — AI output never ships raw (`13`).
- Versioning: keep a `_BACKUP` section at deck end; Hide Slide for variants
  instead of separate files.

## F. Morph compatibility traps (new since `01`)

- **Charts don't morph internally** — chart *objects* move/scale, but data
  morphs only when the chart is the *same object* duplicated (values
  interpolate in M365; older versions cross-fade). For guaranteed bar growth
  in hand-built decks, draw bars as rectangles with `!!` names.
- **Tables/SmartArt** morph as whole objects only — explode to shapes if
  cells must move independently.
- **Grouped objects** morph as units; `!!` on the group name works, members
  inside won't individually match across different groups.
- **Fonts missing on presenting machine** silently reflow text (and wreck
  tight morphs) — embed fonts (File ▸ Options ▸ Save ▸ Embed fonts) when the
  license allows, or ship Calibri-safe.
- Morph + heavy media = slideshow jank on weak hardware: compress pictures
  (Picture Format ▸ Compress, 150–220 ppi for projection), trim video.


## Implemented: clickable decks in the generator

Set `"links": true` in a spec and agenda items become real slide-jump
hyperlinks (`ppaction://hlinksldjump` + slide relationships), auto-matched
to sections by title/kicker vocabulary or pinned with `targets: [...]` on
the agenda slide. `link_url`/`link_text` put external links on any slide;
closing emails become `mailto:`. `"loop": true` stamps presProps
loop+useTimings for kiosk playback with the auto-timings. Section-Zoom
thumbnails still need native PowerPoint, but hyperlink navigation — the
working core of menu decks — is now generated, QA-able, and covered by
tests. Photos join the system too: `tools/imgkit.py` (duotone/tint/crop/
shade to the theme) + the `image_full` and portrait `photo` builders.
