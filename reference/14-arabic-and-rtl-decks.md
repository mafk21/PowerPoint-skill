# 14 — Arabic & RTL Decks (عروض تقديمية بالعربية)

Building PlacePlate-style decks in Arabic (or bilingual AR/EN). Arabic is not
"English but mirrored text" — it has its own typographic physics. Rules below
adapt `02`/`03` without breaking the design system.

## The #1 mistake (and the correct mental model)

**Right-aligning text ≠ RTL.** You must set **paragraph direction** to
Right-to-Left (Home ▸ Paragraph ▸ Right-to-Left; in shape text options too).
Alignment alone leaves bullets on the wrong side, scrambles mixed Arabic +
Latin/number order, and puts final punctuation at the line start. PowerPoint
handles the Unicode BiDi algorithm (Arabic letters flow RTL; Latin words,
URLs, phone numbers stay LTR *inside* the line) — but only when paragraph
direction is set correctly.

## Mirror the layout, not just the text

| Element | LTR (EN) | RTL (AR) |
|---|---|---|
| Headline anchor | top-**left** | top-**right** |
| Reading path | Z-pattern ↘ | **mirrored Z** ↙ (start top-right, end bottom-left) |
| Bullets / numbers | left edge | right edge |
| Progress, timelines, process flow | left → right | **right → left**; flip arrows (← = "next") |
| Agenda dots / chrome | brand bottom-left | brand bottom-right; page no. bottom-left |
| Chart category axis | left → right | mirror layout, but keep time series natural — test with native readers; labels always Arabic-side |

**Never mirror:** logos, photos of people, clocks, media controls, the digits
inside a number, scientific notation/formulas, maps.

## Arabic typography law (adapts `02`)

- **Use a real Arabic font.** Arial/Calibri/Times "support" Arabic with ugly
  kerning and missing ligatures. Free professional set (Google Fonts, OFL):
  **Cairo** (geometric headlines — the Montserrat of Arabic), **Tajawal**
  (clean body workhorse), **IBM Plex Sans Arabic** (tech/corporate),
  **Almarai** (Saudi-market standard), **Noto Sans Arabic** (universal
  fallback), **Amiri** (Naskh serif — formal/academic/religious tone),
  Dubai Font (gov). Pair: Cairo ↔ Inter/Montserrat (match x-height/stroke).
- **Size up ~+2 pt and line-height 1.35–1.5×** vs Latin equivalents: Arabic
  glyphs carry more detail, extend above/below the line, and diacritics
  (tashkeel) need vertical room. Arabic at the "same pt" reads smaller.
- **Letter-spacing = 0. Always.** Arabic is a connected script; tracking
  breaks letter joins and renders text illegible. (If an EN style uses
  tracked caps for kickers, drop tracking in AR — use bold + accent instead.)
- **No italics** (Arabic has no true italic; faux-oblique looks broken) —
  emphasize with **bold or accent color**. No ALL-CAPS equivalent either.
- **Numerals**: choose per market — Western digits (123) common in Saudi/tech
  business; Arabic-Indic (١٢٣) in formal/government/educational contexts.
  Be consistent; dates: state calendar (Gregorian/Hijri) explicitly if shared.
- Calligraphic display fonts (Thuluth, Ruqaa styles) = decoration only
  (title slide word-art, Ramadan/Eid motifs) — never body text.

## Morph in RTL (what changes)

- Travelling elements should travel **right→left** (orb path mirrored) so
  motion agrees with reading direction.
- Chevron/process shapes must point **left**; "next" indicators flip.
- `!!` naming, duplication workflow, durations: unchanged.
- Test mixed-direction lines (Arabic + "PowerPoint" + 2026) on every text
  slide — BiDi surprises hide there.

## Bilingual deck patterns

1. **One language per slide** (duplicate deck, flip fully) — cleanest for gov/
   formal; maintain one master content table to keep versions in sync.
2. **Split slides** — AR right column, EN left column, shared central visual —
   for mixed audiences; halve text budgets (laws still apply per column).
3. **Arabic-first with English captions** — headlines AR (assertion), small EN
   translation as caption under visuals — common in academic defenses.
   Reverse for multinational business decks presented in English.

## Cultural/visual notes (Saudi/Gulf business context)

- Green carries religious/national weight — powerful accent, use with respect;
  gold+dark (our `obsidian-star`) reads premium-formal; avoid color choices
  that fight a client's government/brand identity colors.
- Geometric girih/arabesque motifs map beautifully onto the PlacePlate
  "motif" system (star, radial, lotus → eight-point star "khatam") — one
  motif, tastefully repeated, never a clip-art mashup.
- Imagery: prefer local/regional photography when presenting to Saudi
  audiences; follow client/venue norms on people imagery; mosque/religious
  imagery only when contextually appropriate and treated with care.
- Formal decks open with a title slide carrying both Arabic and English
  names of presenter/institution — standard in Saudi universities.

## Implementation notes for the generator

✅ **Implemented** (`"rtl": true` in the spec): the generator runs a polish
pass per slide — mirrors every shape's x (`x' = slideW − (x + w)`), sets
`rtl="1"` on each paragraph + right alignment (centered text stays centered),
flips chevrons/arrows via `flipH`, and remaps Montserrat→Cairo / Inter→Tajawal
(override with `heading_font`/`body_font`). Verified on
`examples/arabic-pitch-demo.pptx` (128 RTL paragraphs, 5 flipped directionals,
Cairo/Tajawal only). Known v1 limits: native chart internals (axis direction,
chart fonts) stay LTR — PowerPoint applies complex-script fallback so Arabic
labels shape correctly, but eyeball charts in the app; grouped shapes are
mirrored as units. Still QA every mixed-direction line (Latin brand names,
numbers, URLs) per the BiDi rules above.


## Implemented in the generator (Arabic Pro)

- **`bilingual` slide type** — Arabic panel right (reads first), English
  left, hairline divider. AR shapes are named `ar!…` and get `rtl="1"` +
  Cairo/Tajawal in a dedicated pass; EN shapes are named `ltr!…` and the
  RTL polish mirrors their *position only* (never direction/fonts), so the
  same slide is correct in RTL and LTR decks. Fields: `ar_title`/
  `en_title`, `ar_bullets`/`en_bullets` (`{head, body}` ≤4), `ar_body`/
  `en_body`.
- **`"numerals": "arabic-indic"`** — converts Western digits to ٠١٢٣٤٥٦٧٨٩
  in every run *except* `ltr!` (English-panel) shapes. Works in RTL and
  bilingual decks, landscape and portrait. Tip: write stats like `87%` —
  conversion yields the correct reading order automatically in RTL runs.
- **Hijri dates** — put `{hijri}` (or `{gregorian}`) in any spec string;
  it expands from `spec.date` (YYYY-MM-DD) or today, using the arithmetical
  (tabular/Kuwaiti) algorithm in `generate_pptx.gregorian_to_hijri`
  (verified: 2026-10-19 → 7 جمادى الأولى 1448هـ; 622-07-19 → 1-1-1).
  Tabular Hijri may differ ±1 day from Umm al-Qura — fine for covers.
  CLI: `python tools/hijri.py [YYYY-MM-DD]`.
- **Saudi themes** in `themes.json`: `saudi-green` (flag green #006C35 —
  auto-brightened to #007B3D on dark for contrast — + sand-gold) and
  `desert-sand` (light bronze + Saudi green secondary). 19 themes total.
- QA font rule is per-script now: ≤2 Latin + ≤2 Arabic families passes.
- Demo: `examples/arabic-bilingual-demo.pptx` (spec:
  `examples/specs/arabic-bilingual.json`) — verified in XML: 15 rtl
  paragraphs on the bilingual slide, EN paragraphs untouched, fonts split
  Cairo/Tajawal vs Montserrat/Inter, hero stat rendered ٨٧.
