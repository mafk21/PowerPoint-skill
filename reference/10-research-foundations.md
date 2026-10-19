# 10 — Research Foundations: the science under the style

PlacePlate's rules (and this skill's laws) aren't taste — they line up with
decades of cognitive research. This file gives the generation rules *and their
receipts*, so decks can be defended to any stakeholder, examiner or client.
Effect sizes are Cohen's *d* (≈0.2 small, 0.5 medium, 0.8+ large).

## 1 · Richard Mayer's Cognitive Theory of Multimedia Learning

The most-tested framework in the field (200+ experiments; *Multimedia
Learning*, Cambridge UP; 15 principles in the 2021 3rd ed.). The slide-relevant
core, with median effects from Mayer's reviews:

| Principle | Generation rule | Evidence |
|---|---|---|
| **Coherence** | Delete anything that doesn't serve the message — decorative images, logos-per-slide clutter, seductive details | 23/23 tests, *d*≈0.86 |
| **Signaling** | Cue what matters: accent color on ONE element, bold key terms, kickers/headlines that preview structure | 24–26/28 tests, *d*≈0.41–0.70 |
| **Redundancy** | Never put your script on the slide and read it. Visual + spoken words beat visual + spoken + identical on-screen text | 16/16 tests (fast-paced), *d*≈0.86 in 2014 review |
| **Spatial contiguity** | Labels ON/next to the chart parts they describe (direct labels; no legend round-trips) | 22/22 tests, *d*≈1.10 |
| **Temporal contiguity** | Reveal visual + narration together (progressive disclosure timed to speech; Morph keeps context continuous) | 9/9 tests, *d*≈1.22–1.30 |
| **Segmenting** | Break content into digestible, paced chunks — one idea per slide, section dividers, agenda returns | strong, repeatedly replicated |
| **Pre-training** | Define key terms/components before the complex model (glossary/agenda mini-slide) | moderate |
| **Modality** | With complex visuals, explain by *voice*, not paragraphs | most-studied principle |
| **Multimedia** | Words **+** relevant pictures beat words alone | *d*≈1.67 (5 tests) |
| **Personalization/Voice** | Conversational wording; human (not robotic) voice-over in videos | moderate |
| **Image** | A decorative speaker photo/avatar does *not* help learning — spend pixels on evidence | ≈null — cut it |

**Mapping to this skill:** coherence → one-idea law & whitespace; signaling →
3-color discipline + strategic gray; redundancy → text budgets + speaker notes;
spatial contiguity → `05` direct labels; segmenting → `04` structures.

## 2 · Cognitive Load Theory (John Sweller et al.)

Working memory is tiny (~3–4 interacting elements for new material). Loads:

- **Intrinsic** — the real difficulty of the content (can only be sequenced).
- **Extraneous** — load created by *bad design* (fully under our control).
- **Germane** — effort that builds schema (reframed by Sweller 2010 as
  resources aimed at intrinsic load; design goal: free capacity for it).

Slide-critical effects (Chandler & Sweller, 1991 onward):

- **Split-attention effect** — diagram here + explanation there forces mental
  integration and burns memory. *Fix:* physically integrate (annotate on the
  visual; keep insight rail adjacent — but never make the chart unintelligible
  without the text).
- **Redundancy effect** — the flip side: text that merely repeats a
  self-explanatory visual *hurts*. *Fix:* if the visual is clear, delete the
  caption; if you read it aloud, don't print it.
- **Transient information** — spoken/animated info vanishes; dense animations
  overload. *Fix:* leave the final state visible; Morph gently, end static.
- **Worked-example / completion** — for teaching decks, show a solved example
  before asking for problem-solving.

## 3 · Assertion–Evidence research (Michael Alley, Penn State)

The direct experimental backing for "headlines are assertions" (our Law 4):

- Garner & Alley (2013): 110 engineering students, same recorded talk —
  assertion-headline + visual-evidence slides beat topic-bullet slides on
  comprehension (esp. higher-order questions), produced **fewer
  misconceptions**, **lower perceived cognitive load**, and **better recall a
  week later**. Effect on dynamic-process understanding: *d*≈0.60.
- The bullet group saw ~73% more words (334 vs 193) and remembered *no more*
  written info — more text ≠ more learning.
- Garner et al. (2016): *making* assertion-evidence slides also deepens the
  **presenter's own** understanding (significant post-test advantage) — the
  format forces clear thinking upstream.
- Format spec used in studies: sentence headline (assertion) top-left,
  body = visual evidence, minimal callout text, ~28 pt headline, **no bullets**.

(Constructive critique exists — Farkas notes headline benefits are the
best-supported part; the full package needs the visual-evidence discipline too.)

## 4 · Memory effects to design with

- **Picture Superiority Effect** (Paivio, Dual Coding): pictures are
  remembered better than words; classic paired tests show word+picture combos
  recalling roughly **2×** word-only. Caveat from newer work (2025): the edge
  also rides on *distinctiveness* — generic stock photos everyone has seen
  lose it. *Rule:* distinctive, content-bearing visuals > decorative stock.
- **Von Restorff (isolation) effect** (1933): the one item that differs from a
  homogeneous set is remembered. *Rule:* hero stats, one accent bar, one odd
  slide in the rhythm (the STAR moment, `12`) — but only ONE isolate per
  context, against a consistent background.
- **Serial position (primacy/recency)** + **peak–end rule**: openings and
  endings (and the emotional peak) dominate recall. *Rule:* invest design
  effort disproportionately in slide 1, the climax, and the closer; end on the
  ask, not "Thank you/Q&A".
- **Gestalt laws** (via Knaflic): proximity, similarity, enclosure, closure,
  continuity, connection — group with spacing/cards before adding boxes,
  lines and labels. (Our card grids are enclosure+similarity machines.)

## 5 · Chart perception & honest data (Tufte lineage)

- **Data-ink ratio / chartjunk** (Edward Tufte): erase non-data ink; our declutter
  protocol (`05`) is Tufte operationalized. **Small multiples** beat spaghetti.
- Cleveland & McGill's perception ranking: position > length > angle > area —
  why bars beat pies. (Codified in `05`.)
- Tufte's *"The Cognitive Style of PowerPoint"* (2003) remains the canonical
  warning: low-resolution bullet hierarchies can obscure causal thinking
  (famous Columbia-shuttle slide analysis). *Rule:* dense technical evidence
  belongs in handouts/appendix docs; slides carry the assertion + key exhibit.
- **Accessibility numbers** (WCAG 2.x): text ≥ **4.5:1** contrast (3:1 large
  text ≥18 pt / 14 pt bold); meaningful graphics ≥ **3:1**; never encode by
  color alone (add labels/patterns). Colorblind-safe categorical palettes:
  **Okabe-Ito** (Nature Methods/Wong 2011): Orange `E69F00`, Sky Blue `56B4E9`,
  Bluish Green `009E73`, Yellow `F0E442` (fills only — 1.32:1 on white!),
  Blue `0072B2`, Vermilion `D55E00`, Reddish Purple `CC79A7`, Black `000000`.
  Palette-safe ≠ contrast-safe: check both. Avoid red/green pairings (~8% of
  men have red-green CVD); prefer blue/orange oppositions.

## 6 · 🚫 Debunked — never cite these (instant credibility loss)

| Myth | Reality |
|---|---|
| "We remember 10% of what we read, 20% heard… 90% of what we do" (Learning Pyramid / Dale's Cone %) | Dale's 1946 cone had **no numbers**; percentages are unattributed folklore (~1967, Treichler), debunked repeatedly (Subramony et al. 2014; Masters 2013/2020). NTL admits the data doesn't exist. |
| "Communication is 7% words, 38% tone, 55% body language" | Mehrabian (1967) measured *like/dislike judgments under inconsistent cues only*; he himself disclaims general use. Invalid for informational talks. |
| "65% of people are visual learners" / learning-styles matching | No credible evidence that matching "styles" improves outcomes (Pashler et al. consensus). Design for the *content*, not imagined styles. |
| "Attention span is 8 seconds / shorter than a goldfish" | Untraceable stat (misattributed to a Microsoft Canada consumer report); attention is task-dependent. Use segmentation/contrast instead of panic-truncation. |
| "Millennials/Gen Z need everything in 30 seconds" | Generational learning claims lack evidence; clarity & pacing help *everyone* (segmenting principle). |
| Any unsourced "X% more persuasive with visuals" vendor stat | Replace with Mayer/Alley effect sizes above, which are real and citable. |

## Key sources (for citation in decks/handouts)

Mayer (2021) *Multimedia Learning* 3e, Cambridge UP · Mayer & Fiorella (2014),
Cambridge Handbook ch.12 (effect sizes) · Sweller, van Merriënboer & Paas
(1998/2019) Cognitive Architecture & Instructional Design · Chandler &
Sweller (1991) split-attention/redundancy · Garner & Alley (2013) *Int. J.
Engineering Education* · Alley (2013) *The Craft of Scientific Presentations* ·
Paivio (1971/1991) dual coding · von Restorff (1933) · Tufte (1983/2001)
*Visual Display of Quantitative Information*; (2003) *Cognitive Style of
PowerPoint* · Wong (2011) *Nature Methods* 8:441 (color blindness) · WCAG 2.2
(W3C) · Masters (2013) *Medical Teacher* (pyramid debunk) · Lapakko (2007)
(Mehranian legend). Writing/sample slides: writing.engr.psu.edu.
