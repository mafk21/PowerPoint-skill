# Connected by Technology · Conversation — interactive edition

Rebuilt from **`source/Demo lesson Oct 21.pptx`** (23 slides, static) into a **24-slide, click-by-click
interactive lesson** for the classroom projector.

**Deliverable:** `Demo lesson Oct 21 - Interactive.pptx` (24 slides, 16:9, no macros, 2.9 MB)

---

## 1. What the original deck was

A Grade-7 English lesson (Ministry of Education / Madrasati / NCVT branding) built around one audio
dialogue: **Grandma texts her granddaughter Cara for the first time** and learns `DYWT`, `Gr8`, `B ovr n
10 mins`. The 23 slides walked the class through warm-up → prediction → listening → reading → Real Talk
phrases → worksheets → speaking tasks → quiz.

Its material was good; its delivery was flat. Every answer, tick and correction was **already printed on
the slide**, so nothing could be asked before it was answered, and 23 static slides carried no sense of
progress.

## 2. What changed — and what did not

**Nothing in the lesson material was reworded, shortened or deleted.** Every transcript line, question,
worksheet item, Real Talk gloss, true/false statement, quiz option and audio file is the source text,
verbatim. Two purely typographic normalisations are disclosed at the end.

What changed is *everything around* the material:

| Layer | Before | Now |
|---|---|---|
| Design system | ad-hoc fonts, drop shadows, clip-art | one palette (navy / ink / accent blue / cream), two fonts (Trebuchet MS + Calibri), 12-pt minimum, 14 fills total |
| Layout | centred blocks of text | assertion headlines, one idea per slide, left-aligned reading column, right-hand support rail |
| Motion | none | a shared Morph anchor pair on all 24 slides + a per-slide motion story |
| Chrome | none | progress hairline + `NN / 24` + branding lockup + footer signature |
| Answers | printed | **invisible until a student commits** — reveal by clicking the card |
| Feedback | none | green tick / red cross badges, "Not quite — listen again", explain-why bands |
| Navigation | linear | clickable *Start the lesson* → deck, *Back to the start* → title (loops for the next class) |
| Teacher support | none | 24 speaker-note scripts with timings, answers and differentiation |

## 3. How the interactivity works (no macros, no add-ins)

Everything is native PowerPoint **trigger animations** (`interactiveSeq` on click → entrance effects).
Nothing is a hyperlink to a hidden "answer slide"; the answer appears *on the same slide*, so the class
never loses the question. 106 triggers across 19 slides.

* Tap the **card / chip / option** → its reveal fades in where the eye already is.
* Every click is repeatable (the effect holds), so you can re-ask the same question in the next class.
* "Show all answers" / "Show the order" / "Show answer" buttons exist for the moment you want to move fast.
* In **edit view** you see every answer: present in Slide Show (or F5) to see the deck the class sees.

### Slide-by-slide interaction map

| # | Slide | Tap this | To reveal |
|---|---|---|---|
| 1 | Title | *Start the lesson* | jumps to slide 2 |
| 2 | Lesson Objectives | a row | green tick as the class covers that aim |
| 3 | Warm-up YES / NO | YES · NO | the vote tick |
| 4 | WHAT? | *Reveal the words* | `texting`, `abbreviation` |
| 5 | e.g. | a chip | `PLZ = please`, `Gr8 = great`, `B4 = before` |
| 6 | Prediction | *Check* | "Cara is texting Maria." |
| 7 | Listen to the audio | any of 6 question cards | its answer line |
| 8 | Listen 'N' read | — (audio + script) | — |
| 9 | About the Conversation | a task card | tick once the pair has answered |
| 10 | Role-Play | *Grandma* card / *Cara* card | that role's lines only — half the class each |
| 11 | Language Builder | any of 8 word cards | its meaning (+ permanent Real Talk rail) |
| 12 | Find the phrases | a phrase chip | the line where it appears lights up |
| 13 | Real Talk (phrases) | a sentence | what the blue phrase means there |
| 14 | Practice Time | a frame | tick |
| 15 | True or false? | TRUE / FALSE | evidence line **or** "Not quite — listen again" |
| 16 | Clarify | a phrase | green = asks for clarification, red = doesn't |
| 17 | Confirm | a phrase | only "I get it." is green |
| 18 | Check | a phrase | "Does that make sense?" is green |
| 19 | Talk Time | a task card | phrase hints inside the card |
| 20 | Choose the phrase | an option | green ring + tick, red cross on the others |
| 21 | Choose the phrase | an option | same |
| 22 | Choose the phrase | an option | same |
| 23 | Put it in order | a message | its number appears |
| 24 | Thank you | *Back to the start* | jumps to slide 1 |

## 4. Fidelity

* **Audio:** the original `media1.mp3` is embedded and playable on slides 6, 7, 8 and 12 — same file, no re-encode.
* **Text:** all 87 source text lines were diffed against the new deck. 82 match character-for-character
  after whitespace/quote normalisation; the remaining 5 are the `PLZ = please` style vocabulary rows,
  which are split across three shapes (word, `=`, gloss) on one card but contain exactly the same
  characters.
* **Disclosed normalisations:** (1) the acronym table's `=` is drawn as a separate glyph, (2) the source's
  `1 – ` / `1- ` numbering prefixes are carried by numeral chips on slide 2 and re-joined to the text on
  slide 15. No word was changed.
* **No source art was redrawn**; the original illustrations, emoji and branding strip are used as shipped.

## 5. QA

`tools/qa_deck.py` on the delivered file:

```
PASS  Canvas 16:9 13.33x7.50
PASS  Morph on slides 2..N, none on #1
PASS  Travelling !! anchors on every slide
FAIL  Words/slide (max 217)                       <- see note
PASS  No text < 12pt (0 tiny runs)
WARN  Text 12-16pt (296 runs, body/rail text)
PASS  Font families <= 2
WARN  Headlines look like assertions (4 source titles kept verbatim)
WARN  Color discipline (14 distinct fills)
PASS  No leaked placeholder text
PASS  Speaker notes 24/24
PASS  Text overflow sentinel
PASS  Text contrast vs background (WCAG AA)
```

**The one FAIL is deliberate.** Slides 8, 10, 12 and 13 carry the *verbatim* transcript and the *verbatim*
Real Talk table. Cutting words to satisfy a density heuristic would break the one rule of this job, so the
script slides stay dense and every other slide is held to the word budget. They are reference slides, read
with the audio, not slides the class reads cold.

## 6. Rebuilding

```bash
cd examples/lesson-interactive
../../.venv/bin/python build.py        # writes the .pptx
```

* `deckkit.py` — design tokens, primitives (`rect`, `text`, `badge`, `line`…), the `Deck` class
  (Morph, notes, triggers, hyperlinks, audio-rel normalisation).
* `build.py` — the 24 slides; **all source wording lives in the constants at the top**
  (`OBJECTIVES`, `TRANSCRIPT`, `REAL_TALK`, `FRAMES`, `TF`, `QUIZ`, `REORDER`, `WORKSHEET`, `TASK`),
  so wording can only change in one place.
* `tools/render_pptx.py` — PPTX→PNG renderer used to proof every slide (approximate fonts: DejaVu stands
  in for Calibri/Trebuchet).
* `assets/` — original media plus trimmed copies in `assets/t/` used by the layout.

### Audio

The play buttons on slides 6, 7, 8 and 12 embed the original `media1.mp3` and are wired the
same way PowerPoint itself writes them: an `<a:audioFile>` picture whose click action is
`ppaction://media`, an `audio/mpeg` content-type registration, and a `<p:audio><p:cMediaNode>`
cue in the slide timing (the node python-pptx omits, and without which a click just advances
the slide). Each play button is also raised to the top of the z-order so no text box can
swallow the click.

Present in **PowerPoint 2019 / Microsoft 365** (Morph + triggers + audio). Google Slides and Keynote will show
the slides but flatten the trigger reveals; export to video if you need a self-playing copy.
