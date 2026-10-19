#!/usr/bin/env python3
"""
Deck QA analyzer — the automated form of reference/06-qa-checklist.

    python tools/qa_deck.py my-deck.pptx [--strict]

Checks (heuristics are labeled as such — a human still owns taste):
  * 16:9 canvas
  * Morph transition present on slides 2..N (real XML inspection)
  * !! forced-morph anchors shared across slides (travelling-element health)
  * words per slide (PlacePlate budget: ~30 pitch / ~60 general ceiling)
  * minimum font sizes (nothing <12pt; <16pt flagged except footer furniture)
  * font-family count (<=2 target)
  * headline shape (heuristic: <=3-word largest text ~ topic label, not assertion)
  * distinct fill colors (3-role discipline; >10 distinct hues flagged)
  * speaker-notes coverage
Exit code 0 = pass (warnings allowed), 1 = failures, 2 = file problem.
Requires python-pptx.
"""
import argparse
import re
import sys
import zipfile
from collections import Counter

from pptx import Presentation
from pptx.enum.dml import MSO_FILL

PASS, WARN, FAIL = "PASS", "WARN", "FAIL"


def slide_xmls(path):
    z = zipfile.ZipFile(path)
    names = sorted((n for n in z.namelist()
                    if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                   key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    return [z.read(n).decode("utf-8", "ignore") for n in names]


def shape_text(shape):
    if not shape.has_text_frame:
        return "", []
    words, runs = [], []
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip():
                words.append(r.text)
                runs.append(r)
    return " ".join(words), runs


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--strict", action="store_true",
                    help="treat warnings as failures")
    ap.add_argument("--word-ceiling", type=int, default=80)
    a = ap.parse_args(argv)

    try:
        prs = Presentation(a.pptx)
        xmls = slide_xmls(a.pptx)
    except Exception as e:
        print(f"FAIL  cannot open: {e}")
        return 2

    results = []  # (status, label, detail)

    w_in, h_in = prs.slide_width.inches, prs.slide_height.inches
    ratio_ok = abs(w_in / h_in - 16 / 9) < 0.02
    results.append((PASS if ratio_ok else FAIL, "Canvas 16:9",
                    f"{w_in:.2f}x{h_in:.2f} in"))

    n = len(xmls)
    morph = ["morph" in x for x in xmls]
    morph_ok = (n <= 1) or (not morph[0] and all(morph[1:]))
    results.append((PASS if morph_ok else FAIL, "Morph on slides 2..N, none on #1",
                    f"morph map: {''.join('M' if m else '.' for m in morph)}"))

    bang_per_slide = [set(re.findall(r'name="(!![^"]+)"', x)) for x in xmls]
    shared = set.intersection(*bang_per_slide) if bang_per_slide else set()
    results.append((PASS if len(shared) >= 2 else WARN,
                    "Travelling !! anchors on every slide",
                    f"{len(shared)} shared: {', '.join(sorted(shared)[:6]) or 'none'}"))

    fonts, tiny, small = Counter(), [], []
    word_counts, label_heads, notes_have = [], [], 0
    fill_colors = Counter()
    for i, slide in enumerate(prs.slides, 1):
        total_words, biggest = 0, (0, "")
        for shape in slide.shapes:
            text, runs = shape_text(shape)
            total_words += len(text.split())
            for r in runs:
                if r.font.name:
                    fonts[r.font.name] += 1
                if r.font.size:
                    pt = r.font.size.pt
                    if pt > biggest[0]:
                        biggest = (pt, text)
                    if pt < 12:
                        tiny.append((i, pt, r.text[:24]))
                    elif pt < 16:
                        small.append((i, pt))
            try:
                if shape.fill and shape.fill.type == MSO_FILL.SOLID:
                    fill_colors[str(shape.fill.fore_color.rgb)] += 1
            except Exception:
                pass
        word_counts.append(total_words)
        # skip hero numbers (38%, 02, $4.2M) — biggest text containing a digit
        # is a stat/section marker, not a lazy label headline
        if (biggest[1] and len(biggest[1].split()) <= 3 and i not in (1,)
                and not any(c.isdigit() for c in biggest[1])):
            label_heads.append((i, biggest[1][:40]))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            notes_have += 1

    over = [(i + 1, c) for i, c in enumerate(word_counts) if c > a.word_ceiling]
    dense = [(i + 1, c) for i, c in enumerate(word_counts) if 60 < c <= a.word_ceiling]
    wstatus = FAIL if over else (WARN if dense else PASS)
    results.append((wstatus,
                    f"Words/slide (fail>{a.word_ceiling}, dense>60)",
                    f"max={max(word_counts or [0])} avg={sum(word_counts)/max(len(word_counts),1):.0f}"
                    + (f" over: {over[:5]}" if over else (f" dense exhibits: {dense[:5]}" if dense else ""))))
    results.append((FAIL if tiny else PASS, "No text < 12pt",
                    f"{len(tiny)} tiny runs" + (f" e.g. slide {tiny[0][0]} {tiny[0][1]:.0f}pt '{tiny[0][2]}'" if tiny else "")))
    results.append((WARN if small else PASS, "Text 12-16pt (footer/chart zone only)",
                    f"{len(small)} runs"))
    results.append((PASS if len(fonts) <= 2 else WARN, "Font families <= 2",
                    ", ".join(f"{k}×{v}" for k, v in fonts.most_common(4)) or "none set"))
    results.append((WARN if label_heads else PASS,
                    "Headlines look like assertions (heuristic)",
                    f"possible topic-labels: {label_heads[:4]}" if label_heads else "all >3 words"))
    results.append((PASS if len(fill_colors) <= 10 else WARN,
                    "Color discipline (distinct fills)",
                    f"{len(fill_colors)} distinct fill colors"))
    results.append((PASS if notes_have >= max(n - 1, 0) else WARN,
                    "Speaker notes coverage", f"{notes_have}/{n} slides"))

    print(f"QA: {a.pptx} — {n} slides")
    worst = PASS
    rank = {PASS: 0, WARN: 1, FAIL: 2}
    for status, label, detail in results:
        worst = status if rank[status] > rank[worst] else worst
        print(f"  {status:4s}  {label} — {detail}")
    fails = sum(1 for s, _, _ in results if s == FAIL)
    warns = sum(1 for s, _, _ in results if s == WARN)
    print(f"Result: {fails} fail, {warns} warn -> {'GREEN' if fails == 0 and (warns == 0 or not a.strict) else 'NEEDS WORK'}")
    return 1 if (fails or (a.strict and warns)) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
