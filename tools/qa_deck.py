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
  * text overflow sentinel (estimated rendered height vs box — no PPT needed)
  * text contrast vs effective background (WCAG 4.5:1 / 3:1 large)
  * canvas: 16:9 standard; 9:16 / 4:5 / 1:1 / 4:3 accepted (social formats)
Exit code 0 = pass (warnings allowed), 1 = failures, 2 = file problem.
Requires python-pptx.
"""
import argparse
import math
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


def _lum(hexs):
    def f(c):
        c = c / 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hexs[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast(a, b):
    la, lb = _lum(a), _lum(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _solid(shape):
    try:
        if shape.fill and shape.fill.type == MSO_FILL.SOLID:
            return str(shape.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def overflow_ratio(shape):
    """Estimated text height / available height. >1 means likely spill.
    Model: avg glyph advance ~0.50em (0.58em all-caps), line height from
    paragraph spacing (default 1.19). Calibrated conservatively — the
    sentinel flags suspects; the renderer (tools/render.py) shows them."""
    if not (shape.width and shape.height) or shape.height.inches < 0.3:
        return None
    tf = shape.text_frame
    if tf.word_wrap is False:
        return None
    # NB: Length(0) is falsy — must test `is not None`, or zeroed margins
    # silently become PPT defaults and every snug box "overflows".
    ml = tf.margin_left.inches if tf.margin_left is not None else 0.1
    mr = tf.margin_right.inches if tf.margin_right is not None else 0.1
    mt = tf.margin_top.inches if tf.margin_top is not None else 0.05
    mb = tf.margin_bottom.inches if tf.margin_bottom is not None else 0.05
    inner_w = max(shape.width.inches - ml - mr, 0.2)
    inner_h = max(shape.height.inches - mt - mb, 0.1)
    est = 0.0
    for para in tf.paragraphs:
        text = "".join(r.text for r in para.runs)
        if not text.strip():
            continue
        pts = [r.font.size.pt for r in para.runs if r.font.size]
        pt = max(pts) if pts else 18.0
        alpha = [c for c in text if c.isalpha()]
        cw = 0.58 if alpha and sum(c.isupper() for c in alpha) / len(alpha) > 0.7 else 0.50
        cpl = max(inner_w / (cw * pt / 72), 1.0)
        lines = sum(math.ceil(len(seg) / cpl)
                    for seg in re.split(r"[\n\v]", text)) or 1
        raw_ls = para.line_spacing
        if isinstance(raw_ls, float):
            est += lines * pt * raw_ls / 72
        elif raw_ls is not None:  # exact Pt line height
            est += lines * raw_ls.inches
        else:
            est += lines * pt * 1.19 / 72
        if para.space_after:
            est += para.space_after.inches
        if para.space_before:
            est += para.space_before.inches
    return (est / inner_h) if est else None


def effective_bg(shapes, j, slide_bg, slide_fill_fallback):
    """Background color under shapes[j]: its own fill (buttons/chips),
    else the nearest filled shape beneath containing its center,
    else slide background. Returns (hex, confident)."""
    sh = shapes[j]
    own = _solid(sh)
    if own:
        return own, True
    if not (sh.width and sh.height and sh.left is not None):
        return (slide_bg or slide_fill_fallback), bool(slide_bg)
    cx = sh.left.inches + sh.width.inches / 2
    cy = sh.top.inches + sh.height.inches / 2
    for k in range(j - 1, -1, -1):
        o = shapes[k]
        col = _solid(o)
        if not col or not (o.width and o.height and o.left is not None):
            continue
        if (o.left.inches - 0.05 <= cx <= o.left.inches + o.width.inches + 0.05
                and o.top.inches - 0.05 <= cy <= o.top.inches + o.height.inches + 0.05):
            return col, True
    return (slide_bg or slide_fill_fallback), bool(slide_bg)


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
    RATIOS = {"16:9": 16 / 9, "9:16": 9 / 16, "4:5": 4 / 5, "1:1": 1.0,
              "4:3": 4 / 3}
    rname = next((k for k, v in RATIOS.items() if abs(w_in / h_in - v) < 0.02),
                 None)
    results.append((PASS if rname else FAIL,
                    f"Canvas {rname or 'non-standard'}",
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
    overflow, tight = [], []
    PLACEHOLDERS = {"Presentation Title", "Your hook here"}
    placeholder_hits = []
    contrast_fails, contrast_warns = [], []
    for i, slide in enumerate(prs.slides, 1):
        total_words, biggest = 0, (0, "")
        shapes = list(slide.shapes)
        slide_bg = None
        try:
            if slide.background.fill.type == MSO_FILL.SOLID:
                slide_bg = str(slide.background.fill.fore_color.rgb)
        except Exception:
            pass
        slide_fills = Counter()
        for _sh in shapes:
            _c = _solid(_sh)
            if _c and _sh.width and _sh.height:
                slide_fills[_c] += _sh.width.inches * _sh.height.inches
        slide_fallback = (slide_fills.most_common(1)[0][0]
                          if slide_fills else "FFFFFF")
        for shape in shapes:
            text, runs = shape_text(shape)
            total_words += len(text.split())
            if text.strip() in PLACEHOLDERS:
                placeholder_hits.append((i, text.strip()))
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
        # --- overflow sentinel + contrast auditor (need z-order) ---
        for j, shape in enumerate(shapes):
            if not shape.has_text_frame:
                continue
            stext = shape.text_frame.text.strip()
            if not stext:
                continue
            ratio = overflow_ratio(shape)
            if ratio and ratio > 1.12:
                overflow.append((i, ratio, stext[:26]))
            elif ratio and ratio > 1.0:
                tight.append((i, ratio))
            bg, confident = effective_bg(shapes, j, slide_bg, slide_fallback)
            if not bg:
                continue
            for para in shape.text_frame.paragraphs:
                for r in para.runs:
                    t = r.text.strip()
                    if len(t) <= 2:  # glyphs/dingbats: graphics, not text
                        continue
                    try:
                        if not (r.font.color and r.font.color.type is not None):
                            continue
                        fg = str(r.font.color.rgb)
                    except Exception:
                        continue
                    pt = r.font.size.pt if r.font.size else 18.0
                    floor = 3.0 if pt >= 24 or (pt >= 18 and r.font.bold) else 4.5
                    cr = _contrast(fg, bg)
                    if cr < floor:
                        row = (i, round(cr, 1), floor, fg, bg, t[:22])
                        (contrast_fails if confident else contrast_warns).append(row)
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
    AR_FONTS = ("cairo", "tajawal", "amiri", "almarai", "arabic", "noto")
    ar_fams = [f for f in fonts if any(k in f.lower() for k in AR_FONTS)]
    la_fams = [f for f in fonts if f not in ar_fams]
    fonts_ok = len(fonts) <= 2 or (len(ar_fams) <= 2 and len(la_fams) <= 2)
    results.append((PASS if fonts_ok else WARN,
                    "Font families <= 2 (per script)",
                    ", ".join(f"{k}×{v}" for k, v in fonts.most_common(5))
                    or "none set"))
    results.append((WARN if label_heads else PASS,
                    "Headlines look like assertions (heuristic)",
                    f"possible topic-labels: {label_heads[:4]}" if label_heads else "all >3 words"))
    results.append((PASS if len(fill_colors) <= 10 else WARN,
                    "Color discipline (distinct fills)",
                    f"{len(fill_colors)} distinct fill colors"))
    results.append((FAIL if placeholder_hits else PASS,
                    "No leaked placeholder/default text",
                    (f"{len(placeholder_hits)} e.g. slide {placeholder_hits[0][0]} "
                     f"'{placeholder_hits[0][1]}' (spec field typo?)"
                     if placeholder_hits else "none")))
    results.append((PASS if notes_have >= max(n - 1, 0) else WARN,
                    "Speaker notes coverage", f"{notes_have}/{n} slides"))
    ostatus = FAIL if overflow else (WARN if tight else PASS)
    results.append((ostatus, "Text overflow sentinel (est. height vs box)",
                    (f"{len(overflow)} likely spills e.g. slide {overflow[0][0]} "
                     f"{overflow[0][1]:.2f}x '{overflow[0][2]}'"
                     if overflow else
                     (f"{len(tight)} tight boxes (1.0-1.12x)" if tight
                      else "no shape overflows"))))
    cstatus = FAIL if contrast_fails else (WARN if contrast_warns else PASS)
    cf = contrast_fails or contrast_warns
    results.append((cstatus, "Text contrast vs background (WCAG)",
                    (f"{len(contrast_fails)} fail / {len(contrast_warns)} "
                     f"uncertain-bg e.g. slide {cf[0][0]} #{cf[0][3]} on "
                     f"#{cf[0][4]} = {cf[0][1]}:1 (floor {cf[0][2]}) '{cf[0][5]}'"
                     if cf else "all measured runs clear their floor")))

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
