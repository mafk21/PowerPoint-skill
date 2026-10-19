#!/usr/bin/env python3
"""Deck Decompiler — any .pptx -> generator spec.json (rescue legacy decks).

Reads a foreign deck's real content (headlines, bullets, chart data,
pictures, notes) and maps each slide to the closest generator type, with
the theme guessed by tools/anatomy.py as theme_overrides so the rebuild
keeps the original colors. Pictures are extracted to an assets folder.
Losses are reported honestly (tables flatten to bullets, exotic shapes
become text) — then regenerate, QA, and diff.py against the original to
measure fidelity.

Usage:
  python tools/decompile.py old.pptx [--out spec.json] [--build] 
  # then: python tools/qa_deck.py old.rescued.pptx
  #       python tools/diff.py old.pptx old.rescued.pptx
"""
import argparse
import json
import os
import re
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
import anatomy            # noqa: E402
import draft              # noqa: E402

NUM_RE = re.compile(
    r"^\s*[-+]?\$?[\d٠-٩][\d٠-٩,]*\.?\d*\s*(%|x|k|m|bn|mo|mos|days?|hrs?|yrs?|pts|users)?\s*$",
    re.I)
ASSETS = []
LOSSES = []


def shape_runs(shape):
    """[(text, max_pt)] paragraphs + (max_pt, all_text)."""
    paras, max_pt = [], 0.0
    for para in shape.text_frame.paragraphs:
        txt = "".join(r.text for r in para.runs).strip()
        pts = [r.font.size.pt for r in para.runs if r.font.size]
        pt = max(pts) if pts else 0.0
        max_pt = max(max_pt, pt)
        if txt:
            paras.append((txt, pt))
    return paras, max_pt


def extract_picture(shape, assets_dir, idx):
    try:
        img = shape.image
        ext = img.ext or "png"
        name = f"img{idx}.{ext}"
        os.makedirs(assets_dir, exist_ok=True)
        with open(os.path.join(assets_dir, name), "wb") as f:
            f.write(img.blob)
        ASSETS.append(name)
        return os.path.join(os.path.basename(assets_dir), name)
    except Exception:
        return None


def chart_data(shape):
    try:
        ch = shape.chart
        plot = ch.plots[0]
        cats = [str(c.label) for c in plot.categories]
        ctype = str(ch.chart_type)
        series = []
        for s in ch.series:
            try:
                series.append((str(s.name), [float(v) for v in s.values
                                             if v is not None]))
            except Exception:
                pass
        return ctype, cats, series
    except Exception:
        return None, [], []


def slide_texts(slide):
    out = []
    for sh in anatomy.walk(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip():
            paras, _ = shape_runs(sh)
            out += [t for t, _ in paras]
    return out


def find_furniture(prs):
    """Short texts repeated across >=50% of slides = chrome (brand, footer,
    page furniture), not content. Returns (furniture_set, brand_guess)."""
    from collections import Counter
    n = len(prs.slides)
    counts = Counter()
    for slide in prs.slides:
        for t in set(slide_texts(slide)):
            if len(t) <= 42:
                counts[t.strip()] += 1
    furn = {t for t, c in counts.items() if c >= max(2, n / 2)}
    brand = ""
    for t, c in counts.most_common():
        if t in furn and not re.fullmatch(r"[\d٠-٩]+", t) and len(t) >= 3:
            brand = t
            break
    return furn, brand


def classify_slide(slide, idx, total, assets_dir, portrait, furniture=frozenset()):
    texts = []   # (top, left, max_pt, [paras], shape)
    pics, charts, tables = [], [], []
    for sh in anatomy.walk(slide.shapes):
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            pics.append(sh)
        elif sh.shape_type == MSO_SHAPE_TYPE.CHART:
            charts.append(sh)
        if getattr(sh, "has_table", False) and sh.has_table:
            tables.append(sh)
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            paras, mpt = shape_runs(sh)
            if paras:
                texts.append((sh.top.inches if sh.top is not None else 99,
                              sh.left.inches if sh.left is not None else 0,
                              mpt, paras, sh))
    texts.sort(key=lambda r: (r[0], r[1]))
    all_paras = [p for _, _, _, ps, _ in texts for p in ps]
    clean = [p for p in all_paras if p[0].strip() not in furniture
             and not re.fullmatch(r"[\d٠-٩]{1,3}", p[0].strip())]
    if clean:
        all_paras = clean
    if not all_paras and not charts:
        return None, "empty slide skipped"

    # headline = largest WORDY type in the top half (a giant "$4.2M" is a
    # stat, not a headline); fallback: largest anything.
    top_half = [p for (top, _, _, ps, _) in texts if top < 4.5 for p in ps] \
        or all_paras
    wordy = [p for p in top_half if not NUM_RE.match(p[0])
             and p[0].strip() not in furniture and len(p[0].split()) >= 2]
    headline, hpt = max(wordy or top_half, key=lambda p: p[1])
    rest = [t for t, _ in all_paras if t != headline]
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    sd = {"notes": notes} if notes else {}
    # kicker = small ALL-CAPS line floating above the headline shape
    htop = next((top for (top, _, _, ps, _) in texts
                 if any(t == headline for t, _ in ps)), 99)
    for (top, _, _, ps, _) in texts:
        if top < htop - 0.04:
            for t, pt in ps:
                alpha = [c for c in t if c.isalpha()]
                if pt <= 14.5 and 2 <= len(t) <= 26 and alpha and \
                        sum(c.isupper() for c in alpha) / len(alpha) > 0.72:
                    sd["kicker"] = t.strip()
                    break
        if "kicker" in sd:
            break

    # ---- charts win
    if charts:
        ctype, cats, series = chart_data(charts[0])
        if "DOUGHNUT" in ctype or "PIE" in ctype:
            vals = series[0][1] if series else []
            return dict(type="donut" if not portrait else "bars",
                        title=headline, categories=cats[:4], values=vals[:4],
                        highlight=vals.index(max(vals)) if vals else 0,
                        insight=rest[-1] if rest else "", **sd), f"chart {ctype}"
        if "LINE" in ctype and not portrait:
            return dict(type="line_chart", title=headline, categories=cats,
                        series=[{"name": n, "values": v} for n, v in series],
                        highlight=len(series) - 1 if series else 0,
                        insight=rest[-1] if rest else "", **sd), "line chart"
        name, vals = series[0] if series else ("Value", [])
        return dict(type="bars" if portrait else "chart", title=headline,
                    categories=cats, values=vals, series_name=name,
                    highlight=vals.index(max(vals)) if vals else 0,
                    insight=rest[-1] if rest else "", **sd), f"chart {ctype}"

    # ---- tables flatten to bullets (reported loss)
    if tables:
        rows = []
        for row in tables[0].table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                rows.append(dict(head=cells[0],
                                 body=" · ".join(cells[1:])))
        LOSSES.append(f"slide {idx + 1}: table flattened to bullets")
        return dict(type="content", title=headline, bullets=rows[:4], **sd), \
            "table (flattened)"

    joined = " ".join(rest)
    # ---- title slide
    if idx == 0 and hpt >= 36:
        lines = headline.split("\n")
        return dict(type="hook" if portrait else "title",
                    **({"title_lines": lines} if not portrait else
                       {"title_lines": lines}),
                    subtitle=rest[0] if rest else "",
                    meta=rest[1] if len(rest) > 1 else "", **sd), "title"
    # ---- stat strip: >=2 biggish numbers; label/context = texts below
    # the number IN ITS COLUMN (reading order interleaves columns).
    cells = []  # (num_text, top, left, shape_w)
    for (top, left, mpt, ps, sh) in texts:
        for t, pt in ps:
            if NUM_RE.match(t) and pt >= 26 and t != headline \
                    and t.strip() not in furniture:
                cells.append((t.strip(), top, left,
                              sh.width.inches if sh.width else 2.0))
    if len(cells) >= 2 and not portrait:
        def below(ctop, cleft, cw):
            out = []
            for (top, left, mpt, ps, sh) in texts:
                if top <= ctop + 0.05:
                    continue
                cx = cleft + cw / 2
                sx = left + (sh.width.inches / 2 if sh.width else 1.0)
                if abs(sx - cx) > max(cw, 2.6) * 0.62:
                    continue
                for t, pt in ps:
                    if t.strip() not in furniture and not NUM_RE.match(t) \
                            and t != headline:
                        out.append((top, t))
            out.sort()
            return [t for _, t in out]
        stats = []
        for val, ctop, cleft, cw in sorted(cells, key=lambda c: c[2])[:4]:
            under = below(ctop, cleft, cw)
            stats.append(dict(value=val, label=under[0] if under else "",
                              context=under[1] if len(under) > 1 else ""))
        if sum(1 for s_ in stats if s_["label"]) >= 2:
            return dict(type="stat_strip", title=headline, stats=stats,
                        **sd), f"stat strip ({len(stats)})"
    # ---- hero stat
    big = max(all_paras, key=lambda p: p[1])
    if big[1] >= 54 and NUM_RE.match(big[0]) and \
            (big[1] >= 80 or len(all_paras) <= 5):
        others = [t for t, _ in all_paras if t != big[0]]
        return dict(type="bigstat" if portrait else "hero_stat",
                    stat=big[0].strip(), label=others[0] if others else "",
                    context=others[1] if len(others) > 1 else "", **sd), \
            "big number"
    # ---- quote
    if any(t.strip().startswith(("“", '"', "«")) for t, _ in all_paras) \
            and len(all_paras) <= 4:
        q = max(all_paras, key=lambda p: len(p[0]))[0].strip("“”\"«»")
        attr = next((t.lstrip("—- ").strip() for t, _ in all_paras
                     if t.strip().startswith(("—", "--"))), "")
        return dict(type="quote", quote=q, attribution=attr, **sd), "quote"
    # ---- closing
    low = " ".join(t for t, _ in all_paras).lower()
    if idx == total - 1 and any(k in low for k in
                                ("thank", "شكرا", "contact", "@", "http",
                                 "join us", "raising")):
        return dict(type="cta" if portrait else "closing",
                    **({"ask": headline} if portrait else
                       {"title_lines": [headline]}),
                    subtitle=rest[0] if rest else "", **sd), "closing"
    # ---- agenda-ish
    if any(k in headline.lower() for k in ("agenda", "outline", "contents",
                                           "المحاور", "agenda")) and rest:
        items = [re.sub(r"^\d+[.)]\s*", "", t) for t in rest][:6]
        if portrait:
            return dict(type="steps", title=headline,
                        steps=[dict(head=i, body="") for i in items],
                        **sd), "agenda"
        return dict(type="agenda", title=headline, items=items, **sd), "agenda"

    # ---- default: content / tips from body paragraphs
    bullets = [dict(zip(("head", "body"), draft.split_head_body(t)))
               for t in rest if t]
    img_path = None
    if pics and not portrait:
        img_path = extract_picture(pics[0], assets_dir, len(ASSETS) + 1)
    if portrait:
        items = [dict(icon=draft.icon_for(b["head"]), **b)
                 for b in bullets[:4]]
        return dict(type="tips", title=headline, items=items, **sd), \
            "text → tips"
    out = dict(type="content", title=headline, bullets=bullets[:4], **sd)
    if img_path:
        out["image"] = img_path
    if not bullets and rest:
        out["bullets"] = [dict(head=headline, body=joined[:200])]
    return out, "text → content"


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck")
    ap.add_argument("--out", help="spec path (default <deck>.spec.json)")
    ap.add_argument("--build", action="store_true",
                    help="also regenerate <deck>.rescued.pptx")
    ap.add_argument("--theme-name", default="",
                    help="save guessed theme into themes.json under NAME "
                         "instead of theme_overrides")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.deck):
        print(f"no such file: {a.deck}")
        return 2
    prs = Presentation(a.deck)
    W, H = prs.slide_width.inches, prs.slide_height.inches
    portrait_canvas = H > W
    out = a.out or os.path.splitext(a.deck)[0] + ".spec.json"
    assets_dir = os.path.join(os.path.dirname(os.path.abspath(out)), "assets")

    furniture, brand_guess = find_furniture(prs)
    slides, report = [], []
    for i, slide in enumerate(prs.slides):
        sd, why = classify_slide(slide, i, len(prs.slides), assets_dir,
                                 portrait_canvas, furniture)
        report.append((i + 1, sd.get("type", "—") if sd else "—", why))
        if sd:
            slides.append(sd)

    spec = {"title": f"Rescued: {os.path.basename(a.deck)}", "slides": slides}
    if brand_guess:
        spec["brand"] = brand_guess
    if portrait_canvas:
        spec["format"] = "reel" if abs(W / H - 9 / 16) < 0.03 else "carousel"
    ana = anatomy.analyze(a.deck)
    theme_pack, _ = anatomy.guess_theme(ana)
    if theme_pack:
        theme, _rep = theme_pack
        if a.theme_name:
            from palette_lab import save_theme
            save_theme(a.theme_name, theme)
            spec["theme"] = a.theme_name
        else:
            spec["theme"] = "teal-orbit"
            spec["theme_overrides"] = {k: theme[k] for k in
                                       ("bg", "card", "ink", "muted", "accent",
                                        "accent2", "graybar", "dark")}
    # RTL detection: any rtl paragraph attr in XML
    from zipfile import ZipFile
    z = ZipFile(a.deck)
    xml0 = z.read("ppt/slides/slide1.xml").decode("utf-8", "ignore")
    if 'rtl="1"' in xml0:
        spec["rtl"] = True

    with open(out, "w", encoding="utf-8") as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)
    print(f"decompiled {len(slides)}/{len(prs.slides)} slides -> {out}"
          f"{' (portrait)' if portrait_canvas else ''}")
    for n, t, why in report:
        print(f"  {n:2d}. {t:12s} ({why})")
    if ASSETS:
        print(f"extracted pictures: {', '.join(ASSETS)} (in assets/)")
    for l in LOSSES:
        print(f"  LOSS: {l}")
    if a.build:
        deck_out = os.path.splitext(a.deck)[0] + ".rescued.pptx"
        if portrait_canvas:
            import portrait as P
            P.build_portrait(spec, deck_out)
        else:
            import generate_pptx as G
            G.build(spec, deck_out)
        print(f"regenerated {deck_out}\nnext: python tools/qa_deck.py "
              f"{deck_out} && python tools/diff.py {a.deck} {deck_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
