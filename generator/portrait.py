#!/usr/bin/env python3
"""Portrait engine — Reel (9:16) and Carousel (4:5) decks from JSON specs.

Sibling of generate_pptx.py: same themes, same Morph XML, same helpers,
vertical canvases and phone-first type sizes. Reel layouts keep key
content inside the platform safe band (~16–76% of height; reference/13 §6);
carousels get progress dots, reels get a progress bar + counter.

Spec: {"format": "reel"|"carousel", "theme": ..., "brand": ..., "slides": [
  {"type": "hook",     "badge": "...", "title_lines": [...], "subtitle": ...},
  {"type": "bigstat",  "stat": "87%", "label": ..., "context": ...},
  {"type": "tips",     "kicker": ..., "title": ..., "items": [{icon,head,body}]},
  {"type": "bars",     "kicker": ..., "title": ..., "categories": [...],
                       "values": [...], "highlight": 3, "insight": ...},
  {"type": "steps",    "kicker": ..., "title": ..., "steps": [{head,body}]},
  {"type": "quote",    "quote": ..., "attribution": ...},
  {"type": "compare",  "title": ..., "before": {"head","rows"},
                       "after": {"head","rows"}},
  {"type": "cta",      "ask": ..., "handle": "@you", "subtitle": ...},
]}
Any slide may carry "seconds" (reel pacing; default 4s or notes estimate)
and "notes" (the voiceover line — tools/social.py turns it into a shotlist).

Usage:  python generator/portrait.py spec.json out.pptx
"""
import json
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_pptx as G  # noqa: E402

FORMATS = {"reel": (7.5, 13.333), "carousel": (10.667, 13.333)}


# ------------------------------------------------------------------- chrome
def _chrome(slide, t, idx, total, brand, W, H, fmt):
    # travelling orb (top-right; builders may add a second, bigger one)
    G.add_shape(slide, MSO_SHAPE.OVAL, W - 1.15, H - 1.05, 1.7, 1.7,
                fill=t["accent"], name="!!p-orb")
    G.add_shape(slide, MSO_SHAPE.OVAL, W - 0.92, H - 0.82, 1.24, 1.24,
                fill=t["bg"], name="!!p-orb-hole")
    if brand:
        G.add_text(slide, 0.55, 0.62, W * 0.6, 0.4, brand.upper(), size=13,
                   bold=True, font=G.HEADING_FONT, color=t["muted"],
                   name="!!brand")
    if fmt == "reel":
        G.add_shape(slide, MSO_SHAPE.RECTANGLE, 0.55, 0.38, W - 1.1, 0.045,
                    fill=t["graybar"], name="!!p-track")
        G.add_shape(slide, MSO_SHAPE.RECTANGLE, 0.55, 0.38,
                    (W - 1.1) * (idx + 1) / max(total, 1), 0.045,
                    fill=t["accent"], name="!!p-progress")
        G.add_text(slide, W - 1.7, 0.62, 1.15, 0.4, f"{idx + 1}/{total}",
                   size=13, bold=True, font=G.HEADING_FONT, color=t["muted"],
                   align=PP_ALIGN.RIGHT, name="!!pageno")
    else:  # carousel: progress dots
        dots_w = total * 0.34
        x0 = (W - dots_w) / 2
        for k in range(total):
            G.add_shape(slide, MSO_SHAPE.OVAL, x0 + k * 0.34, H - 0.85,
                        0.17, 0.17,
                        fill=t["accent"] if k == idx else t["graybar"],
                        name=f"!!p-dot-{k}")
        G.add_text(slide, W - 1.7, 0.62, 1.15, 0.4, f"{idx + 1}/{total}",
                   size=13, bold=True, font=G.HEADING_FONT, color=t["muted"],
                   align=PP_ALIGN.RIGHT, name="!!pageno")


def _title_block(slide, t, d, W, y=2.3, size=38):
    if d.get("kicker"):
        G.add_text(slide, 0.55, y, W - 1.1, 0.45, d["kicker"].upper(),
                   size=15, bold=True, font=G.HEADING_FONT,
                   color=t["accent"] if not t["dark"] else t["accent2"])
        y += 0.55
    if d.get("title"):
        G.add_text(slide, 0.55, y, W - 1.1, 1.9, d["title"], size=size,
                   bold=True, font=G.HEADING_FONT, color=t["ink"],
                   line_spacing=1.08)
    return y


# ----------------------------------------------------------------- builders
def p_hook(slide, t, d, idx, total, brand, W, H, fmt):
    G.add_shape(slide, MSO_SHAPE.OVAL, W * 0.42, H * 0.52, W * 0.85, W * 0.85,
                fill=t["accent"], name="!!p-hero-orb")
    if d.get("badge"):
        bw = 0.42 + 0.115 * len(d["badge"])
        G.add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, H * 0.17,
                    bw, 0.52, fill=t["accent"], adj=0.5, name="!!p-badge")
        G.add_text(slide, 0.55, H * 0.17 + 0.09, bw, 0.36, d["badge"].upper(),
                   size=14, bold=True, font=G.HEADING_FONT,
                   color=t["bg"] if t["dark"] else "FFFFFF",
                   align=PP_ALIGN.CENTER)
    lines = d.get("title_lines") or [d.get("title", "Your hook here")]
    size = 56 if W < 9 else 64
    G.add_text(slide, 0.55, H * 0.215, W - 1.1, 3.6, lines, size=size,
               bold=True, font=G.HEADING_FONT, color=t["ink"],
               line_spacing=1.05, name="!!p-hook")
    if d.get("subtitle"):
        G.add_text(slide, 0.55, H * 0.215 + 0.42 + 0.86 * len(lines), W - 1.1,
                   1.95, d["subtitle"], size=19 if W < 9 else 21,
                   color=t["muted"], line_spacing=1.25)
    hint = d.get("hint") or ("swipe →" if fmt == "carousel" else "▶ watch")
    G.add_text(slide, 0.55, H * 0.70, 3.4, 0.5, hint, size=17, bold=True,
               font=G.HEADING_FONT, color=t["ink"])


def p_bigstat(slide, t, d, idx, total, brand, W, H, fmt):
    _title_block(slide, t, d, W, y=H * 0.17, size=30)
    stat = str(d.get("stat", "3x"))
    G.add_text(slide, 0.4, H * 0.36, W - 0.8, 2.2, stat,
               size=120 if len(stat) <= 5 else 92, bold=True,
               font=G.HEADING_FONT, color=t["accent"],
               align=PP_ALIGN.CENTER, name="!!p-stat")
    G.add_text(slide, 0.7, H * 0.55, W - 1.4, 0.95, d.get("label", ""),
               size=26, bold=True, font=G.HEADING_FONT, color=t["ink"],
               align=PP_ALIGN.CENTER, line_spacing=1.15)
    if d.get("context"):
        G.add_text(slide, 0.9, H * 0.635, W - 1.8, 1.2, d["context"],
                   size=18, color=t["muted"], align=PP_ALIGN.CENTER,
                   line_spacing=1.3)


def p_tips(slide, t, d, idx, total, brand, W, H, fmt):
    _title_block(slide, t, d, W, y=H * 0.155, size=36)
    items = (d.get("items") or d.get("bullets") or [])[:4]
    n = max(len(items), 1)
    y = H * (0.315 if fmt == "reel" else 0.345)
    end = H * (0.86 if fmt == "reel" else 0.88)
    gap = 0.3
    ch = max(1.35, min(1.78, (end - y - gap * (n - 1)) / n))
    for i, it in enumerate(items):
        G.add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, y, W - 1.1, ch,
                    fill=t["card"], adj=0.09, name=f"!!p-card-{i}")
        G.add_shape(slide, MSO_SHAPE.OVAL, 0.85, y + 0.32, 0.72, 0.72,
                    fill=t["accent"])
        G.add_text(slide, 0.85, y + 0.44, 0.72, 0.5, G._icon_glyph(it),
                   size=21, bold=True,
                   color=t["bg"] if t["dark"] else "FFFFFF",
                   align=PP_ALIGN.CENTER)
        G.add_text(slide, 1.85, y + 0.2, W - 2.55, 0.72,
                   it.get("head") or it.get("title", ""),
                   size=19 if W < 9 else 21, bold=True,
                   font=G.HEADING_FONT, color=t["ink"], line_spacing=1.1)
        G.add_text(slide, 1.85, y + 0.78, W - 2.55, ch - 0.86,
                   it.get("body", ""), size=15.5 if W < 9 else 16.5,
                   color=t["muted"], line_spacing=1.22)
        y += ch + gap


def p_bars(slide, t, d, idx, total, brand, W, H, fmt):
    _title_block(slide, t, d, W, y=H * 0.155, size=36)
    cats = d.get("categories", [])[:6]
    vals = [float(v) for v in d.get("values", [])][:6]
    hi = int(d.get("highlight", 0))
    vmax = max(vals) if vals else 1.0
    y = H * 0.36
    fmt_v = d.get("format", "{:g}")
    for i, (c, v) in enumerate(zip(cats, vals)):
        G.add_text(slide, 0.55, y, W - 2.6, 0.45, str(c), size=18, bold=True,
                   font=G.HEADING_FONT, color=t["ink"])
        G.add_text(slide, W - 2.15, y, 1.6, 0.45, fmt_v.format(v), size=19,
                   bold=True, font=G.HEADING_FONT,
                   color=t["accent"] if i == hi else t["muted"],
                   align=PP_ALIGN.RIGHT)
        bw = (W - 1.1) * (v / vmax) if vmax else 0
        G.add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, y + 0.5,
                    max(bw, 0.25), 0.44,
                    fill=t["accent"] if i == hi else t["graybar"],
                    adj=0.5, name=f"!!p-bar-{i}")
        y += 1.28
    if d.get("insight"):
        G.add_text(slide, 0.55, y + 0.15, W - 1.1, 1.5, d["insight"],
                   size=18, color=t["ink"], line_spacing=1.3)


def p_steps(slide, t, d, idx, total, brand, W, H, fmt):
    _title_block(slide, t, d, W, y=H * 0.155, size=36)
    steps = d.get("steps", [])[:5]
    y = H * 0.35
    rh = max(1.5, min(1.95, (H * 0.86 - y) / max(len(steps), 1)))
    if len(steps) > 1:
        G.add_shape(slide, MSO_SHAPE.RECTANGLE, 1.03, y + 0.5, 0.05,
                    rh * (len(steps) - 1), fill=t["graybar"], name="!!p-line")
    for i, st in enumerate(steps):
        yy = y + i * rh
        G.add_shape(slide, MSO_SHAPE.OVAL, 0.68, yy, 0.75, 0.75,
                    fill=t["accent"] if i == 0 else t["card"],
                    name=f"!!p-step-{i}")
        G.add_text(slide, 0.68, yy + 0.16, 0.75, 0.45, str(i + 1), size=20,
                   bold=True, font=G.HEADING_FONT,
                   color=(t["bg"] if t["dark"] else "FFFFFF") if i == 0
                   else t["accent"],
                   align=PP_ALIGN.CENTER)
        G.add_text(slide, 1.8, yy - 0.02, W - 2.35, 0.72,
                   st.get("head") or st.get("title", ""),
                   size=19 if W < 9 else 21, bold=True,
                   font=G.HEADING_FONT, color=t["ink"], line_spacing=1.1)
        G.add_text(slide, 1.8, yy + 0.68, W - 2.35, 0.8, st.get("body", ""),
                   size=15.5 if W < 9 else 16.5, color=t["muted"],
                   line_spacing=1.22)


def p_quote(slide, t, d, idx, total, brand, W, H, fmt):
    G.add_text(slide, 0.5, H * 0.22, 1.6, 1.2, "“", size=110, bold=True,
               font=G.HEADING_FONT, color=t["accent"], name="!!p-qmark")
    G.add_text(slide, 0.7, H * 0.33, W - 1.4, 4.4, d.get("quote", ""),
               size=33, bold=True, font=G.HEADING_FONT, color=t["ink"],
               line_spacing=1.22)
    if d.get("attribution"):
        G.add_text(slide, 0.7, H * 0.66, W - 1.4, 0.6,
                   "— " + d["attribution"], size=18, color=t["muted"])


def p_compare(slide, t, d, idx, total, brand, W, H, fmt):
    _title_block(slide, t, d, W, y=H * 0.155, size=36)
    panels = [(d.get("before", {}), t["card"], t["muted"], "✕ "),
              (d.get("after", {}), t["accent"],
               t["bg"] if t["dark"] else "FFFFFF", "✓ ")]
    y = H * 0.345
    for k, (panel, fill, ink, mark) in enumerate(panels):
        ph = H * 0.235
        G.add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, y, W - 1.1, ph,
                    fill=fill, adj=0.07, name=f"!!p-panel-{k}")
        G.add_text(slide, 0.95, y + 0.28, W - 1.9, 0.55,
                   panel.get("head", ""), size=22, bold=True,
                   font=G.HEADING_FONT, color=ink)
        rows = [mark + str(r).lstrip("+- ") for r in panel.get("rows", [])[:4]]
        G.add_text(slide, 0.95, y + 0.95, W - 1.9, ph - 1.1, rows, size=17.5,
                   color=ink, line_spacing=1.3)
        y += ph + 0.42


def p_cta(slide, t, d, idx, total, brand, W, H, fmt):
    G.add_shape(slide, MSO_SHAPE.OVAL, -W * 0.28, H * 0.56, W * 0.8, W * 0.8,
                fill=t["accent"], name="!!p-hero-orb")
    ask = d.get("ask") or " ".join(d.get("title_lines", [])) or "Follow for more"
    asize = 50 if W < 9 else 58
    if len(ask) > 26:
        asize = 38
    if len(ask) > 48:
        asize = 31
    G.add_text(slide, 0.55, H * 0.24, W - 1.1, 3.4, ask, size=asize,
               bold=True, font=G.HEADING_FONT, color=t["ink"],
               line_spacing=1.08, name="!!p-ask")
    if d.get("subtitle"):
        G.add_text(slide, 0.55, H * 0.44, W - 1.1, 1.0, d["subtitle"],
                   size=20, color=t["muted"], line_spacing=1.3)
    handle = d.get("handle") or d.get("contact", "")
    if handle:
        hw = 0.7 + 0.15 * len(handle)
        G.add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, H * 0.555, hw,
                    0.68, fill=t["accent"], adj=0.5, name="!!p-handle")
        G.add_text(slide, 0.55, H * 0.555 + 0.13, hw, 0.45, handle, size=19,
                   bold=True, font=G.HEADING_FONT,
                   color=t["bg"] if t["dark"] else "FFFFFF",
                   align=PP_ALIGN.CENTER)


BUILDERS = {"hook": p_hook, "bigstat": p_bigstat, "tips": p_tips,
            "bars": p_bars, "steps": p_steps, "quote": p_quote,
            "compare": p_compare, "cta": p_cta}


def build_portrait(spec, out_path):
    fmt = spec.get("format", "reel")
    if fmt not in FORMATS:
        raise SystemExit(f"format must be one of {sorted(FORMATS)}")
    W, H = FORMATS[fmt]
    t = dict(G.THEMES.get(spec.get("theme", "teal-orbit"),
                          G.THEMES["teal-orbit"]))
    t.update(spec.get("theme_overrides", {}))
    brand = spec.get("brand", "")
    slides = spec.get("slides", [])
    if not slides:
        raise SystemExit("Spec has no slides.")
    rtl = bool(spec.get("rtl"))
    indic = str(spec.get("numerals", "")).lower() == "arabic-indic"
    font_map = {}
    if spec.get("heading_font"):
        font_map[G.HEADING_FONT] = spec["heading_font"]
    if spec.get("body_font"):
        font_map[G.BODY_FONT] = spec["body_font"]
    if rtl:
        font_map.setdefault(G.HEADING_FONT, "Cairo")
        font_map.setdefault(G.BODY_FONT, "Tajawal")

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    prs.core_properties.title = spec.get("title", f"{fmt} deck")
    blank = prs.slide_layouts[6]
    saved_w = G.SLIDE_W  # _polish mirrors against the module global
    try:
        G.SLIDE_W = Inches(W)
        for idx, sd in enumerate(slides):
            sd = G._expand_tokens(sd, spec)
            slide = prs.slides.add_slide(blank)
            G.set_bg(slide, t["bg"])
            BUILDERS.get(sd.get("type", "tips"), p_tips)(
                slide, t, sd, idx, len(slides), brand, W, H, fmt)
            _chrome(slide, t, idx, len(slides), brand, W, H, fmt)
            G._fix_arabic_shapes(slide)
            if rtl or font_map or indic:
                G._polish(slide, rtl, font_map, indic)
            secs = sd.get("seconds") or spec.get("seconds")
            if not secs and fmt == "reel":
                secs = G._est_seconds(dict(sd, notes=sd.get("notes", "")),
                                      {"timings": True}) or 4.0
            adv = int(float(secs) * 1000) if secs else None
            if idx > 0:
                G.add_morph(slide,
                            duration_ms=int(sd.get("morph_ms",
                                                   spec.get("morph_ms", 700))),
                            adv_ms=adv)
            elif adv:
                G.add_fade_timing(slide, adv)
            if sd.get("notes"):
                slide.notes_slide.notes_text_frame.text = sd["notes"]
    finally:
        G.SLIDE_W = saved_w
    prs.save(out_path)
    return len(slides)


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    with open(argv[1], encoding="utf-8") as f:
        spec = json.load(f)
    out = argv[2] if len(argv) > 2 else os.path.splitext(argv[1])[0] + ".pptx"
    n = build_portrait(spec, out)
    print(f"Built {out}: {n} slides, format={spec.get('format', 'reel')}, "
          f"theme={spec.get('theme', 'teal-orbit')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
