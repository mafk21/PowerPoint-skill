#!/usr/bin/env python3
"""
PlacePlate-style PowerPoint generator.

Builds a 16:9 .pptx deck in the style documented in ../SKILL.md:
  * one motif + one color story (named themes inspired by the PlacePlate catalog)
  * REAL PowerPoint Morph transitions injected as slide XML (python-pptx has no
    native API for transitions, so we add the <p:transition><p14:morph/></p:transition>
    element by hand -- this is the exact markup PowerPoint itself writes)
  * "!!" forced-morph shape naming (the travelling orb, accent bar, kicker and
    footer exist on every slide with the same name, so Morph animates them
    continuously from slide to slide instead of fading)
  * PlacePlate design rules baked in: one idea per slide, assertion headlines,
    3-color discipline, hero stats, "strategic gray" charts, >=16pt text,
    generous margins, Z-pattern layout.

Usage:
    python generate_pptx.py deck-spec.json [output.pptx]
    python generate_pptx.py --themes           # list built-in themes
    python generate_pptx.py --demo out.pptx    # build the bundled demo spec

Requires: python-pptx >= 1.0  (pip install python-pptx)
Present the result in Microsoft PowerPoint 2019+ / Microsoft 365 -- Morph does
not play in Google Slides, Keynote or PDF. Export to MP4 to share the motion.
"""

import json
import os
import sys

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION,
                             XL_LEGEND_POSITION, XL_TICK_MARK)
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

# ----------------------------------------------------------------------------
# Canvas & type scale (16:9, from reference/02-design-system.md)
# ----------------------------------------------------------------------------
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.67)            # >=5% of slide width on every side
CONTENT_W = Inches(13.333 - 1.34)

HEADING_FONT = "Montserrat"       # free font; falls back gracefully
BODY_FONT = "Inter"               # PP-1 Pro Pitch Deck ships Onest/Inter

# ----------------------------------------------------------------------------
# Themes: one dominant color story per deck. Names/moods follow the PlacePlate
# catalog (PP-xx); hex values are reconstructions in the same spirit, sampled
# to match the named palette and preview descriptions -- see
# reference/07-template-catalog.md. Tune freely per project.
# Keys: bg, card (surface), ink (main text), muted (secondary text),
#       accent (the ONE saturated color), accent2 (tint of accent), graybar
# ----------------------------------------------------------------------------
THEMES = {
    "teal-orbit": dict(          # PP-23 style -- startup pitch, dark teal
        bg="04333F", card="0B4A5A", ink="F2FBFA", muted="A9CCC9",
        accent="14C8B4", accent2="7FE7D2", graybar="5E7A80", dark=True),
    "aqua-neumorphic": dict(     # PP-20 style -- thesis/academic, soft light
        bg="E9F5F5", card="FFFFFF", ink="0E3737", muted="587272",
        accent="099B92", accent2="3EC6BE", graybar="B9CFCE", dark=False),
    "royal-blue": dict(          # PP-43 style -- formal academic/corporate
        bg="FFFFFF", card="EEF3FB", ink="10295E", muted="5B6B8C",
        accent="1D4ED8", accent2="7FA6F2", graybar="C3CEDF", dark=False),
    "obsidian-star": dict(       # PP-49/PP-26 style -- premium dark + gold
        bg="0B0B10", card="17171F", ink="F5F2EA", muted="A09CB0",
        accent="D4A537", accent2="F0CE84", graybar="4A4A55", dark=True),
    "steel-blue": dict(          # PP-46 style (their free one) -- SaaS/investor
        bg="0F2233", card="17344D", ink="F2F7FC", muted="A9C2D8",
        accent="4FA3E3", accent2="9CC8EE", graybar="57708A", dark=True),
    "crimson-diamond": dict(     # PP-39 style -- bold business
        bg="1A0A10", card="2A1219", ink="FDF2F4", muted="C9A3AC",
        accent="E11D48", accent2="FB7185", graybar="6E4A52", dark=True),
    "emerald-ring": dict(        # PP-38 style -- growth/sustainability
        bg="06231A", card="0B3527", ink="ECFCF5", muted="A7CDBE",
        accent="10B981", accent2="6EE7B7", graybar="4E7265", dark=True),
    "lilac-soft": dict(          # PP-44 style -- soft creative light
        bg="F6F2FC", card="FFFFFF", ink="35245E", muted="71669A",
        accent="8B5CF6", accent2="C4B5FD", graybar="CFC6E8", dark=False),
}

# External themes (tools/palette_lab.py, brandkit.py, anatomy.py write here)
# merge over the built-ins, so DB palettes are usable by name in specs.
_EXT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                    "database", "data", "themes.json")
if os.path.exists(_EXT):
    try:
        with open(_EXT, encoding="utf-8") as _f:
            THEMES.update(json.load(_f))
    except Exception:
        pass

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def C(hexstr):
    return RGBColor.from_string(hexstr)


# ----------------------------------------------------------------------------
# Morph injection -- the heart of the PlacePlate look.
# PowerPoint writes Morph as a p14 extension inside <p:transition>.
# type: byObject | byWord | byChar  (Effect Options in the UI)
# ----------------------------------------------------------------------------
def _insert_transition(slide, el):
    # Schema order inside <p:sld>: cSld, clrMapOvr?, transition?, timing?, extLst?
    sld = slide._element
    insert_at = 0
    for i, child in enumerate(sld):
        if child.tag.endswith("}cSld") or child.tag.endswith("}clrMapOvr"):
            insert_at = i + 1
    sld.insert(insert_at, el)


def _adv_attrs(adv_ms):
    return f' advClick="1" advTm="{int(adv_ms)}"' if adv_ms else ""


def add_morph(slide, duration_ms=900, morph_type="byObject", speed="med",
              adv_ms=None):
    # p: = presentationml/2006/main, p14: = powerpoint/2010/main (dur + morph)
    xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:p14="{P14}" spd="{speed}" p14:dur="{duration_ms}"{_adv_attrs(adv_ms)}>'
        f'<p14:morph type="{morph_type}"/></p:transition>'
    )
    _insert_transition(slide, etree.fromstring(xml))


def add_fade_timing(slide, adv_ms):
    """Slide 1 can't Morph (no before-state) — a plain fade carrying advTm
    gives kiosk/video exports correct pacing from the first slide."""
    xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'spd="med"{_adv_attrs(adv_ms)}><p:fade/></p:transition>'
    )
    _insert_transition(slide, etree.fromstring(xml))


HIJRI_MONTHS = ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى",
                "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة",
                "ذو الحجة"]


def gregorian_to_hijri(y, m, d):
    """Arithmetical (tabular/Kuwaiti) Hijri conversion. ±1 day vs observed
    calendars is normal — label generated dates accordingly if official."""
    a = (14 - m) // 12
    y2, m2 = y + 4800 - a, m + 12 * a - 3
    jd = d + (153 * m2 + 2) // 5 + 365 * y2 + y2 // 4 - y2 // 100 + y2 // 400 - 32045
    l = jd - 1948440 + 10632
    n = (l - 1) // 10631
    l = l - 10631 * n + 354
    j = ((10985 - l) // 5316) * ((50 * l) // 17719) + (l // 5670) * ((43 * l) // 15238)
    l = l - ((30 - j) // 15) * ((17719 * j) // 50) - (j // 16) * ((15238 * j) // 43) + 29
    hm = (24 * l) // 709
    hd = l - (709 * hm) // 24
    return 30 * n + j - 30, hm, hd


def hijri_str(y, m, d):
    hy, hm, hd = gregorian_to_hijri(y, m, d)
    return f"{hd} {HIJRI_MONTHS[hm - 1]} {hy}هـ"


_INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def _expand_tokens(obj, spec):
    """Replace {hijri} / {gregorian} tokens in slide strings.
    Date from spec['date'] (YYYY-MM-DD) or today."""
    import datetime as _dt
    raw = str(spec.get("date", ""))
    try:
        dt = _dt.date.fromisoformat(raw) if raw else _dt.date.today()
    except ValueError:
        dt = _dt.date.today()
    repl = {"{hijri}": hijri_str(dt.year, dt.month, dt.day),
            "{gregorian}": dt.isoformat()}

    def walk(o):
        if isinstance(o, str):
            for k, v in repl.items():
                o = o.replace(k, v)
            return o
        if isinstance(o, dict):
            return {k: walk(v) for k, v in o.items()}
        if isinstance(o, list):
            return [walk(v) for v in o]
        return o
    return walk(obj)


def _fix_arabic_shapes(slide):
    """Shapes named 'ar!…' (bilingual builder) get true Arabic paragraph
    direction + Cairo/Tajawal even inside an LTR deck (reference/14)."""
    for shape in slide.shapes:
        if not (shape.name or "").startswith("ar!") or not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            para._p.get_or_add_pPr().set("rtl", "1")
            if para.alignment in (None, PP_ALIGN.LEFT):
                para.alignment = PP_ALIGN.RIGHT
            for run in para.runs:
                if run.font.name == HEADING_FONT:
                    run.font.name = "Cairo"
                elif run.font.name == BODY_FONT:
                    run.font.name = "Tajawal"


def _est_seconds(sd, spec):
    """Auto-timing: explicit seconds win; else estimate from speaker notes
    at ~130 wpm when spec.timings is on; clamp 6-45s."""
    if sd.get("seconds"):
        return float(sd["seconds"])
    if spec.get("timings") in (True, "auto"):
        words = len(str(sd.get("notes", "")).split())
        return max(6.0, min(45.0, words / 130 * 60)) if words else 8.0
    return None


# ----------------------------------------------------------------------------
# Shape/text helpers. Every helper accepts `name=` -- pass "!!something" for
# elements that must Morph-match across slides (forced morph naming trick).
# ----------------------------------------------------------------------------
def _no_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def add_shape(slide, kind, x, y, w, h, fill=None, line=None, name=None, adj=None):
    sp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        sp.name = name
    if adj is not None:
        try:
            sp.adjustments[0] = adj
        except Exception:
            pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = C(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = C(line)
        sp.line.width = Pt(1.0)
    _no_shadow(sp)
    sp.text_frame.word_wrap = True
    return sp


def add_text(slide, x, y, w, h, runs, size=18, color="FFFFFF", bold=False,
             font=BODY_FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             name=None, line_spacing=1.12, space_after=Pt(4)):
    """runs: str, or list of paragraphs; each paragraph is str or list of
    (text, dict-overrides) run tuples."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str):
        runs = [runs]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = space_after
        if isinstance(para, str):
            para = [(para, {})]
        for text, over in para:
            r = p.add_run()
            r.text = text
            f = r.font
            f.name = over.get("font", font)
            f.size = Pt(over.get("size", size))
            f.bold = over.get("bold", bold)
            f.color.rgb = C(over.get("color", color))
    return tb


def set_bg(slide, hexstr):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = C(hexstr)


# ----------------------------------------------------------------------------
# Persistent "chrome": elements repeated on (almost) every slide under the
# SAME !! names. Morph treats them as one continuous object, which produces
# PlacePlate's signature travelling-element / floating-motion feel.
# The orb follows a path: huge on the title slide, then tucks into corners.
# ----------------------------------------------------------------------------
ORB_PATH = [  # (x, y, d) per slide index; cycles
    (9.3, -2.2, 6.4),    # title: giant orb top-right (hero)
    (11.9, 5.9, 2.6),    # bottom-right
    (-0.9, 5.6, 2.8),    # bottom-left
    (11.7, -1.0, 2.9),   # top-right
    (-1.1, -1.1, 3.0),   # top-left
    (11.9, 5.9, 2.6),
    (-0.9, 5.6, 2.8),
    (11.7, -1.0, 2.9),
]


def add_chrome(slide, theme, idx, total, brand="", kicker=None):
    t = theme
    x, y, d = ORB_PATH[idx % len(ORB_PATH)]
    # Travelling duo: outer ring + solid core, always the same !! names.
    add_shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill=None,
              line=t["accent2"], name="!!hero-ring")
    add_shape(slide, MSO_SHAPE.OVAL, x + d * 0.22, y + d * 0.22, d * 0.56, d * 0.56,
              fill=t["accent"], name="!!hero-orb")
    # Accent bar under every headline position.
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.67, 1.52, 0.85, 0.09,
              fill=t["accent"], name="!!accent-bar", adj=0.5)
    if kicker:
        add_text(slide, 0.67, 0.52, 8.5, 0.35, kicker.upper(), size=13,
                 color=t["accent"] if not t["dark"] else t["accent2"],
                 bold=True, font=HEADING_FONT, name="!!kicker")
    if brand:
        add_text(slide, 0.67, 7.02, 6.0, 0.3, brand, size=12.5, color=t["muted"],
                 name="!!brand")
    add_text(slide, 12.1, 7.02, 0.6, 0.3, f"{idx + 1:02d}", size=12.5,
             color=t["muted"], align=PP_ALIGN.RIGHT, name="!!pageno")


def headline(slide, theme, text, y=0.92, size=34):
    """Assertion headline: the takeaway sentence, not a topic label."""
    add_text(slide, 0.67, y, 11.4, 0.95, text, size=size, color=theme["ink"],
             bold=True, font=HEADING_FONT, name="!!headline", line_spacing=1.05)


# ----------------------------------------------------------------------------
# Slide builders -- one function per recipe in reference/03-slide-recipes.md
# ----------------------------------------------------------------------------
ICON_GLYPHS = {  # keyword -> widely-supported text glyph (icon engine)
    "check": "\u2713", "cross": "\u2715", "arrow": "\u2192", "star": "\u2605",
    "bolt": "\u26a1", "plus": "\u271a", "dot": "\u25cf", "diamond": "\u25c6",
    "square": "\u25a0", "triangle": "\u25b2", "target": "\u25ce",
    "chart": "\u25ae", "growth": "\u2197", "shield": "\u26e8",
    "gear": "\u2699", "globe": "\u25cd", "idea": "\u2726",
    "money": "$", "time": "\u25f7", "layers": "\u25a4", "link": "\u221e",
    "heart": "\u2665", "flag": "\u2691",
}


def _icon_glyph(item):
    if item.get("glyph"):
        return item["glyph"]
    return ICON_GLYPHS.get(str(item.get("icon", "")).lower(), "\u25c6")


def _mix(c1, c2, f):
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * f):02X}" for x, y in zip(a, b))


def _insight_rail(slide, t, text, label="KEY INSIGHT"):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 9.35, 2.1, 3.3, 4.4,
              fill=t["card"], adj=0.07, name="!!insight-card")
    add_text(slide, 9.65, 2.4, 2.75, 0.5, label, size=12, bold=True,
             font=HEADING_FONT,
             color=t["accent"] if not t["dark"] else t["accent2"])
    add_text(slide, 9.65, 2.85, 2.75, 3.4, text, size=15,
             color=t["ink"], line_spacing=1.3)


def s_title(slide, t, d, idx, total, brand):
    # `title_lines` (list) is an accepted alias of `title` (string with \n);
    # producers in tools/ emit both spellings — never leak the default.
    title = d.get("title") or "\n".join(d.get("title_lines") or []) \
        or "Presentation Title"
    add_text(slide, 0.67, 1.55, 8.6, 2.2, title,
             size=48, color=t["ink"], bold=True, font=HEADING_FONT,
             name="!!hero-title", line_spacing=1.04)
    if d.get("subtitle"):
        add_text(slide, 0.67, 3.55, 7.6, 1.1, d["subtitle"], size=20,
                 color=t["muted"], name="!!hero-sub", line_spacing=1.25)
    if d.get("meta"):
        add_text(slide, 0.67, 5.15, 7.6, 0.8, d["meta"], size=14,
                 color=t["accent2"] if t["dark"] else t["accent"], bold=True,
                 font=HEADING_FONT)


def s_agenda(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "What we'll cover"))
    items = d.get("items", [])[:6]
    active = int(d.get("active", 0))  # reuse as progress tracker in later sections
    y = 2.0
    for i, item in enumerate(items):
        on = (i == active)
        done = (i < active)
        # Navigation dots share !! names -> they "expand" out of the orb system,
        # and re-using this slide with a new `active` walks the highlight (Morph).
        add_shape(slide, MSO_SHAPE.OVAL, 0.67, y + 0.06, 0.52, 0.52,
                  fill=t["accent"] if (on or done) else t["card"],
                  line=None if (on or done) else t["accent"], name=f"!!nav-dot-{i}")
        add_text(slide, 0.67, y + 0.13, 0.52, 0.4, "\u2713" if done else str(i + 1),
                 size=15, bold=True, font=HEADING_FONT, align=PP_ALIGN.CENTER,
                 color=(t["bg"] if ((on or done) and t["dark"]) else
                        ("FFFFFF" if (on or done) else t["ink"])))
        add_text(slide, 1.45, y, 9.5, 0.62, item, size=19,
                 color=t["ink"] if on else t["muted"],
                 bold=on, font=HEADING_FONT)
        y += 0.78


def s_section(slide, t, d, idx, total, brand):
    add_text(slide, 0.67, 2.35, 3.0, 1.6, d.get("number", "01"), size=88,
             color=t["accent"], bold=True, font=HEADING_FONT, name="!!section-no")
    add_text(slide, 0.67, 4.05, 10.5, 1.2, d.get("title", "Section"), size=42,
             color=t["ink"], bold=True, font=HEADING_FONT, name="!!section-title")
    if d.get("subtitle"):
        add_text(slide, 0.67, 5.05, 9.0, 0.9, d["subtitle"], size=18,
                 color=t["muted"], line_spacing=1.25)


def s_content(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "Headline states the takeaway"))
    bullets = d.get("bullets", [])[:4]     # max 3-4: slides are anchors, not scripts
    y = 2.05
    for b in bullets:
        if isinstance(b, str):
            b = {"head": b, "body": ""}
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.67, y, 7.15, 1.02,
                  fill=t["card"], name=None, adj=0.10)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.67, y, 0.12, 1.02,
                  fill=t["accent"], adj=0.5)
        paras = [[(b.get("head", ""), {"size": 17, "bold": True,
                                       "font": HEADING_FONT, "color": t["ink"]})]]
        if b.get("body"):
            paras.append([(b["body"], {"size": 13.5, "color": t["muted"]})])
        add_text(slide, 1.0, y + 0.13, 6.6, 0.8, paras, line_spacing=1.12)
        y += 1.18
    # Visual zone (right): image if provided, else motif panel with big glyph no.
    if d.get("image") and os.path.exists(d["image"]):
        slide.shapes.add_picture(d["image"], Inches(8.25), Inches(2.05),
                                 Inches(4.4), Inches(4.55))
    else:
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 8.25, 2.05, 4.4, 4.55,
                  fill=t["accent"], name="!!visual-panel", adj=0.06)
        add_text(slide, 8.55, 2.55, 3.8, 1.4, d.get("visual_big", "★"), size=60,
                 color=t["bg"] if t["dark"] else "FFFFFF", bold=True,
                 font=HEADING_FONT)
        add_text(slide, 8.55, 4.6, 3.8, 1.7, d.get("visual_caption", ""),
                 size=15, color=t["bg"] if t["dark"] else "FFFFFF",
                 line_spacing=1.25)


def s_icon_grid(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "Features"))
    items = d.get("items", [])[:6]         # 3-6 cards replace a bullet list
    cols = 3 if len(items) > 3 else max(len(items), 1)
    cw, ch, gx, gy = 3.86, 2.02, 0.27, 0.3
    x0, y0 = 0.67, 2.05
    for i, it in enumerate(items):
        if isinstance(it, str):
            it = {"head": it, "body": "", "glyph": "◆"}
        cx = x0 + (i % cols) * (cw + gx)
        cy = y0 + (i // cols) * (ch + gy)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch,
                  fill=t["card"], adj=0.08, name=f"!!card-{i}")
        add_shape(slide, MSO_SHAPE.OVAL, cx + 0.28, cy + 0.26, 0.62, 0.62,
                  fill=t["accent"], name=f"!!card-dot-{i}")
        add_text(slide, cx + 0.28, cy + 0.36, 0.62, 0.45, _icon_glyph(it),
                 size=16, color=t["bg"] if t["dark"] else "FFFFFF", bold=True,
                 align=PP_ALIGN.CENTER, font=HEADING_FONT)
        add_text(slide, cx + 0.28, cy + 1.02, cw - 0.5, 0.42, it.get("head", ""),
                 size=16.5, color=t["ink"], bold=True, font=HEADING_FONT)
        add_text(slide, cx + 0.28, cy + 1.42, cw - 0.5, 0.55, it.get("body", ""),
                 size=12.5, color=t["muted"], line_spacing=1.18)


def s_process(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "How it works"))
    steps = d.get("steps", [])[:5]
    n = max(len(steps), 1)
    cw = (12.0 - (n - 1) * 0.42) / n
    # Connecting spine (morphs/stretches between slides under !! name)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.67, 3.02, 12.0, 0.07,
              fill=t["graybar"], name="!!spine", adj=0.5)
    for i, st in enumerate(steps):
        if isinstance(st, str):
            st = {"head": st, "body": ""}
        cx = 0.67 + i * (cw + 0.42)
        add_shape(slide, MSO_SHAPE.CHEVRON, cx, 2.25, cw, 0.62,
                  fill=t["accent"] if i == 0 else t["card"],
                  line=None if i == 0 else t["accent"], adj=0.35)
        add_text(slide, cx + 0.15, 2.38, cw - 0.3, 0.4, f"STEP {i + 1}",
                 size=12.5, bold=True, font=HEADING_FONT, align=PP_ALIGN.CENTER,
                 color=(t["bg"] if (i == 0 and t["dark"]) else
                        ("FFFFFF" if i == 0 else t["ink"])))
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 3.3, cw, 2.2,
                  fill=t["card"], adj=0.08)
        add_text(slide, cx + 0.24, 3.52, cw - 0.48, 0.75, st.get("head", ""),
                 size=16, color=t["ink"], bold=True, font=HEADING_FONT,
                 line_spacing=1.1)
        add_text(slide, cx + 0.24, 4.35, cw - 0.48, 1.05, st.get("body", ""),
                 size=12.5, color=t["muted"], line_spacing=1.2)


def s_comparison(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "Comparison"))
    cols = d.get("columns", [])[:3]
    n = max(len(cols), 1)
    cw = (12.0 - (n - 1) * 0.35) / n
    for i, col in enumerate(cols):
        cx = 0.67 + i * (cw + 0.35)
        hero = col.get("highlight", i == n - 1)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 2.05, cw, 4.5,
                  fill=t["accent"] if hero else t["card"], adj=0.06,
                  name=f"!!cmp-{i}")
        base = (t["bg"] if t["dark"] else "FFFFFF") if hero else t["ink"]
        sub = base if hero else t["muted"]
        add_text(slide, cx + 0.32, 2.3, cw - 0.6, 0.55, col.get("head", ""),
                 size=19, color=base, bold=True, font=HEADING_FONT)
        paras = []
        for row in col.get("rows", [])[:7]:
            mark, txt = ("✓ ", row[2:]) if row.startswith("+ ") else \
                        (("✕ ", row[2:]) if row.startswith("- ") else ("• ", row))
            paras.append([(mark, {"bold": True, "color": base}),
                          (txt, {"color": sub})])
        add_text(slide, cx + 0.32, 2.95, cw - 0.6, 3.4, paras, size=13.5,
                 line_spacing=1.3, space_after=Pt(7))


def s_chart(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "The chart title states the finding"))
    cats = d.get("categories", ["A", "B", "C", "D"])
    vals = d.get("values", [10, 20, 30, 40])
    highlight = d.get("highlight", len(vals) - 1)   # strategic gray: one accent bar
    cd = ChartData()
    cd.categories = cats
    cd.add_series(d.get("series_name", "Value"), vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                Inches(0.67), Inches(2.1), Inches(8.3), Inches(4.4), cd)
    chart = gf.chart
    chart.has_legend = False
    chart.has_title = False
    ser = chart.series[0]
    ser.has_data_labels = True
    dl = ser.data_labels
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(13)
    dl.font.bold = True
    dl.font.name = BODY_FONT
    dl.font.color.rgb = C(t["ink"])
    for i, pt in enumerate(ser.points):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = C(t["accent"] if i == highlight else t["graybar"])
    va = chart.value_axis
    va.has_major_gridlines = False
    va.has_title = False
    va.tick_labels.font.size = Pt(12)
    va.tick_labels.font.color.rgb = C(t["muted"])
    va.minimum_scale = 0                      # never truncate the axis
    ca = chart.category_axis
    ca.tick_labels.font.size = Pt(12)
    ca.tick_labels.font.color.rgb = C(t["ink"])
    ca.major_tick_mark = XL_TICK_MARK.OUTSIDE
    chart.plots[0].gap_width = 60
    # Insight rail on the right: the "so what"
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 9.35, 2.1, 3.3, 4.4,
              fill=t["card"], adj=0.07, name="!!insight-card")
    add_text(slide, 9.65, 2.4, 2.75, 0.5, "KEY INSIGHT", size=12, bold=True,
             font=HEADING_FONT, color=t["accent"] if not t["dark"] else t["accent2"])
    add_text(slide, 9.65, 2.85, 2.75, 3.4, d.get("insight", ""), size=15,
             color=t["ink"], line_spacing=1.3)


def s_hero_stat(slide, t, d, idx, total, brand):
    # One number, enormous -- the keynote moment. Use at section breaks.
    add_text(slide, 0.67, 1.9, 12.0, 2.3, d.get("stat", "127%"), size=120,
             color=t["accent"], bold=True, font=HEADING_FONT,
             name="!!hero-stat", line_spacing=1.0)
    add_text(slide, 0.67, 4.35, 10.5, 0.7, d.get("label", ""), size=24,
             color=t["ink"], font=HEADING_FONT, name="!!hero-stat-label")
    add_text(slide, 0.67, 5.1, 9.5, 0.6, d.get("context", ""), size=16,
             color=t["muted"])


def s_market(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "A large, reachable market"))
    rings = d.get("rings", [           # TAM -> SAM -> SOM, outermost first
        {"label": "TAM", "value": "$48B", "note": "Total market"},
        {"label": "SAM", "value": "$12B", "note": "Serviceable market"},
        {"label": "SOM", "value": "$1.2B", "note": "Obtainable in 3 years"},
    ])[:3]
    cols = [t["graybar"], t["accent2"], t["accent"]]
    cx, cy, dmax = 1.1, 2.0, 4.6
    for i, r in enumerate(rings):
        dd = dmax - i * 1.28
        off = (dmax - dd) / 2
        add_shape(slide, MSO_SHAPE.OVAL, cx + off, cy + off, dd, dd,
                  fill=cols[i], name=f"!!ring-{i}")
        add_text(slide, 6.6, 2.25 + i * 1.42, 6.0, 1.2,
                 [[(f"{r['value']}  ", {"size": 30, "bold": True,
                                         "font": HEADING_FONT, "color": t["ink"]}),
                   (r["label"], {"size": 15, "bold": True, "font": HEADING_FONT,
                                 "color": t["accent"] if not t["dark"] else t["accent2"]})],
                  [(r.get("note", ""), {"size": 14, "color": t["muted"]})]],
                 line_spacing=1.15)


def s_team(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "The team to execute this"))
    people = d.get("people", [])[:4]
    n = max(len(people), 1)
    cw = (12.0 - (n - 1) * 0.35) / n
    for i, p in enumerate(people):
        cx = 0.67 + i * (cw + 0.35)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 2.1, cw, 4.3,
                  fill=t["card"], adj=0.07, name=f"!!person-{i}")
        initials = "".join(w[0] for w in p.get("name", "?").split()[:2]).upper()
        add_shape(slide, MSO_SHAPE.OVAL, cx + cw / 2 - 0.62, 2.5, 1.24, 1.24,
                  fill=t["accent"])
        add_text(slide, cx + cw / 2 - 0.62, 2.88, 1.24, 0.5, initials, size=20,
                 bold=True, font=HEADING_FONT, align=PP_ALIGN.CENTER,
                 color=t["bg"] if t["dark"] else "FFFFFF")
        add_text(slide, cx + 0.3, 4.0, cw - 0.6, 0.5, p.get("name", ""), size=17,
                 bold=True, font=HEADING_FONT, color=t["ink"], align=PP_ALIGN.CENTER)
        add_text(slide, cx + 0.3, 4.5, cw - 0.6, 0.4, p.get("role", ""), size=13,
                 color=t["accent"] if not t["dark"] else t["accent2"],
                 bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, cx + 0.35, 5.0, cw - 0.7, 1.2, p.get("bio", ""), size=12,
                 color=t["muted"], align=PP_ALIGN.CENTER, line_spacing=1.25)


def s_quote(slide, t, d, idx, total, brand):
    add_text(slide, 0.67, 1.7, 1.6, 1.4, "“", size=96, color=t["accent"],
             bold=True, font=HEADING_FONT, name="!!quote-mark")
    add_text(slide, 0.67, 3.0, 11.6, 2.2, d.get("quote", ""), size=30,
             color=t["ink"], font=HEADING_FONT, line_spacing=1.25)
    add_text(slide, 0.67, 5.35, 10.0, 0.6, d.get("attribution", ""), size=15,
             color=t["muted"], bold=True)


def s_closing(slide, t, d, idx, total, brand):
    ctitle = d.get("title") or "\n".join(d.get("title_lines") or []) \
        or "Thank you"
    add_text(slide, 0.67, 2.0, 11.5, 1.6, ctitle, size=54,
             color=t["ink"], bold=True, font=HEADING_FONT, name="!!closing-title")
    if d.get("ask"):
        add_text(slide, 0.67, 3.6, 10.5, 1.0, d["ask"], size=22,
                 color=t["accent"] if not t["dark"] else t["accent2"],
                 bold=True, font=HEADING_FONT, line_spacing=1.25)
    if d.get("contact"):
        add_text(slide, 0.67, 4.95, 10.0, 0.9, d["contact"], size=16,
                 color=t["muted"], line_spacing=1.3)


def s_stat_strip(slide, t, d, idx, total, brand):
    # Executive summary strip: 3-4 mini hero stats (QBR openers, proof walls).
    headline(slide, t, d.get("title", "The quarter in numbers"))
    stats = d.get("stats", [])[:4]
    n = max(len(stats), 1)
    cw = (12.0 - (n - 1) * 0.35) / n
    for i, s in enumerate(stats):
        cx = 0.67 + i * (cw + 0.35)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 2.25, cw, 3.3,
                  fill=t["card"], adj=0.07, name=f"!!stat-{i}")
        add_text(slide, cx + 0.3, 2.6, cw - 0.6, 1.1, s.get("value", "\u2014"),
                 size=42, color=t["accent"], bold=True, font=HEADING_FONT)
        add_text(slide, cx + 0.3, 3.85, cw - 0.6, 0.7, s.get("label", ""),
                 size=15.5, color=t["ink"], bold=True, font=HEADING_FONT,
                 line_spacing=1.15)
        add_text(slide, cx + 0.3, 4.55, cw - 0.6, 0.85, s.get("context", ""),
                 size=12.5, color=t["muted"], line_spacing=1.25)


def s_before_after(slide, t, d, idx, total, brand):
    # Transformation slide (service pages, redesigns, what-is -> what-could-be).
    headline(slide, t, d.get("title", "From typical to transformed"))
    panels = [(d.get("before", {}), False, 0.67, "\u2715 "),
              (d.get("after", {}), True, 7.17, "\u2713 ")]
    for panel, hero, cx, mark in panels:
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 2.05, 5.5, 4.5,
                  fill=t["accent"] if hero else t["card"], adj=0.06,
                  name="!!ba-after" if hero else "!!ba-before")
        base = (t["bg"] if t["dark"] else "FFFFFF") if hero else t["ink"]
        sub = base if hero else t["muted"]
        add_text(slide, cx + 0.35, 2.32, 4.8, 0.55,
                 panel.get("head", "After" if hero else "Before"),
                 size=19, color=base, bold=True, font=HEADING_FONT)
        paras = [[(mark, {"bold": True, "color": base}), (r, {"color": sub})]
                 for r in panel.get("rows", [])[:6]]
        add_text(slide, cx + 0.35, 3.0, 4.85, 3.3, paras, size=14,
                 line_spacing=1.3, space_after=Pt(8))
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW, 6.36, 3.95, 0.62, 0.7,
              fill=t["accent"], name="!!ba-arrow", adj=0.55)


def s_logo_wall(slide, t, d, idx, total, brand):
    # Social proof wall: client/partner wordmarks in a calm grid.
    headline(slide, t, d.get("title", "Trusted by teams like yours"))
    names = d.get("names", [])[:8]
    cols, cw, ch, gx, gy = 4, 2.86, 1.7, 0.19, 0.25
    for i, nm in enumerate(names):
        cx = 0.67 + (i % cols) * (cw + gx)
        cy = 2.2 + (i // cols) * (ch + gy)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch,
                  fill=t["card"], adj=0.10, name=f"!!logo-{i}")
        add_text(slide, cx + 0.2, cy + 0.62, cw - 0.4, 0.5, nm, size=18,
                 color=t["ink"], bold=True, font=HEADING_FONT,
                 align=PP_ALIGN.CENTER)
    if d.get("caption"):
        add_text(slide, 0.67, 6.15, 12.0, 0.5, d["caption"], size=14,
                 color=t["muted"], align=PP_ALIGN.CENTER)


def s_line_chart(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "The trend is the finding"))
    series = d.get("series") or [{"name": d.get("series_name", "Value"),
                                   "values": d.get("values", [1, 2, 3])}]
    highlight = int(d.get("highlight", len(series) - 1))
    cd = ChartData()
    cd.categories = d.get("categories", [])
    for srs in series:
        cd.add_series(srs.get("name", "s"), srs.get("values", []))
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.67), Inches(2.1), Inches(8.3),
                                Inches(4.4), cd)
    chart = gf.chart
    chart.has_title = False
    chart.has_legend = len(series) > 1
    if len(series) > 1:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(11)
    for i, ser in enumerate(chart.series):
        col = t["accent"] if i == highlight else t["graybar"]
        ser.format.line.color.rgb = C(col)
        ser.format.line.width = Pt(2.75 if i == highlight else 1.75)
        ser.smooth = False
    va, ca = chart.value_axis, chart.category_axis
    va.has_major_gridlines = False
    va.tick_labels.font.size = Pt(12)
    va.tick_labels.font.color.rgb = C(t["muted"])
    ca.tick_labels.font.size = Pt(12)
    ca.tick_labels.font.color.rgb = C(t["ink"])
    _insight_rail(slide, t, d.get("insight", ""))


def s_donut(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "Where the whole splits"))
    cats = d.get("categories", [])[:4]      # part-to-whole: max 4 (ref/05)
    vals = d.get("values", [])[:4]
    highlight = int(d.get("highlight", 0))
    tints = [t["graybar"], _mix(t["graybar"], t["accent2"], 0.45), t["muted"]]
    cd = ChartData()
    cd.categories = cats
    cd.add_series("share", vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
                                Inches(0.9), Inches(1.95), Inches(5.9),
                                Inches(4.7), cd)
    chart = gf.chart
    chart.has_title = False
    chart.has_legend = False
    ser = chart.series[0]
    ser.has_data_labels = True
    dl = ser.data_labels
    dl.show_percentage = True
    dl.font.size = Pt(13)
    dl.font.bold = True
    dl.font.color.rgb = C(t["ink"])
    ti = 0
    for i, pt in enumerate(ser.points):
        pt.format.fill.solid()
        if i == highlight:
            pt.format.fill.fore_color.rgb = C(t["accent"])
        else:
            pt.format.fill.fore_color.rgb = C(tints[ti % len(tints)]); ti += 1
    total_v = sum(vals) or 1
    add_text(slide, 2.75, 3.85, 2.2, 1.0,
             d.get("center", f"{round(vals[highlight] / total_v * 100)}%"
                   if vals else ""),
             size=34, color=t["accent"], bold=True, font=HEADING_FONT,
             align=PP_ALIGN.CENTER, name="!!donut-center")
    y = 2.3
    ti = 0
    for i, (c, v) in enumerate(zip(cats, vals)):
        col = t["accent"] if i == highlight else tints[ti % len(tints)]
        if i != highlight:
            ti += 1
        add_shape(slide, MSO_SHAPE.OVAL, 7.6, y + 0.08, 0.3, 0.3, fill=col)
        add_text(slide, 8.1, y, 4.5, 0.5,
                 [[(f"{c}  ", {"bold": True, "color": t["ink"]}),
                   (f"{v} ({round(v / total_v * 100)}%)", {"color": t["muted"]})]],
                 size=15)
        y += 0.62
    add_text(slide, 7.6, y + 0.25, 5.0, 1.6, d.get("insight", ""), size=14.5,
             color=t["ink"], line_spacing=1.3)


def s_waterfall(slide, t, d, idx, total, brand):
    """Shape-built waterfall (native charts can't Morph bar-by-bar; !! bars can).
    steps: [{label, value, kind: start|delta|total}] — start/total absolute,
    deltas signed. Down = Okabe vermilion (semantic, CVD-safe)."""
    headline(slide, t, d.get("title", "How we get from A to B"))
    steps = d.get("steps", [])[:7]
    cum, bars = 0.0, []
    for st in steps:
        kind = st.get("kind", "delta")
        v = float(st.get("value", 0))
        if kind in ("start", "total"):
            bars.append((st.get("label", ""), 0.0, v, kind, v)); cum = v
        else:
            bars.append((st.get("label", ""), cum, cum + v, kind, v)); cum += v
    vmax = max([max(a, b) for _, a, b, _, _ in bars] + [1.0])
    x0, base_y, plot_h, plot_w = 0.9, 5.85, 3.55, 11.5
    n = max(len(bars), 1)
    bw = min(1.15, (plot_w / n) * 0.62)
    step_w = plot_w / n
    prev_top = None
    for i, (label, lo, hi, kind, v) in enumerate(bars):
        cx = x0 + i * step_w + (step_w - bw) / 2
        top = base_y - (hi / vmax) * plot_h
        bot = base_y - (lo / vmax) * plot_h
        if kind == "total":
            col = t["accent"]
        elif kind == "start":
            col = t["graybar"]
        elif v < 0:
            col = "D55E00"
        else:
            col = t["accent2"] if t["dark"] else t["accent"]
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, top, bw, bot - top,
                  fill=col, adj=0.06, name=f"!!wf-{i}")
        if prev_top is not None:
            add_shape(slide, MSO_SHAPE.RECTANGLE, cx - (step_w - bw) / 2 - 0.02,
                      prev_top, (step_w - bw) / 2 + 0.04, 0.022, fill=t["muted"])
        prev_top = top if (kind != "delta" or v >= 0) else bot
        fmt = d.get("format", "{:+.0f}" if kind == "delta" else "{:.0f}")
        add_text(slide, cx - 0.3, top - 0.42, bw + 0.6, 0.35, fmt.format(v),
                 size=13.5, bold=True, color=t["ink"], align=PP_ALIGN.CENTER)
        add_text(slide, cx - 0.35, base_y + 0.12, bw + 0.7, 0.55, label,
                 size=12, color=t["muted"], align=PP_ALIGN.CENTER,
                 line_spacing=1.1)
    add_shape(slide, MSO_SHAPE.RECTANGLE, x0 - 0.2, base_y, plot_w + 0.4, 0.03,
              fill=t["muted"], name="!!wf-axis")
    if d.get("insight"):
        add_text(slide, 0.9, 6.55, 11.5, 0.4, d["insight"], size=13.5,
                 color=t["ink"])


def s_matrix(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "Where we win"))
    px, py, pw, ph = 0.67, 2.05, 8.3, 4.45
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, ph,
              fill=t["card"], adj=0.04, name="!!matrix-plot")
    add_shape(slide, MSO_SHAPE.RECTANGLE, px + pw / 2 - 0.012, py + 0.15, 0.024,
              ph - 0.3, fill=t["graybar"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, px + 0.15, py + ph / 2 - 0.012,
              pw - 0.3, 0.024, fill=t["graybar"])
    quads = (d.get("quadrants") or ["", "", "", ""])[:4]  # TL TR BL BR
    quad_xy = [(px + 0.25, py + 0.18), (px + pw - 2.35, py + 0.18),
               (px + 0.25, py + ph - 0.55), (px + pw - 2.35, py + ph - 0.55)]
    for (qx, qy), q in zip(quad_xy, quads):
        if q:
            add_text(slide, qx, qy, 2.1, 0.35, q.upper(), size=12.5, bold=True,
                     font=HEADING_FONT, color=t["muted"],
                     align=PP_ALIGN.RIGHT if qx > px + pw / 2 else PP_ALIGN.LEFT)
    for i, it in enumerate(d.get("items", [])[:8]):
        cx = px + 0.5 + float(it.get("x", 0.5)) * (pw - 1.0)
        cy = py + 0.45 + (1 - float(it.get("y", 0.5))) * (ph - 0.95)
        hero = it.get("highlight", False)
        dsz = 0.34 if hero else 0.26
        add_shape(slide, MSO_SHAPE.OVAL, cx - dsz / 2, cy - dsz / 2, dsz, dsz,
                  fill=t["accent"] if hero else t["graybar"],
                  name=f"!!mx-{i}")
        add_text(slide, cx + 0.24, cy - 0.16, 2.2, 0.35, it.get("label", ""),
                 size=13, bold=hero, color=t["ink"] if hero else t["muted"])
    add_text(slide, px, py + ph + 0.14, pw, 0.35, d.get("x_label", ""),
             size=12.5, color=t["muted"], bold=True, align=PP_ALIGN.CENTER)
    _insight_rail(slide, t, d.get("insight", ""),
                  label=d.get("y_label", "KEY INSIGHT"))


def s_funnel(slide, t, d, idx, total, brand):
    headline(slide, t, d.get("title", "Where the funnel leaks"))
    stages = d.get("stages", [])[:5]
    vmax = max([float(s.get("value", 0)) for s in stages] + [1.0])
    cx_mid, top, bh, gap = 4.6, 2.1, 0.78, 0.16
    for i, st in enumerate(stages):
        v = float(st.get("value", 0))
        w = 1.6 + (v / vmax) * 5.6
        col = _mix(t["accent"], t["accent2"], i / max(len(stages) - 1, 1))
        yy = top + i * (bh + gap)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx_mid - w / 2, yy, w, bh,
                  fill=col, adj=0.14, name=f"!!fn-{i}")
        base = t["bg"] if t["dark"] else "FFFFFF"
        if w >= 3.0:  # wide bar: label inside, on the color
            add_text(slide, cx_mid - w / 2 + 0.25, yy + 0.2, w - 0.5, 0.42,
                     [[(f"{st.get('label', '')}  ", {"bold": True, "color": base}),
                       (f"{v:,.0f}", {"color": base})]], size=14.5)
        else:  # narrow bar: inside text would wrap+spill (QA sentinel catch)
            lx = cx_mid + w / 2 + 0.22
            add_text(slide, lx, yy + 0.2, max(8.3 - lx, 1.2), 0.42,
                     [[(f"{st.get('label', '')}  ", {"bold": True, "color": t["ink"]}),
                       (f"{v:,.0f}", {"color": t["muted"]})]], size=14.5)
        if i > 0:
            prev = float(stages[i - 1].get("value", 0)) or 1
            add_text(slide, cx_mid + 3.9, yy + 0.18, 2.4, 0.4,
                     f"{round(v / prev * 100)}% \u2193", size=14, bold=True,
                     color=t["accent"] if not t["dark"] else t["accent2"])
    add_text(slide, 9.0, 2.3, 3.65, 3.6, d.get("insight", ""), size=14.5,
             color=t["ink"], line_spacing=1.35)


def s_bilingual(slide, t, d, idx, total, brand):
    """Arabic (right, reads first) + English (left) twins on one slide.
    AR shapes are named 'ar!…' -> _fix_arabic_shapes gives them rtl=1 +
    Cairo/Tajawal; EN shapes 'ltr!…' -> _polish mirrors position only, so
    the slide is correct in both RTL and LTR decks (reference/14)."""
    accent = t["accent"] if not t["dark"] else t["accent2"]
    add_text(slide, 6.77, 0.92, 5.9, 1.15,
             d.get("ar_title") or d.get("title", ""), size=29, bold=True,
             font=HEADING_FONT, color=t["ink"], align=PP_ALIGN.RIGHT,
             line_spacing=1.12, name="ar!ar-title")
    add_text(slide, 0.67, 0.92, 5.9, 1.15,
             d.get("en_title") or d.get("title", ""), size=27, bold=True,
             font=HEADING_FONT, color=t["ink"], align=PP_ALIGN.LEFT,
             line_spacing=1.12, name="ltr!en-title")
    add_shape(slide, MSO_SHAPE.RECTANGLE, 6.652, 2.15, 0.03, 4.35,
              fill=t["graybar"], name="!!bi-divider")
    panels = [  # (x, label, bullets key, body key, align, prefix)
        (6.77, "العربية", "ar_bullets", "ar_body", PP_ALIGN.RIGHT, "ar!"),
        (0.67, "ENGLISH", "en_bullets", "en_body", PP_ALIGN.LEFT, "ltr!"),
    ]
    for x, label, bkey, bodykey, align, pfx in panels:
        add_text(slide, x, 2.18, 5.9, 0.4, label, size=12.5, bold=True,
                 font=HEADING_FONT, color=accent, align=align,
                 name=pfx + "label")
        y = 2.72
        for i, b in enumerate(d.get(bkey, [])[:4]):
            add_text(slide, x, y, 5.9, 0.42, b.get("head", ""), size=17,
                     bold=True, font=HEADING_FONT, color=t["ink"],
                     align=align, name=f"{pfx}h{i}")
            if b.get("body"):
                add_text(slide, x, y + 0.4, 5.9, 0.62, b["body"], size=13.5,
                         color=t["muted"], align=align, line_spacing=1.2,
                         name=f"{pfx}b{i}")
            y += 1.02
        if d.get(bodykey):
            add_text(slide, x, y, 5.9, 2.2, d[bodykey], size=14.5,
                     color=t["ink"], align=align, line_spacing=1.35,
                     name=pfx + "body")


BUILDERS = {
    "title": s_title, "agenda": s_agenda, "section": s_section,
    "content": s_content, "icon_grid": s_icon_grid, "process": s_process,
    "timeline": s_process, "comparison": s_comparison, "chart": s_chart,
    "hero_stat": s_hero_stat, "market": s_market, "tam_sam_som": s_market,
    "team": s_team, "quote": s_quote, "closing": s_closing,
    "stat_strip": s_stat_strip, "before_after": s_before_after,
    "logo_wall": s_logo_wall,
    "line_chart": s_line_chart, "donut": s_donut, "waterfall": s_waterfall,
    "matrix": s_matrix, "funnel": s_funnel, "bilingual": s_bilingual,
}

# Which slides get the standard headline chrome vs. a bespoke layout
CHROME_KICKER = {"agenda", "content", "icon_grid", "process", "timeline",
                 "comparison", "chart", "market", "tam_sam_som", "team",
                 "stat_strip", "before_after", "logo_wall",
                 "line_chart", "donut", "waterfall", "matrix", "funnel", "bilingual"}


# ----------------------------------------------------------------------------
# RTL / font polish pass (reference/14): mirrors geometry, sets paragraph
# direction rtl=1 + right alignment, flips directional shapes, remaps fonts.
# ----------------------------------------------------------------------------
_FLIP_TYPES = None  # lazily built (MSO_SHAPE members)


def _polish(slide, rtl, font_map, indic=False):
    global _FLIP_TYPES
    if _FLIP_TYPES is None:
        _FLIP_TYPES = {MSO_SHAPE.CHEVRON, MSO_SHAPE.RIGHT_ARROW,
                       MSO_SHAPE.PENTAGON}
    for shape in slide.shapes:
        # 'ltr!…' shapes (English panel of bilingual slides): mirror the
        # position like everything else, but never touch direction/fonts.
        ltr_panel = (shape.name or "").startswith("ltr!")
        if shape.has_text_frame and not ltr_panel:
            for para in shape.text_frame.paragraphs:
                if rtl:
                    para._p.get_or_add_pPr().set("rtl", "1")
                    if para.alignment in (None, PP_ALIGN.LEFT):
                        para.alignment = PP_ALIGN.RIGHT
                for run in para.runs:
                    if run.font.name in font_map:
                        run.font.name = font_map[run.font.name]
                    if indic and run.text:
                        run.text = run.text.translate(_INDIC)
        if rtl:
            try:
                shape.left = int(SLIDE_W) - (int(shape.left) + int(shape.width))
            except Exception:
                pass
            try:
                if shape.auto_shape_type in _FLIP_TYPES:
                    xfrm = shape._element.spPr.find(qn("a:xfrm"))
                    if xfrm is not None:
                        xfrm.set("flipH", "1")
            except Exception:
                pass


def build(spec, out_path):
    theme_name = spec.get("theme", "teal-orbit")
    t = dict(THEMES.get(theme_name, THEMES["teal-orbit"]))
    t.update(spec.get("theme_overrides", {}))
    brand = spec.get("brand", "")
    slides = spec.get("slides", [])
    rtl = bool(spec.get("rtl"))
    indic = str(spec.get("numerals", "")).lower() == "arabic-indic"
    font_map = {}
    if spec.get("heading_font"):
        font_map[HEADING_FONT] = spec["heading_font"]
    if spec.get("body_font"):
        font_map[BODY_FONT] = spec["body_font"]
    if rtl:  # Arabic defaults (reference/14)
        font_map.setdefault(HEADING_FONT, "Cairo")
        font_map.setdefault(BODY_FONT, "Tajawal")
    if not slides:
        raise SystemExit("Spec has no slides.")

    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    prs.core_properties.title = spec.get("title", "Presentation")
    prs.core_properties.author = spec.get("author", "")
    blank = prs.slide_layouts[6]

    for idx, sd in enumerate(slides):
        sd = _expand_tokens(sd, spec)  # {hijri} / {gregorian} in any string
        stype = sd.get("type", "content")
        slide = prs.slides.add_slide(blank)
        set_bg(slide, t["bg"])
        add_chrome(slide, t, idx, len(slides), brand=brand,
                   kicker=sd.get("kicker") if stype in CHROME_KICKER else None)
        BUILDERS.get(stype, s_content)(slide, t, sd, idx, len(slides), brand)
        _fix_arabic_shapes(slide)
        if rtl or font_map or indic:
            _polish(slide, rtl, font_map, indic)
        secs = _est_seconds(sd, spec)
        adv = int(secs * 1000) if secs else None
        if idx > 0:  # Morph needs a "before" slide -- never on slide 1
            add_morph(slide,
                      duration_ms=int(sd.get("morph_ms", spec.get("morph_ms", 900))),
                      morph_type=sd.get("morph_type", "byObject"), adv_ms=adv)
        elif adv:    # timed kiosk/video pacing from slide 1 (fade carries advTm)
            add_fade_timing(slide, adv)
        if sd.get("notes"):
            slide.notes_slide.notes_text_frame.text = sd["notes"]

    prs.save(out_path)
    return len(slides)


DEMO_SPEC = {
    "title": "Demo — run with --demo; replace with your own spec",
    "theme": "teal-orbit", "brand": "DEMO DECK",
    "slides": [
        {"type": "title", "title": "Your Big Idea,\nPresented Beautifully",
         "subtitle": "A PlacePlate-style Morph deck generated from JSON.",
         "meta": "GENERATOR DEMO"},
        {"type": "agenda", "kicker": "Agenda", "title": "What we'll cover",
         "items": ["The problem", "Our solution", "Market", "The ask"]},
        {"type": "hero_stat", "stat": "127%", "label": "Year-over-year growth",
         "context": "Compared to a 34% industry average."},
        {"type": "closing", "title": "Thank you",
         "ask": "Replace this spec with your own content.",
         "contact": "See deck-spec.example.json"},
    ],
}


def main(argv):
    if len(argv) >= 2 and argv[1] == "--themes":
        for name, t in THEMES.items():
            print(f"{name:18s} bg=#{t['bg']} accent=#{t['accent']} "
                  f"{'dark' if t['dark'] else 'light'}")
        return 0
    if len(argv) >= 2 and argv[1] == "--demo":
        out = argv[2] if len(argv) > 2 else "demo.pptx"
        spec = DEMO_SPEC
    elif len(argv) >= 2:
        with open(argv[1], encoding="utf-8") as f:
            spec = json.load(f)
        out = argv[2] if len(argv) > 2 else os.path.splitext(argv[1])[0] + ".pptx"
    else:
        print(__doc__)
        return 2
    n = build(spec, out)
    print(f"Built {out}: {n} slides, theme={spec.get('theme', 'teal-orbit')}, "
          f"Morph on slides 2..{n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
