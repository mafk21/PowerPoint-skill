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
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_MARK
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
        bg="E9F5F5", card="FFFFFF", ink="0E3737", muted="5F7B7B",
        accent="0AA9A0", accent2="3EC6BE", graybar="B9CFCE", dark=False),
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
        bg="F6F2FC", card="FFFFFF", ink="35245E", muted="7A6FA0",
        accent="8B5CF6", accent2="C4B5FD", graybar="CFC6E8", dark=False),
}

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def C(hexstr):
    return RGBColor.from_string(hexstr)


# ----------------------------------------------------------------------------
# Morph injection -- the heart of the PlacePlate look.
# PowerPoint writes Morph as a p14 extension inside <p:transition>.
# type: byObject | byWord | byChar  (Effect Options in the UI)
# ----------------------------------------------------------------------------
def add_morph(slide, duration_ms=900, morph_type="byObject", speed="med"):
    # p: = presentationml/2006/main, p14: = powerpoint/2010/main (dur attr + morph el)
    xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:p14="{P14}" spd="{speed}" p14:dur="{duration_ms}">'
        f'<p14:morph type="{morph_type}"/></p:transition>'
    )
    el = etree.fromstring(xml)
    sld = slide._element
    # Schema order inside <p:sld>: cSld, clrMapOvr?, transition?, timing?, extLst?
    insert_at = 0
    for i, child in enumerate(sld):
        if child.tag.endswith("}cSld") or child.tag.endswith("}clrMapOvr"):
            insert_at = i + 1
    sld.insert(insert_at, el)


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
def s_title(slide, t, d, idx, total, brand):
    add_text(slide, 0.67, 1.55, 8.6, 2.2, d.get("title", "Presentation Title"),
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
        add_text(slide, cx + 0.28, cy + 0.36, 0.62, 0.45, it.get("glyph", "◆"),
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
    add_text(slide, 0.67, 2.0, 11.5, 1.6, d.get("title", "Thank you"), size=54,
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


BUILDERS = {
    "title": s_title, "agenda": s_agenda, "section": s_section,
    "content": s_content, "icon_grid": s_icon_grid, "process": s_process,
    "timeline": s_process, "comparison": s_comparison, "chart": s_chart,
    "hero_stat": s_hero_stat, "market": s_market, "tam_sam_som": s_market,
    "team": s_team, "quote": s_quote, "closing": s_closing,
    "stat_strip": s_stat_strip, "before_after": s_before_after,
    "logo_wall": s_logo_wall,
}

# Which slides get the standard headline chrome vs. a bespoke layout
CHROME_KICKER = {"agenda", "content", "icon_grid", "process", "timeline",
                 "comparison", "chart", "market", "tam_sam_som", "team",
                 "stat_strip", "before_after", "logo_wall"}


# ----------------------------------------------------------------------------
# RTL / font polish pass (reference/14): mirrors geometry, sets paragraph
# direction rtl=1 + right alignment, flips directional shapes, remaps fonts.
# ----------------------------------------------------------------------------
_FLIP_TYPES = None  # lazily built (MSO_SHAPE members)


def _polish(slide, rtl, font_map):
    global _FLIP_TYPES
    if _FLIP_TYPES is None:
        _FLIP_TYPES = {MSO_SHAPE.CHEVRON, MSO_SHAPE.RIGHT_ARROW,
                       MSO_SHAPE.PENTAGON}
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                if rtl:
                    para._p.get_or_add_pPr().set("rtl", "1")
                    if para.alignment in (None, PP_ALIGN.LEFT):
                        para.alignment = PP_ALIGN.RIGHT
                for run in para.runs:
                    if run.font.name in font_map:
                        run.font.name = font_map[run.font.name]
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
        stype = sd.get("type", "content")
        slide = prs.slides.add_slide(blank)
        set_bg(slide, t["bg"])
        add_chrome(slide, t, idx, len(slides), brand=brand,
                   kicker=sd.get("kicker") if stype in CHROME_KICKER else None)
        BUILDERS.get(stype, s_content)(slide, t, sd, idx, len(slides), brand)
        if rtl or font_map:
            _polish(slide, rtl, font_map)
        if idx > 0:  # Morph needs a "before" slide -- never on slide 1
            add_morph(slide,
                      duration_ms=int(sd.get("morph_ms", spec.get("morph_ms", 900))),
                      morph_type=sd.get("morph_type", "byObject"))
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
