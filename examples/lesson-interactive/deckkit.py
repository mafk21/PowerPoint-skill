#!/usr/bin/env python3
"""deckkit — a small, opinionated design system on top of python-pptx.

Everything the lesson rebuild needs, in one place:

  * design tokens (one motif, one color story, two fonts, one type scale)
  * primitives: cards, pills, chips, badges, chat bubbles, images, icons
  * real **Morph** transitions (p14:morph) on every slide
  * real **click triggers** (interactive timing sequences) so a student can
    tap an answer card and the feedback appears on that same slide
  * hyperlink navigation (ppaction://hlinksldjump) for menus / back buttons
  * embedded audio with the exact `<a:audioFile>` + `p14:media` shape
    PowerPoint writes itself

The rules encoded here come from SKILL.md + reference/02 (design system),
01 (Morph), 11 (interactivity) — see examples/lesson-interactive/README.md.
"""
from __future__ import annotations

import copy
import io
import os

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Inches, Pt

from PIL import Image

# ---------------------------------------------------------------------------
# Tokens
# ---------------------------------------------------------------------------
W, H = 13.333, 7.5
MX = 0.62                      # side margin
CW = W - 2 * MX                # content width

BG = "FFFFFF"
BG_TINT = "F4F8FF"
CARD = "FFFFFF"
LINE = "DCE5F2"

INK = "0F1B33"
INK_SOFT = "46587A"
MUTED = "5C6C86"          # 5.3:1 on white
MUTED_DARKBG = "A9BCD9"   # light-on-navy, >6:1
DARK_FOOT = "93A9CC"
FOOT = "5B6A82"           # 5.5:1 on white

ACCENT = "1D4ED8"
ACCENT_DK = "1638A8"
ACCENT_SOFT = "E8EFFF"
ACCENT_LINE = "C3D4FF"
ACCENT_TEXT = "1B3FA8"

DARK = "0B1B33"
DARK_2 = "14284A"
DARK_LINE = "2A4470"

OK = "0F6B33"
OK_SOFT = "E7F6EC"
NO = "B42318"
NO_SOFT = "FDECEA"
WARN = "B45309"
WARN_SOFT = "FEF3C7"

HEAD = "Trebuchet MS"
BODY = "Calibri"

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def C(hexstr):
    return RGBColor.from_string(hexstr)


def I(v):
    return Emu(int(round(v * 914400)))


# ---------------------------------------------------------------------------
# low level
# ---------------------------------------------------------------------------
def _spPr(shape):
    return shape._element.spPr


def set_alpha(shape, opacity_pct, which="fill"):
    """opacity_pct: 0..100 (100 = fully opaque)."""
    val = int(round(opacity_pct * 1000))
    spPr = _spPr(shape)
    tag = "a:solidFill" if which == "fill" else "a:ln"
    node = spPr.find(qn(tag))
    if node is None:
        return
    if which == "line":
        node = node.find(qn("a:solidFill"))
        if node is None:
            return
    srgb = node.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for old in srgb.findall(qn("a:alpha")):
        srgb.remove(old)
    srgb.append(etree.fromstring(f'<a:alpha {nsdecls("a")} val="{val}"/>'))


def no_line(shape):
    shape.line.fill.background()


def shadow(shape, blur=18, dist=4, opacity=9, color=INK):
    """Soft card shadow (PowerPoint renders it; harmless everywhere else)."""
    spPr = _spPr(shape)
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    xml = (
        f'<a:effectLst {nsdecls("a")}><a:outerShdw blurRad="{int(blur * 12700)}" '
        f'dist="{int(dist * 12700)}" dir="5400000" rotWithShape="0">'
        f'<a:srgbClr val="{color}"><a:alpha val="{int(opacity * 1000)}"/></a:srgbClr>'
        f"</a:outerShdw></a:effectLst>"
    )
    spPr.append(etree.fromstring(xml))


def _name(shape, name):
    if name:
        shape.name = name
    return shape


def rect(s, x, y, w, h, fill=None, line=None, lw=1.0, radius=None,
         opacity=100, line_opacity=100, name=None, prst=None, dash=None):
    if prst is None:
        prst = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = s.shapes.add_shape(prst, I(x), I(y), I(w), I(h))
    _name(shp, name)
    if radius and prst == MSO_SHAPE.ROUNDED_RECTANGLE:
        # radius given in inches -> adjustment fraction of the short side
        try:
            shp.adjustments[0] = min(0.5, radius / min(w, h))
        except Exception:
            pass
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
        if opacity < 100:
            set_alpha(shp, opacity)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(lw)
        if line_opacity < 100:
            set_alpha(shp, line_opacity, which="line")
        if dash:
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    else:
        no_line(shp)
    # kill the default "theme" text styling
    shp.text_frame.word_wrap = True
    shp.text_frame.text = ""
    return shp


def line(s, x1, y1, x2, y2, color=ACCENT, lw=2.0, opacity=100):
    from pptx.enum.shapes import MSO_CONNECTOR
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    cn.line.color.rgb = C(color)
    cn.line.width = Pt(lw)
    if opacity < 100:
        set_alpha(cn, opacity, which="line")
    return cn


# ---------------------------------------------------------------------------
# text
# ---------------------------------------------------------------------------
def _para(tf, first, align, line_spacing, space_before, space_after):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    return p


def _run(p, txt, size, color, bold, font, italic, spc, underline=False):
    r = p.add_run()
    r.text = txt
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    f.color.rgb = C(color)
    if underline:
        f.underline = True
    if spc:
        r._r.get_or_add_rPr().set("spc", str(int(spc)))
    return r


def text(s, x, y, w, h, content, size=18, color=INK, bold=False, font=BODY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=0.95,
         space_before=0, space_after=0, margin=(0, 0, 0, 0), name=None,
         spc=None, italic=False, wrap=True, underline=False):
    """content: str (\\n = new paragraph) or list of paragraph dicts:
       {"runs": [(text, {overrides})], "align": ..., "space_before": ...}.

    All runs inside one paragraph dict stay in the same paragraph."""
    box = s.shapes.add_textbox(I(x), I(y), I(w), I(h))
    _name(box, name)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    l, t, r_, b = margin
    tf.margin_left, tf.margin_top = I(l), I(t)
    tf.margin_right, tf.margin_bottom = I(r_), I(b)
    tf.auto_size = None
    if isinstance(content, str):
        paras = [{"runs": [(t_, {})]} for t_ in content.split("\n")]
    else:
        paras = content
    first = True
    for para in paras:
        p_align = para.get("align", align)
        p_sb = para.get("space_before", space_before)
        p_sa = para.get("space_after", space_after)
        p_size = para.get("size", size)
        p_color = para.get("color", color)
        p_bold = para.get("bold", bold)
        p_ls = para.get("line_spacing", line_spacing)
        p = _para(tf, first, p_align, p_ls, p_sb, p_sa)
        first = False
        for txt, ov in para["runs"]:
            _run(p, txt,
                 ov.get("size", p_size), ov.get("color", p_color),
                 ov.get("bold", p_bold), ov.get("font", font),
                 ov.get("italic", italic), ov.get("spc", spc),
                 ov.get("underline", underline))
    return box


def text_in(shape, content, size=18, color=INK, bold=False, font=BODY,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
            margin=(0.18, 0.06, 0.18, 0.06), line_spacing=0.95, spc=None,
            italic=False):
    """Write text inside an existing autoshape (so the shape itself is the
    click target — tapping anywhere on the card fires the trigger)."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    l, t, r_, b = margin
    tf.margin_left, tf.margin_top = I(l), I(t)
    tf.margin_right, tf.margin_bottom = I(r_), I(b)
    tf.auto_size = None
    if isinstance(content, str):
        paras = [{"runs": [(t_, {})]} for t_ in content.split("\n")]
    else:
        paras = content
    first = True
    for para in paras:
        p = _para(tf, first, para.get("align", align),
                  para.get("line_spacing", line_spacing),
                  para.get("space_before", 0), para.get("space_after", 0))
        first = False
        for txt, ov in para["runs"]:
            _run(p, txt,
                 ov.get("size", para.get("size", size)),
                 ov.get("color", para.get("color", color)),
                 ov.get("bold", para.get("bold", bold)),
                 ov.get("font", font), ov.get("italic", italic),
                 ov.get("spc", spc), ov.get("underline", False))
    return shape


# ---------------------------------------------------------------------------
# icons drawn with geometry (no font/emoji dependency)
# ---------------------------------------------------------------------------
def check(s, cx, cy, d=0.42, ok=True, ring=True, on_dark=False):
    """Circle badge with a drawn tick / cross — renders everywhere."""
    fill = OK if ok else NO
    if ring:
        rect(s, cx - d / 2, cy - d / 2, d, d, fill=fill, prst=MSO_SHAPE.OVAL)
    w = max(1.6, d * 5.2)
    if ok:
        line(s, cx - d * 0.22, cy + d * 0.02, cx - d * 0.05, cy + d * 0.20,
             color="FFFFFF", lw=w)
        line(s, cx - d * 0.05, cy + d * 0.20, cx + d * 0.24, cy - d * 0.20,
             color="FFFFFF", lw=w)
    else:
        line(s, cx - d * 0.17, cy - d * 0.17, cx + d * 0.17, cy + d * 0.17,
             color="FFFFFF", lw=w)
        line(s, cx + d * 0.17, cy - d * 0.17, cx - d * 0.17, cy + d * 0.17,
             color="FFFFFF", lw=w)


def play_icon(s, x, y, d=0.95, fill=ACCENT, name=None):
    c = rect(s, x, y, d, d, fill=fill, prst=MSO_SHAPE.OVAL, name=name)
    tri = rect(s, x + d * 0.36, y + d * 0.30, d * 0.34, d * 0.40,
               fill="FFFFFF", prst=MSO_SHAPE.ISOSCELES_TRIANGLE)
    tri.rotation = 90
    return c


def arrow(s, x, y, w=0.42, h=0.22, color=ACCENT):
    return rect(s, x, y, w, h, fill=color, prst=MSO_SHAPE.RIGHT_ARROW)


def num_circle(s, x, y, d, n, fill=ACCENT, color="FFFFFF", size=None):
    shp = rect(s, x, y, d, d, fill=fill, prst=MSO_SHAPE.OVAL)
    text_in(shp, str(n), size=size or d * 40, color=color, bold=True,
            align=PP_ALIGN.CENTER, margin=(0, 0, 0, 0), font=HEAD)
    return shp


def image_card(s, path, x, y, w, h, pad=0.22, fill=BG_TINT, border=LINE,
               radius=0.28, shadow_on=True, fit="contain", name=None):
    """Soft card with the (transparent) clip art centred inside."""
    card = rect(s, x, y, w, h, fill=fill, line=border, lw=0.75, radius=radius,
                name=name)
    if shadow_on:
        shadow(card, blur=22, dist=5, opacity=7)
    iw, ih = Image.open(path).size
    avail_w, avail_h = w - 2 * pad, h - 2 * pad
    if fit == "contain":
        sc = min(avail_w / iw, avail_h / ih)
    else:
        sc = max(avail_w / iw, avail_h / ih)
    nw, nh = iw * sc, ih * sc
    pic = s.shapes.add_picture(path, I(x + (w - nw) / 2), I(y + (h - nh) / 2),
                              I(nw), I(nh))
    if fit == "cover":
        # crop overflow by shrinking the frame instead of cropping the picture
        pass
    return card, pic


# ---------------------------------------------------------------------------
# slide furniture
# ---------------------------------------------------------------------------
class Deck:
    def __init__(self, path, total=None):
        self.prs = Presentation()
        self.prs.slide_width = I(W)
        self.prs.slide_height = I(H)
        self.blank = self.prs.slide_layouts[6]
        self.path = path
        self.total = total
        self.slides = []
        self._ctn = {}

    # -- slides ---------------------------------------------------------
    def slide(self, dark=False):
        s = self.prs.slides.add_slide(self.blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = C(DARK if dark else BG)
        self.slides.append(s)
        self._ctn[id(s)] = 10
        return s

    # -- helpers attached to the deck for convenience -------------------
    def trigger(self, slide, clicker, targets, delay=0):
        add_trigger(slide, clicker, targets, self._next_ids(slide), delay=delay)

    def _next_ids(self, slide):
        def nxt(n=1):
            self._ctn[id(slide)] += n
            return self._ctn[id(slide)] - n + 1
        return nxt

    def progress(self, slide, index, dark=False):
        track = rect(slide, 0, 0, W, 0.065, fill=DARK_LINE if dark else LINE,
                     name="!!prog")
        fill = rect(slide, 0, 0, W * index / self.total, 0.065,
                    fill=ACCENT if not dark else "5B8CFF", name="!!progfill")
        return track, fill

    def page(self, slide, index, dark=False):
        text(slide, W - 1.55, H - 0.52, 1.0, 0.3,
             f"{index:02d} / {self.total:02d}", size=9, color=MUTED if not dark else "7E93B8",
             align=PP_ALIGN.RIGHT, font=BODY, spc=40)

    def kicker(self, slide, label, x=MX, y=0.40, color=None, dark=False):
        return text(slide, x, y, 6.0, 0.28, label.upper(), size=12,
                    color=color or (("8FB0FF") if dark else ACCENT), bold=True,
                    font=BODY, spc=180)

    def motif(self, slide, x, y, size=0.5, dark=False, opacity=40):
        """The travelling element: a small 'typing' chat bubble. Every slide
        carries the same `!!` names, so Morph glides it through the deck."""
        fill = "1B3A66" if dark else ACCENT_SOFT
        h = size * 0.74
        b = rect(slide, x, y, size, h, fill=fill, line=None,
                 radius=size * 0.26, name="!!bubble")
        for i in range(3):
            dd = size * 0.125
            rect(slide, x + size * (0.24 + i * 0.185), y + h * 0.40, dd, dd,
                 fill=("5B8CFF" if dark else ACCENT), prst=MSO_SHAPE.OVAL,
                 name=f"!!dot{i + 1}")
        return b

    # -- interaction ----------------------------------------------------
    def hyperlink(self, shape, target_slide):
        shape.click_action.target_slide = target_slide

    def morph(self, slide, duration=900, kind="byObject"):
        add_morph(slide, duration_ms=duration, morph_type=kind)

    def notes(self, slide, body):
        slide.notes_slide.notes_text_frame.text = body

    def raise_audio(self):
        """Put every play button on top of the slide's shape tree, so no
        transparent text box can swallow the click that should play the mp3."""
        for slide in self.slides:
            tree = slide.shapes._spTree
            for sh in list(slide.shapes):
                if sh.name.startswith("!!audio"):
                    tree.remove(sh._element)
                    tree.append(sh._element)

    def save(self):
        self.raise_audio()
        self.prs.save(self.path)
        _normalize_audio_rels(self.path)
        add_audio_timing(self.path)


# ---------------------------------------------------------------------------
# Morph
# ---------------------------------------------------------------------------
def add_morph(slide, duration_ms=900, morph_type="byObject", speed="med"):
    xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:p14="{P14}" spd="{speed}" p14:dur="{duration_ms}">'
        f'<p14:morph type="{morph_type}"/></p:transition>'
    )
    sld = slide._element
    insert_at = 0
    for i, child in enumerate(sld):
        if child.tag.endswith("}cSld") or child.tag.endswith("}clrMapOvr"):
            insert_at = i + 1
    sld.insert(insert_at, etree.fromstring(xml))


# ---------------------------------------------------------------------------
# Click triggers (interactive timing sequences)
# ---------------------------------------------------------------------------
TIMING_SKELETON = (
    '<p:timing {ns}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
    'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
    '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst/>'
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
    "</p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"
)


def _ensure_timing(slide):
    sld = slide._element
    timing = sld.find(qn("p:timing"))
    if timing is None:
        timing = etree.fromstring(TIMING_SKELETON.format(ns=nsdecls("p")))
        sld.insert_element_before(timing, "p:extLst")
    return timing


def _flatten(items):
    out = []
    for it in items:
        if isinstance(it, (list, tuple, set)):
            out.extend(_flatten(it))
        elif it is not None:
            out.append(it)
    return out


def add_trigger(slide, clicker, targets, ids, delay=0):
    """Clicking *clicker* reveals every shape in *targets* (in place).

    Emits the structure PowerPoint writes itself: one
    `<p:seq><p:cTn nodeType="interactiveSeq">` per trigger, holding one par per
    effect — the first `clickEffect`, the rest `withEffect` so they play
    together under a single click.
    """
    timing = _ensure_timing(slide)
    root = (timing.find(qn("p:tnLst")).find(qn("p:par")).find(qn("p:cTn"))
            .find(qn("p:childTnLst")))

    seq_id = ids(1)
    par_id = ids(1)
    effects = []
    for i, t in enumerate(_flatten(targets)):
        e_id = ids(1)
        s_id = ids(1)
        a_id = ids(1)
        node = "clickEffect" if i == 0 else "withEffect"
        effects.append(
            '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" fill="hold" '
            'grpId="0" nodeType="%s"><p:stCondLst><p:cond delay="%d"/>'
            '</p:stCondLst><p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"><p:stCondLst>'
            '<p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="%d"/>'
            "</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName>"
            "</p:attrNameLst></p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set>"
            '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="%d" dur="350"/>'
            '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr></p:animEffect>'
            "</p:childTnLst></p:cTn></p:par>"
            % (e_id, node, delay, s_id, t.shape_id, a_id, t.shape_id)
        )

    xml = (
        '<p:seq xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'concurrent="1" nextAc="seek"><p:cTn id="%d" restart="whenNotActive" fill="hold" '
        'evtFilter="cancelBubble" nodeType="interactiveSeq">'
        '<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="%d"/>'
        "</p:tgtEl></p:cond></p:stCondLst>"
        '<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>'
        '<p:childTnLst><p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="0"/>'
        "</p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par></p:childTnLst>"
        "</p:cTn></p:seq>"
        % (seq_id, clicker.shape_id, par_id, "".join(effects))
    )
    root.append(etree.fromstring(xml))


# ---------------------------------------------------------------------------
# Audio (the exact shape PowerPoint writes for an inserted mp3)
# ---------------------------------------------------------------------------
def add_audio(slide, mp3_path, x, y, d=0.95, label=None, name="!!audio"):
    """A round, clickable play button that plays the embedded mp3 in place.

    Built by python-pptx's add_movie (which wires up the media part, the
    p14:media extension and a poster frame) and then relabelled from
    `<a:videoFile>` to `<a:audioFile>` so PowerPoint treats it as sound.
    """
    from pptx.enum.shapes import MSO_SHAPE as _S
    # poster frame = a flat accent disc with a white triangle, rendered here
    poster = _play_poster()
    pic = slide.shapes.add_movie(mp3_path, I(x), I(y), I(d), I(d),
                                 poster_frame_image=poster,
                                 mime_type="audio/mpeg")
    pic.name = name
    el = pic._element
    vf = el.find(".//" + qn("a:videoFile"))
    if vf is not None:
        vf.tag = qn("a:audioFile")
        # the relationship behind <a:audioFile r:link> must be type "audio",
        # not "video" — PowerPoint is strict about this one
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        rId = vf.get(qn("r:link"))
        rel = pic.part.rels[rId]
        rel._reltype = RT.AUDIO
    return pic


def _play_poster(size=380, fill=ACCENT):
    from PIL import ImageDraw
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.ellipse([0, 0, size - 1, size - 1], fill="#0B1B33")
    d.ellipse([6, 6, size - 7, size - 7], fill="#1D4ED8")
    m = size * 0.34
    d.polygon([(m, size * 0.30), (m, size * 0.70), (size - m * 0.92, size * 0.50)],
              fill="white")
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# package fix-ups
# ---------------------------------------------------------------------------
def _normalize_audio_rels(path):
    """python-pptx wires an mp3 as a *video* relationship; PowerPoint wants
    `<a:audioFile>` + a `.../audio` relationship. Rewrite the rels parts."""
    import re
    import shutil
    import zipfile
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".rels") and "/slides/" in item.filename:
                txt = data.decode("utf-8")
                txt = re.sub(
                    r'(Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/)video(")',
                    r"\1audio\2", txt)
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ---------------------------------------------------------------------------
# Audio playback (the part python-pptx does not write)
# ---------------------------------------------------------------------------
_MEDIA_NODE = (
    '<p:audio xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cMediaNode vol="100000">'
    '<p:cTn id="{cid}" fill="hold" display="0">'
    '<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
    '<p:endCondLst><p:cond evt="onStopAudio" delay="0">'
    '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>'
    '</p:cTn>'
    '<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    '</p:cMediaNode></p:audio>'
)


def add_audio_timing(path, delay="indefinite"):
    """Give every embedded-audio picture a <p:audio><p:cMediaNode> cue.

    python-pptx writes the picture, the media part and the click action, plus a
    <p:video> cue — but a *video* cue on a shape that declares <a:audioFile>
    makes PowerPoint ignore the click (it looks like a still picture, so the
    click just advances the slide). PowerPoint's own encoding for
    "Start: In Click Sequence" is <p:audio> with delay="indefinite".
    Here every video/audio cue on a media shape is replaced by that node.
    """
    import re
    import shutil
    import zipfile
    from lxml import etree

    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    AUDIO = ("<p:audio xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
             "<p:cMediaNode vol=\"80000\">"
             "<p:cTn id=\"{cid}\" fill=\"hold\" display=\"0\">"
             "<p:stCondLst><p:cond delay=\"{delay}\"/></p:stCondLst>"
             "<p:endCondLst><p:cond evt=\"onStopAudio\" delay=\"0\">"
             "<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>"
             "</p:cTn>"
             "<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>"
             "</p:cMediaNode></p:audio>")

    def slide_media(root):
        out = []
        for pic in root.iter(P + "pic"):
            nv = pic.find("./" + P + "nvPicPr/" + P + "nvPr")
            if nv is None or nv.find(A + "audioFile") is None:
                continue
            cnv = pic.find("./" + P + "nvPicPr/" + P + "cNvPr")
            if cnv is not None and cnv.get("id"):
                out.append(cnv.get("id"))
        return out

    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", item.filename):
                root = etree.fromstring(data)
                spids = slide_media(root)
                if spids:
                    timing = root.find(P + "timing")
                    if timing is None:
                        timing = etree.SubElement(root, P + "timing")
                        etree.SubElement(timing, P + "tnLst")
                    tnLst = timing.find(P + "tnLst")
                    root_ctn = tnLst.find(P + "par/" + P + "cTn")
                    child = root_ctn.find(P + "childTnLst")
                    if child is None:
                        child = etree.SubElement(root_ctn, P + "childTnLst")
                    # drop every stale cue aimed at a media shape
                    for tag in ("video", "audio"):
                        for node in list(child.findall(P + tag)):
                            tgt = node.find(".//" + P + "spTgt")
                            if tgt is not None and tgt.get("spid") in spids:
                                child.remove(node)
                    ids = [int(c.get("id")) for c in root.iter(P + "cTn")
                           if c.get("id") and c.get("id").isdigit()]
                    nxt = (max(ids) if ids else 1) + 1
                    for spid in spids:
                        child.append(etree.fromstring(
                            AUDIO.format(cid=nxt, spid=spid, delay=delay)))
                        nxt += 1
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
            elif item.filename == "[Content_Types].xml":
                txt = data.decode("utf-8")
                if 'Extension="mp3"' not in txt:
                    txt = re.sub(r'<Override PartName="/ppt/media/[^"]+\.mp3"[^/]*/>', "", txt)
                    txt = txt.replace(
                        "</Types>",
                        '<Default Extension="mp3" ContentType="audio/mpeg"/></Types>')
                    data = txt.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
