#!/usr/bin/env python3
"""Rebuild "Demo lesson Oct 21" as an interactive PowerPoint lesson.

Content law: **not a single word of the lesson material changes.** Every
question, option, phrase, sentence frame, abbreviation, transcript line and
task instruction is reproduced verbatim (including the original spelling like
"Acronym", double spaces and the space-before-"?" style).

What changes is the design system, the layout, and the interaction layer:

  * every question is answerable *on the slide* — tap a card, the answer or the
    feedback fades in right there (native PowerPoint click triggers, no macros,
    no add-ins)
  * the audio is embedded and playable from any listening slide
  * transition = Morph on every slide after the first, with one travelling
    chat-bubble motif that glides through the whole deck
  * speaker notes on every slide tell the teacher how to run the interaction

Run:  python3 build.py ["Demo lesson Oct 21 - Interactive.pptx"]
"""
from __future__ import annotations

import math
import os
import sys

from PIL import Image, ImageFont

from deckkit import (set_alpha, MUTED_DARKBG, DARK_FOOT, FOOT, ACCENT, ACCENT_DK, ACCENT_LINE, ACCENT_SOFT, ACCENT_TEXT,
                     BG, BG_TINT, BODY, CARD, CW, DARK, DARK_2, DARK_LINE, HEAD,
                     H, INK, INK_SOFT, LINE, MUTED, MX, NO, NO_SOFT, OK, OK_SOFT,
                     W, Deck, MSO_ANCHOR, MSO_SHAPE, PP_ALIGN, add_audio, arrow,
                     image_card, line, num_circle, rect, shadow, text, text_in)

HERE = os.path.dirname(os.path.abspath(__file__))
A = lambda n: os.path.join(HERE, "assets", n)
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    HERE, "Demo lesson Oct 21 - Interactive.pptx")

TOTAL = 24

# ---------------------------------------------------------------------------
# verbatim lesson material
# ---------------------------------------------------------------------------
# verbatim from the source slide \u2014 the "1 \u2013 " prefixes included
OBJECTIVES = [
    "1 \u2013 Predict the topic of a conversation .",
    "2 \u2013 Understand the conversation .",
    "3 \u2013 Use some texting abbreviation & Acronym  in a text message .",
    "4 \u2013  Use the \u201creal talk\u201d phrases in a conversation .",
    "5 \u2013 Ask for clarifications in a conversation .",
]

WORKSHEET = ["I don\u2019t understand.", "Can you explain it?", "How do you\u2026?",
             "What does that mean?", "Does that make sense?", "I get it."]

REAL_TALK = [
    ("How on earth did that happen ?", "How is it possible"),
    ("If you can hold on , I\u2019ll go and get some help .", "Wait a moment."),
    ("Where do the youngsters hang out ?", "spend time together informally"),
    ("I don\u2019t get it \u2013 it doesn\u2019t make sense .", "understand"),
    ("I\u2019ll teach you how to use the program \u2013 you\u2019ll get the hang of it after a while.",
     "started to understand something"),
]

FRAMES = [
    "How on earth can you _____?",
    "My friends and I often hang out at ______.",
    "Hold on a minute. I have to _____.",
    "I don\u2019t get it. Why do you _____?",
    "I finally got the hang of ______.",
]

TRANSCRIPT = [
    ("Grandma", "What are you doing, Cara?"),
    ("Cara", "I\u2019m texting Maria. I haven\u2019t seen her all week, so I\u2019m just "
             "saying \u201chi.\u201d"),
    ("Grandma", "What\u2019s that beep?"),
    ("Cara", "It\u2019s Maria. She just sent a text message back to me."),
    ("Grandma", "I don\u2019t understand. How on earth do you type what you want to "
                "say so quickly?"),
    ("Cara", "You don\u2019t actually type the words out. You use abbreviations and "
             "acronyms. Hold on. I\u2019ll show you. Right now I\u2019m asking if she "
             "wants to come over and hang out."),
    ("Grandma", "That\u2019s a lot to type."),
    ("Cara", "Right. So instead I type \u201cDYWT come ovr.\u201d Does that make sense?"),
    ("Grandma", "DYWT? What does that mean?"),
    ("Cara", "DYWT means \u201cDo you want to.\u201d It\u2019s the first letter of each word."),
    ("Grandma", "Oh. OK. Now what?"),
    ("Cara", "She\u2019ll respond in a second. (beep) There."),
    ("Grandma", "What did she say?"),
    ("Cara", "(spelling it out) Gr8. B ovr n 10 mins."),
    ("Grandma", "I think I get it. She\u2019s saying, \u201cGreat. Be over in ten minutes.\u201d"),
    ("Cara", "Now you\u2019ve got the hang of it, Grandma!"),
]

LISTEN_Q = [
    ("What are they talking about ?", "Texting \u2014 Cara is texting Maria."),
    ("What is the granddaughter\u2019s name ?", "Cara."),
    ("What is Cara\u2019s friend  name ?", "Maria."),
    ("What is Cara doing ?", "She is texting Maria."),
    ("What is the sound grandma heard first ?", "A beep."),
    ("Does grandma know how to text ?", "No \u2014 Cara shows her how."),
]

TASK = [
    "What doesn\u2019t the grandmother understand?",
    "What acronym does the grandmother learn? What texting acronyms do you use?",
    "Tell your partner about a time you explained something or learned something new "
    "from a friend.",
]

TF = [
    ("1- The grandmother doesn\u2019t understand how Cara types a text message so quickly .",
     "\u201cHow on earth do you type what you want to say so quickly?\u201d"),
    ("2- The grandmother learns the acronym, \u201cDYWT,\u201d or \u201cdo you want to.\u201d",
     "\u201cDYWT means \u2018Do you want to.\u2019 It\u2019s the first letter of each word.\u201d"),
]

QUIZ = {
    20: dict(lines=[("Hi Ali.  Let\u2019s ...... in the mall.", "in"),
                   ("Hi Ahmed. OK. I\u2019ll call Faris to meet us there.", "out")],
             options=["hold on", "Hang out", "How on earth?"], correct=1,
             art="image25.jpeg",
             why="\u201cLet\u2019s hang out in the mall\u201d = let\u2019s spend time "
                 "together there."),
    21: dict(lines=[("He will send you his place with google map.", "in"),
                   ("Yes, I ...............", "out")],
             options=["get it", "getting the long of it", "hold on"], correct=0,
             art="image24.jpeg",
             why="\u201cYes, I get it\u201d = yes, I understand."),
    22: dict(lines=[("............. google maps to me, please?", "in"),
                   ("Sure", "out")],
             options=["Does that make sense..", "How on earth ..", "Can you explain.."],
             correct=2, art="image20.png",
             why="You are asking someone to explain \u2014 so: \u201cCan you "
                 "explain\u2026\u201d"),
}

REORDER = [("How do you add a picture to Photoshop?", 1),
           ("Just drag & drop from your files first", 2),
           ("Great", 3),
           ("Ahaa, I get it now", 4)]

GRID = {
    16: dict(kicker="Clarify",
             q="Which things would you say when you don\u2019t understand something "
               "or want an explanation ?",
             marks=["ok", "ok", "ok", "ok", "no", "no"],
             note="Tap a phrase: green = it asks for clarification, red = it doesn\u2019t."),
    17: dict(kicker="Confirm",
             q="What do you say when you understand something ?",
             marks=["no", "no", "no", "no", "no", "ok"],
             note="Tap a phrase: only one of these confirms that you understand."),
    18: dict(kicker="Check",
             q="What do you say when you want to know if your explanation is clear ?",
             marks=["no", "no", "no", "no", "ok", "no"],
             note="Tap a phrase: only one of these checks that you were clear."),
}

# ---------------------------------------------------------------------------
# measuring helpers (so nothing overflows the slide)
# ---------------------------------------------------------------------------
_FONTS = {}


def _f(size):
    if size not in _FONTS:
        _FONTS[size] = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            max(8, int(round(size * 96 / 72))))
    return _FONTS[size]


def n_lines(txt, size, width_in, factor=1.0):
    """Approximate wrapped line count at `size` pt inside `width_in` inches."""
    f = _f(size)
    limit = width_in * 96 * 0.97 * factor
    words, lines, cur = txt.split(), 1, ""
    for w in words:
        trial = (cur + " " + w).strip()
        if f.getlength(trial) > limit and cur:
            lines += 1
            cur = w
        else:
            cur = trial
    return lines


def fit_size(rows, width_in, height_in, sizes, line_factor=1.22, gap_in=0.0):
    for size in sizes:
        h = sum(n_lines(r, size, width_in) * size * line_factor / 72.0 for r in rows)
        h += gap_in * (len(rows) - 1)
        if h <= height_in:
            return size
    return sizes[-1]


# ---------------------------------------------------------------------------
# components
# ---------------------------------------------------------------------------
def progress(s, idx, dark=False):
    """Progress drawn as one rectangle whose width is the fraction done, so
    Morph simply stretches it from slide to slide."""
    rect(s, 0, 0, W, 0.055, fill=(DARK_LINE if dark else LINE), name="!!track")
    if idx > 1:
        rect(s, 0, 0, W * idx / TOTAL, 0.055,
             fill=("5B8CFF" if dark else ACCENT), name="!!prog")


def page(s, idx, dark=False):
    """Deck progress: a hairline rule with the slide number at its right end."""
    c = DARK_FOOT if dark else FOOT
    rect(s, MX, H - 0.52, 0.44, 0.028, fill=("5B8CFF" if dark else ACCENT))
    text(s, MX + 0.56, H - 0.60, 3.4, 0.28, "Connected by technology", size=12,
         color=c, spc=40)
    text(s, W - 1.15, H - 0.60, 1.1, 0.28, f"{idx:02d} / {TOTAL:02d}", size=12,
         color=c, align=PP_ALIGN.RIGHT, spc=40)


MOTIF_SPOTS = [(W - 2.72, H - 0.74)] * 4


def motif(d, s, idx, dark=False):
    """The lesson\u2019s texting mark, parked in the footer on every slide so it
    reads as a signature rather than a stray sticker."""
    x, y = MOTIF_SPOTS[0]
    d.motif(s, x, y, 0.40, dark=dark)


def brand(s, dark=False):
    """The school-book branding lockup (Ministry of Education, Vision 2030,
    Madrasati, NCVT) exactly as the original deck stamped it on its slides."""
    strip = A("logos_strip.png")
    iw, ih = Image.open(strip).size
    h = 0.30
    w = h * iw / ih
    x = W - MX - w
    if dark:
        rect(s, x - 0.16, 0.13, w + 0.32, h + 0.18, fill="FFFFFF",
             line=None, radius=(h + 0.18) / 2)
    else:
        rect(s, x - 0.16, 0.13, w + 0.32, h + 0.18, fill="FFFFFF",
             line="E4EBF6", lw=0.8, radius=(h + 0.18) / 2, opacity=64)
    s.shapes.add_picture(strip, int(x * 914400), int(0.21 * 914400),
                         int(w * 914400), int(h * 914400))


def head(d, s, idx, kicker, title=None, size=32, dark=False, tw=CW, ty=0.86,
         sub=None):
    progress(s, idx, dark)
    page(s, idx, dark)
    brand(s, dark)
    d.kicker(s, kicker, dark=dark)
    if title:
        text(s, MX, ty, tw, 1.5, title, size=size, bold=True, font=HEAD,
             color=("FFFFFF" if dark else INK), line_spacing=0.92)
    if sub:
        text(s, MX, ty + 1.02, tw, 0.4, sub, size=14,
             color=("9FB3D4" if dark else MUTED))


def card(s, x, y, w, h, fill=CARD, border=LINE, radius=0.2, sh=True, lw=0.9):
    c = rect(s, x, y, w, h, fill=fill, line=border, lw=lw, radius=radius)
    if sh:
        shadow(c, blur=20, dist=4, opacity=8)
    return c


def button(s, x, y, w, h, label, fill=ACCENT, color="FFFFFF", size=14,
           outline=False, name=None):
    b = rect(s, x, y, w, h, fill=(CARD if outline else fill),
             line=(ACCENT_LINE if outline else None), lw=1.2, radius=h / 2,
             name=name)
    text_in(b, label, size=size, color=(ACCENT_TEXT if outline else color),
            bold=True, font=BODY, align=PP_ALIGN.CENTER,
            margin=(0.1, 0.02, 0.1, 0.02))
    if outline:
        shadow(b, blur=16, dist=3, opacity=7)
    return b


def badge(s, cx, cy, d=0.4, kind="ok"):
    """Feedback disc: green tick, red cross or neutral dash. Returns the list of
    shapes it is made of so a trigger can hide/reveal all of them together."""
    if kind == "ok":
        disc = rect(s, cx - d / 2, cy - d / 2, d, d, fill=OK, prst=MSO_SHAPE.OVAL)
        w = max(1.5, d * 5.0)
        l1 = line(s, cx - d * 0.22, cy + d * 0.01, cx - d * 0.05,
                  cy + d * 0.19, "FFFFFF", w)
        l2 = line(s, cx - d * 0.05, cy + d * 0.19, cx + d * 0.24,
                  cy - d * 0.20, "FFFFFF", w)
    elif kind == "no":
        disc = rect(s, cx - d / 2, cy - d / 2, d, d, fill=NO, prst=MSO_SHAPE.OVAL)
        w = max(1.5, d * 5.0)
        l1 = line(s, cx - d * 0.17, cy - d * 0.17, cx + d * 0.17,
                  cy + d * 0.17, "FFFFFF", w)
        l2 = line(s, cx + d * 0.17, cy - d * 0.17, cx - d * 0.17,
                  cy + d * 0.17, "FFFFFF", w)
    else:
        disc = rect(s, cx - d / 2, cy - d / 2, d, d, fill="E6EBF4",
                    prst=MSO_SHAPE.OVAL)
        l1 = line(s, cx - d * 0.18, cy, cx + d * 0.18, cy, MUTED,
                  max(1.5, d * 4.2))
        l2 = None
    return [disc, l1, l2]


def bubble(s, x, y, w, h, body, size=20, fill=ACCENT_SOFT, line_c=ACCENT_LINE,
           color=INK, bold=False, align=PP_ALIGN.LEFT, speaker=None,
           margin=None, radius=0.3, lw=1.1):
    b = rect(s, x, y, w, h, fill=fill, line=line_c, lw=lw,
             prst=MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT)
    try:
        b.adjustments[0] = 0.22
        b.adjustments[1] = radius
        b.adjustments[2] = -0.42
    except Exception:
        pass
    shadow(b, blur=16, dist=4, opacity=7)
    if speaker:
        text(s, x + 0.26, y + 0.14, 1.6, 0.26, speaker.upper(), size=12,
             color=ACCENT_TEXT, bold=True, spc=120)
        text_in(b, body, size=size, color=color, bold=bold, align=align,
                margin=(0.26, 0.42, 0.30, 0.22))
    else:
        text_in(b, body, size=size, color=color, bold=bold, align=align,
                margin=margin or (0.30, 0.16, 0.34, 0.16))
    return b


def backdrop(s, dark=False, variant=0):
    """Soft geometry behind the content \u2014 the same motif, quiet."""
    if dark:
        rect(s, 8.4 + variant * 0.4, -2.0, 7.2, 7.2, fill="14284A",
             prst=MSO_SHAPE.OVAL, opacity=70, name="!!bg1")
        rect(s, -1.8, 4.8, 4.4, 4.4, fill="14284A", prst=MSO_SHAPE.OVAL,
             opacity=55, name="!!bg2")
    else:
        rect(s, 10.4 + variant * 0.3, -1.9, 5.0, 5.0, fill=BG_TINT,
             prst=MSO_SHAPE.OVAL, name="!!bg1")
        rect(s, -1.5, 5.2, 3.5, 3.5, fill="F4F8FF", prst=MSO_SHAPE.OVAL,
             name="!!bg2")


def place(s, name, x, y, w, h, align="center", flip=False, pad=0.0):
    """Place a trimmed asset inside the box (x, y, w, h) keeping its aspect."""
    path = A("t/" + name)
    iw, ih = Image.open(path).size
    bw, bh = w - 2 * pad, h - 2 * pad
    sc = min(bw / iw, bh / ih)
    nw, nh = iw * sc, ih * sc
    px = x + pad + (bw - nw) / 2 if align == "center" else x + pad
    py = y + pad + (bh - nh) / 2
    pic = s.shapes.add_picture(path, int(px * 914400), int(py * 914400),
                               int(nw * 914400), int(nh * 914400))
    if flip:
        flip_h(pic)
    return pic


def flip_h(pic):
    pic._element.spPr.xfrm.set("flipH", "1")
    return pic


def hint(s, x, y, w, body, size=12.5):
    text(s, x, y, w, 0.5, body, size=size, color=MUTED, italic=True)


# ---------------------------------------------------------------------------
# slides
# ---------------------------------------------------------------------------
def s01_title(d):
    s = d.slide(dark=True)
    backdrop(s, dark=True)
    rect(s, 0.95, 1.24, 0.88, 0.075, fill="5B8CFF")
    text(s, 0.95, 1.48, 7.4, 0.32, "CONVERSATION", size=12, color="8FB0FF",
         bold=True, spc=260, font=BODY)
    text(s, 0.95, 1.92, 7.85, 2.35,
         [{"runs": [("Connected by", {})], "line_spacing": 0.92},
          {"runs": [("Technology", {})], "line_spacing": 0.92},
          {"runs": [("Conversation", {"color": "8FB0FF"})], "line_spacing": 0.92}],
         size=43, bold=True, font=HEAD, color="FFFFFF")

    text(s, 0.95, 4.34, 7.2, 0.32,
         "\u25cf  Interactive \u2014 tap the cards as you go", size=12.5,
         color=MUTED_DARKBG)
    b = button(s, 0.95, 4.86, 3.15, 0.66, "Start the lesson", size=15,
               name="btn_start")

    c = rect(s, 8.42, 1.12, 4.34, 5.26, fill=DARK_2, line=DARK_LINE, lw=1.1,
             radius=0.36)
    shadow(c, blur=40, dist=10, opacity=35, color="000000")
    place(s, "image3.png", 8.68, 1.34, 3.85, 3.55)
    place(s, "image14.png", 8.66, 5.42, 1.20, 0.74, align="left")
    place(s, "image13.png", 10.02, 5.26, 1.45, 1.01, align="left")

    text(s, 0.95, 6.55, 9.6, 0.3,
         "Every question in this lesson answers when you tap it.",
         size=12.5, color=DARK_FOOT)
    brand(s, dark=True)
    d.notes(s, "OPENING \u2014 1 min.\n"
               "Read the title together. The two stickers preview the lesson language "
               "(texting abbreviations).\n"
               "Tap \u201cStart the lesson\u201d (or press \u2192) to move on.\n\n"
               "HOW THIS DECK WORKS: nothing is pre-revealed. In the slideshow, answers "
               "and feedback are invisible until you tap the matching card \u2014 then they "
               "fade in on the same slide. In edit view you will see them all; that is "
               "normal.\n"
               "Present in PowerPoint 2019+ / Microsoft 365 so Morph and the click "
               "triggers work. Exporting to video plays everything back exactly as it "
               "behaves live.")
    return s


def s02_objectives(d):
    s = d.slide()
    backdrop(s, variant=3)
    head(d, s, 2, "Objectives", "Lesson Objectives", size=34, tw=8.3)
    text(s, MX, 1.56, 8.3, 0.34, "The students will be able to :", size=16,
         color=MUTED)
    y = 2.02
    for i, obj in enumerate(OBJECTIVES, 1):
        c = card(s, MX, y, 8.35, 0.82, radius=0.16)
        text_in(c, obj, size=17, color=INK, margin=(0.30, 0.06, 0.85, 0.06))
        tick = badge(s, MX + 7.65, y + 0.425, 0.42, "ok")
        d.trigger(s, c, [tick])
        y += 0.94

    card(s, 9.28, 2.02, 3.4, 2.86, fill=BG_TINT, radius=0.26, sh=False)
    place(s, "image2.png", 9.42, 2.14, 3.12, 2.62)
    text(s, 9.10, 5.22, 3.75, 0.6,
         "Tick a row once the class has covered it.", size=12, color=MUTED,
         align=PP_ALIGN.CENTER)
    d.notes(s, "OBJECTIVES \u2014 1 min.\n"
               "Read the five aims. Tap each row as the lesson reaches it \u2014 a green "
               "tick fades in. Five empty slots give the class a visible sense of progress "
               "through the lesson.")
    return s


def s03_warmup(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 3, "Warm-up", "Is texting a common means of communication ?",
         size=31, tw=7.0)

    yes = bubble(s, MX, 3.00, 2.55, 0.95, "YES", size=20, bold=True,
                 align=PP_ALIGN.CENTER, fill=CARD, line_c="BFE3CB")
    no = bubble(s, 3.40, 3.00, 2.55, 0.95, "NO", size=20, bold=True,
                align=PP_ALIGN.CENTER, fill=CARD, line_c="F3C9C4")
    ty = badge(s, MX + 2.20, 3.475, 0.34, "ok")
    tn = badge(s, 3.40 + 2.20, 3.475, 0.34, "ok")
    d.trigger(s, yes, [ty])
    d.trigger(s, no, [tn])

    show = button(s, MX, 4.40, 3.35, 0.60, "Show examples", outline=True)
    hint(s, MX, 5.15, 4.9, "Vote first \u2014 hands up for YES, hands up for NO.")

    ex = rect(s, 7.95, 1.26, 4.75, 3.62, fill=CARD, line=LINE, lw=0.9,
              radius=0.28)
    shadow(ex, blur=24, dist=6, opacity=9)
    place(s, "image4.png", 8.08, 1.38, 4.49, 3.20)
    cap = text(s, 7.95, 3.66, 4.75, 0.34, "texting abbreviations", size=12.5,
               color=ACCENT_TEXT, bold=True, align=PP_ALIGN.CENTER)
    d.trigger(s, show, [ex, cap])

    card(s, 9.95, 5.02, 2.55, 1.55, fill=BG_TINT, radius=0.22, sh=False)
    place(s, "image3.png", 10.05, 5.10, 2.35, 1.40)
    d.notes(s, "WARM-UP \u2014 4 min.\n"
               "Ask the question and take a quick class vote: tap YES or NO \u2014 a tick "
               "marks the option that won.\n"
               "Then tap \u201cShow examples\u201d: the acronym cloud from the original "
               "slide appears. The class reads a few aloud. Ask: do you use any of these "
               "when you text in Arabic?")
    return s


def s04_what(d):
    s = d.slide()
    backdrop(s, variant=2)
    head(d, s, 4, "Look", "WHAT ?", size=44, tw=3.2, ty=0.94)

    c = rect(s, 3.92, 1.28, 6.10, 4.80, fill=CARD, line=LINE, lw=0.9, radius=0.32)
    shadow(c, blur=26, dist=6, opacity=9)
    place(s, "image4.png", 4.05, 1.42, 5.84, 4.30)
    place(s, "image6.png", 9.55, 4.62, 1.55, 1.60)

    word1 = rect(s, MX, 2.52, 3.10, 0.76, fill=ACCENT_SOFT, line=ACCENT_LINE,
                 lw=1.2, radius=0.16)
    text_in(word1, "texting", size=25, color=ACCENT_TEXT, bold=True, font=HEAD,
            align=PP_ALIGN.CENTER, margin=(0.1, 0.04, 0.1, 0.04))
    word2 = rect(s, MX, 3.44, 3.10, 0.76, fill=ACCENT_SOFT, line=ACCENT_LINE,
                 lw=1.2, radius=0.16)
    text_in(word2, "abbreviation", size=22, color=ACCENT_TEXT, bold=True,
            font=HEAD, align=PP_ALIGN.CENTER, margin=(0.1, 0.04, 0.1, 0.04))
    reveal = button(s, MX, 4.48, 3.10, 0.58, "Reveal the words", outline=True,
                    size=13)
    d.trigger(s, reveal, [word1, word2])
    hint(s, MX, 5.24, 3.15, "What can you see?\nSay the words before you tap.")
    d.notes(s, "LOOK \u2014 2 min.\n"
               "Ask what the cloud shows. Elicit \u201ctexting\u201d and "
               "\u201cabbreviation\u201d first; then tap \u201cReveal the words\u201d to "
               "confirm both labels.")
    return s


def s05_language(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 5, "Think", "How do people text in your language?", size=31, tw=9.6)

    eg = rect(s, MX, 2.35, 2.60, 1.05, fill=ACCENT_SOFT, line=ACCENT_LINE,
              lw=1.4, radius=0.22, name="btn_eg")
    text_in(eg, "e.g.", size=30, color=ACCENT_TEXT, bold=True, font=HEAD,
            align=PP_ALIGN.CENTER, margin=(0.1, 0.05, 0.1, 0.05))
    shadow(eg, blur=18, dist=4, opacity=8)

    cards = []
    for i, (term, mean) in enumerate([("PLZ", "please"), ("Gr8", "great"),
                                      ("B4", "before")]):
        x = MX + i * 3.92
        c = rect(s, x, 4.22, 3.62, 0.98, fill=CARD, line=LINE, lw=0.9, radius=0.2)
        shadow(c, blur=18, dist=4, opacity=8)
        text_in(c, [{"runs": [(term + "   =   ", {"color": ACCENT, "bold": True}),
                              (mean, {"color": INK})]}],
                size=20, margin=(0.3, 0.08, 0.2, 0.08))
        cards.append(c)
    d.trigger(s, eg, cards)

    card(s, 9.80, 1.28, 2.90, 2.36, fill=BG_TINT, radius=0.24, sh=False)
    place(s, "image3.png", 9.95, 1.40, 2.60, 2.12)
    hint(s, MX, 5.55, 8.8,
         "How do you shorten words when you text in Arabic? Say two examples, then tap "
         "e.g. for three English ones.", size=13)
    d.notes(s, "THINK \u2014 3 min.\n"
               "Pair talk first: how do people shorten words when texting in Arabic? Take "
               "two examples from the class. Then tap \u201ce.g.\u201d to reveal three "
               "English text abbreviations the lesson reuses later (PLZ, Gr8, B4).")
    return s


def s06_predict(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 6, "Listening", "What do you think they are talking about ?",
         size=30, tw=7.2)

    add_audio(s, A("media1.mp3"), MX, 2.00, 0.80, name="!!audio")
    text(s, 1.58, 2.18, 4.4, 0.5, "Listen to the audio", size=15, bold=True,
         color=INK, font=HEAD)

    g1 = bubble(s, MX, 2.96, 2.90, 1.05, "football?", size=24,
                align=PP_ALIGN.CENTER, fill=CARD, line_c=ACCENT_LINE,
                margin=(0.3, 0.16, 0.75, 0.16))
    g2 = bubble(s, 3.80, 2.96, 2.50, 1.05, "\u2026.?", size=24,
                align=PP_ALIGN.CENTER, fill=CARD, line_c=LINE,
                margin=(0.3, 0.16, 0.75, 0.16))
    t1 = badge(s, MX + 2.52, 3.49, 0.34, "ok")
    t2 = badge(s, 3.80 + 2.14, 3.49, 0.34, "ok")
    d.trigger(s, g1, [t1])
    d.trigger(s, g2, [t2])

    chk = button(s, MX, 4.40, 2.30, 0.56, "Check", outline=True, size=13)
    ans = rect(s, MX, 5.16, 6.50, 0.80, fill=OK_SOFT, line="BFE3CB", lw=1.0,
               radius=0.18)
    text(s, MX + 0.26, 5.34, 5.9, 0.45,
         [{"runs": [("Cara is texting Maria.", {"bold": True, "color": OK})]}],
         size=18)
    d.trigger(s, chk, [ans])
    hint(s, MX, 6.18, 6.6, "Guess first, then tap Check.")

    card(s, 8.30, 1.28, 4.42, 5.00, fill=BG_TINT, radius=0.28, sh=False)
    place(s, "image8.png", 8.55, 1.42, 3.92, 4.28)
    text(s, 8.30, 5.78, 4.42, 0.34, "Grandma  \u00b7  Cara", size=12,
         color=MUTED, align=PP_ALIGN.CENTER)
    d.notes(s, "LISTENING 1 \u2014 prediction (4 min).\n"
               "Tap the round play button to hear the conversation. Before the second "
               "listen, ask the class to guess: tap \u201cfootball?\u201d or "
               "\u201c\u2026.?\u201d \u2014 a tick records which guess the class made.\n"
               "Then tap Check: the real topic appears under the question.")
    return s


def s07_questions(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 7, "Listening", "Listen to the audio", size=26, tw=7.6, ty=0.92)
    add_audio(s, A("media1.mp3"), 11.95, 0.58, 0.58, name="!!audio")

    reveals = []
    for i, (q, a) in enumerate(LISTEN_Q):
        col, row = i % 2, i // 2
        x = MX + col * 6.16
        y = 1.62 + row * 1.42
        c = card(s, x, y, 5.96, 1.28, radius=0.18)
        text_in(c, q, size=16.5, color=INK, margin=(0.24, 0.10, 0.46, 0.10))
        strip = rect(s, x + 0.14, y + 0.80, 5.66, 0.38, fill=OK_SOFT,
                     line=None, radius=0.12)
        text_in(strip, a, size=12.5, color=OK, bold=True,
                margin=(0.14, 0.02, 0.14, 0.02))
        tick = badge(s, x + 5.63, y + 0.32, 0.33, "ok")
        d.trigger(s, c, [strip, tick])
        reveals += [strip, tick]

    allb = button(s, 9.75, 6.22, 2.60, 0.50, "Show all answers", outline=True,
                  size=12)
    d.trigger(s, allb, reveals)
    hint(s, MX, 6.33, 8.6, "Tap a question to check the answer.", size=12)
    d.notes(s, "LISTENING 2 \u2014 comprehension (6 min).\n"
               "Play the audio again (small play button, top right). Pairs answer, then tap "
               "each question card: the answer fades in inside the card that asked it, so "
               "the class checks itself.\n"
               "\u201cShow all answers\u201d reveals the whole grid for a final review.")
    return s


def s08_read(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 8, "Read", "Listen \u2018N\u2019 read", size=28, tw=5.0, ty=0.94)
    text(s, 4.60, 1.06, 2.4, 0.3, "Full script", size=12, color=MUTED)
    add_audio(s, A("media1.mp3"), 6.95, 0.72, 0.58, name="!!audio")

    card(s, MX, 1.52, 8.62, 5.22, radius=0.22)
    rows = [w + ":  " + l for w, l in TRANSCRIPT]
    size = fit_size(rows, 7.95, 4.60, [13.5, 13, 12.5, 12], line_factor=1.24,
                    gap_in=0.072)
    y = 1.74
    for who, ln in TRANSCRIPT:
        nl = n_lines(who + ":  " + ln, size, 7.95)
        h = nl * size * 1.24 / 72.0
        text(s, MX + 0.28, y, 7.95, h + 0.05,
             [{"runs": [(who + ":  ", {"bold": True, "color": ACCENT_TEXT}),
                        (ln, {"color": INK})], "line_spacing": 0.94}], size=size)
        y += h + 0.072

    card(s, 9.48, 1.52, 3.24, 5.22, fill=BG_TINT, radius=0.22, sh=False)
    text(s, 9.72, 1.72, 2.8, 0.28, "NOTICE THESE", size=12, color=ACCENT,
         bold=True, spc=160)
    for i, ph in enumerate(["How on earth", "Hold on", "hang out", "get it",
                            "got the hang of it"]):
        pill = rect(s, 9.72, 2.08 + i * 0.52, 2.76, 0.42, fill=CARD,
                    line=ACCENT_LINE, lw=1.0, radius=0.21)
        text_in(pill, ph, size=12, color=ACCENT_TEXT, bold=True,
                align=PP_ALIGN.CENTER, margin=(0.05, 0, 0.05, 0))
    text(s, 9.72, 4.85, 2.8, 1.3,
         "These are the phrases we practise all lesson \u2014 spot them as you read.",
         size=12, color=MUTED)
    d.notes(s, "READ \u2014 5 min.\n"
               "Play the audio once more while students follow the script. Then read it in "
               "two halves: teacher = Grandma, class = Cara (then swap).\n"
               "The rail on the right lists the five phrases recycled later; ask students "
               "to wave when they hear one.")
    return s


def s09_task(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 9, "About the Conversation", "About the Conversation", size=30,
         tw=9.0)
    text(s, MX, 1.58, 7.2, 0.36, "In pairs, ask and answer the questions. "
                                 "Then switch roles.", size=15, color=MUTED)

    for i, q in enumerate(TASK, 1):
        y = 2.12 + (i - 1) * 1.42
        c = card(s, MX, y, 7.55, 1.26, radius=0.18)
        text_in(c, q, size=16, color=INK, margin=(1.02, 0.10, 0.70, 0.10))
        num = rect(s, MX + 0.26, y + 0.40, 0.46, 0.46, fill=ACCENT_SOFT,
                   line=None, prst=MSO_SHAPE.OVAL)
        text_in(num, str(i), size=15, color=ACCENT_TEXT, bold=True,
                align=PP_ALIGN.CENTER, margin=(0, 0, 0, 0))
        ring = rect(s, MX + 0.21, y + 0.35, 0.56, 0.56, fill=None, line=OK,
                    lw=2.2, prst=MSO_SHAPE.OVAL)
        tick = badge(s, MX + 0.49, y + 0.63, 0.34, "ok")
        d.trigger(s, c, [ring, tick])

    card(s, 8.40, 2.12, 4.25, 3.92, fill=BG_TINT, radius=0.24, sh=False)
    place(s, "image8.png", 8.95, 2.24, 3.15, 2.86)
    text(s, 8.55, 5.14, 3.95, 0.8,
         "Switch roles after every question so both partners ask and answer.",
         size=12, color=MUTED, align=PP_ALIGN.CENTER)
    hint(s, MX, 6.52, 7.4, "Tap a question once the pair has answered it.", size=12)
    d.notes(s, "TASK CARD \u2014 pair work (6 min).\n"
               "Same instructions as the original slide. Pairs ask and answer; tap a "
               "question when the pair is done to mark it.\n"
               "Q2 has two parts: the acronym grandma learns (DYWT) and the acronyms "
               "students use themselves. Q3 opens into a short story \u2014 take one or "
               "two in open class.")
    return s


def s10_roleplay(d):
    s = d.slide(dark=True)
    backdrop(s, dark=True)
    progress(s, 10, True)
    page(s, 10, True)
    brand(s, True)
    d.kicker(s, "Speaking", dark=True)
    motif(d, s, 10, True)
    text(s, MX, 0.80, 5.6, 0.8, "Role-Play Time", size=38, bold=True, font=HEAD,
         color="FFFFFF")
    text(s, MX, 1.54, 8.4, 0.4, "Two roles, one script. Read it, then swap.",
         size=13.5, color=MUTED_DARKBG)

    for i, role in enumerate(["Grandma", "Cara"]):
        x = 0.95 + i * 5.95
        c = rect(s, x, 1.86, 5.55, 4.56, fill=DARK_2, line=DARK_LINE, lw=1.1,
                 radius=0.26)
        shadow(c, blur=28, dist=7, opacity=26, color="000000")
        lbl = rect(s, x + 0.24, 1.66, 1.72, 0.44, fill=ACCENT, radius=0.22)
        text_in(lbl, role, size=13.5, color="FFFFFF", bold=True,
                align=PP_ALIGN.CENTER, margin=(0.05, 0, 0.05, 0))

        card(s, x + 0.20, 2.26, 1.22, 2.94, fill="14284A", radius=0.16, sh=False)
        place(s, "image8.png", x + 0.24, 2.31, 1.14, 2.84,
              flip=(role == "Cara"))

        lines = [ln for w, ln in TRANSCRIPT if w == role]
        size = 12
        yy = 1.98
        for ln in lines:
            nl = n_lines(ln, size, 3.98)
            h = nl * size * 1.26 / 72.0
            t = text(s, x + 1.52, yy, 3.98, h + 0.04, ln, size=size,
                     color="E2EBF9", line_spacing=0.98)
            d.trigger(s, c, [t])
            yy += h + 0.026

    text(s, MX, 6.52, 11.5, 0.3,
         "Tap a card to reveal that role\u2019s lines \u2014 each half of the class reads its own part.",
         size=12, color=MUTED_DARKBG)
    d.notes(s, "ROLE-PLAY \u2014 5 min.\n"
               "Split the class in half (or into pairs). Tap a character card: only that "
               "role\u2019s lines appear, so each half knows exactly what to read.\n"
               "Run the whole script once, then tap the other card and swap roles. Keep the "
               "audio as a model: play a line, the class repeats it.")
    return s


def s11_flash(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 11, "Vocabulary", "Language Builder", size=31, tw=8.6, ty=0.80)
    hint(s, MX, 1.46, 8.6, "Tap a word when the class has guessed its meaning.",
         size=12.5)

    cols = [("abbreviation", "a short form of a word",
             [("PLZ", "please"), ("ovr", "over."), ("Gr8", "great"), ("B4", "before")],
             0.62),
            ("acronym", "first letters of several words to write a phrase.",
             [("OMG", "Ohh my God"), ("DYWT", "Do you want to."),
              ("WU?", "What\u2019s up?"), ("B4N", "bye for now")], 5.20)]
    for word, gloss, items, x in cols:
        text(s, x, 1.92, 4.45, 0.40, word, size=21, bold=True, font=HEAD,
             color=ACCENT_TEXT)
        text(s, x, 2.34, 4.40, 0.34, gloss, size=12, color=MUTED, italic=True)
        for i, (term, meaning) in enumerate(items):
            y = 2.84 + i * 0.90
            c = rect(s, x, y, 4.40, 0.78, fill=CARD, line=LINE, lw=0.9, radius=0.18)
            shadow(c, blur=16, dist=3, opacity=7)
            text(s, x + 0.24, y + 0.20, 1.34, 0.40, term, size=18, bold=True,
                 color=ACCENT, font=HEAD)
            text(s, x + 1.36, y + 0.21, 0.30, 0.40, "=", size=15,
                 color="9AA9BF", bold=True)
            m = text(s, x + 1.68, y + 0.16, 2.44, 0.50, meaning, size=12.5,
                     color=INK, line_spacing=0.96)
            d.trigger(s, c, [m])

    # the source deck's "Real Talk" reference box, kept as a permanent rail
    panel = rect(s, 10.02, 1.84, 2.70, 4.72, fill="FBEFC6", line="E8D7A2",
                 lw=1.0, radius=0.14)
    shadow(panel, blur=24, dist=6, opacity=10)
    tab = rect(s, 10.02, 1.84, 1.72, 0.40, fill=ACCENT, line=None, radius=0.10)
    text_in(tab, "Real Talk", size=12, color="FFFFFF", bold=True, font=HEAD,
            margin=(0.12, 0.02, 0.06, 0.02))
    text(s, 10.20, 2.36, 2.35, 0.26, "Real Talk (phrases)", size=12,
         color="6E5E2E", italic=True)
    rows = [("How on earth?", "How is it possible"),
            ("Hold on.", "Wait a moment."),
            ("hang out", "spend time together informally"),
            ("get it", "understand"),
            ("got the hang of it", "started to understand something")]
    yy = 2.66
    for term, mean in rows:
        line_txt = term + " = " + mean
        nl = n_lines(line_txt, 12, 2.22)
        h = nl * 12 * 1.34 / 72.0
        text(s, 10.20, yy, 2.24, h + 0.02,
             [{"runs": [(term + " = ", {"bold": True, "color": "8A4B12"}),
                        (mean, {"color": "3A3320"})], "line_spacing": 0.98}],
             size=12)
        yy += h + 0.13
    d.notes(s, "VOCABULARY \u2014 6 min.\n"
               "Two word families: abbreviations (letters left out) and acronyms (the "
               "first letters of several words).\n"
               "Elicit the meaning first, then tap the card: it appears to the right of the "
               "word. The cream box on the right is the original slide\u2019s Real Talk "
               "reference \u2014 keep it visible, the class will need it all lesson.\n"
               "Finish by asking for one more example of each from the class\u2019s own "
               "texting.")
    return s


def s12_find(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 12, "Read again", "Find the phrases in the script", size=27, tw=8.0,
         ty=0.94)
    add_audio(s, A("media1.mp3"), 7.10, 0.62, 0.56, name="!!audio")

    card(s, MX, 1.46, 7.35, 5.42, radius=0.22)
    size = 12
    lh = size * 1.30 / 72.0
    ys, y = [], 1.62
    for i, (who, ln) in enumerate(TRANSCRIPT):
        nl = n_lines(who + ":  " + ln, size, 6.78)
        ys.append((y, nl))
        y += nl * lh + 0.050

    phrase_at = [("How on earth", 4), ("Hold on", 5), ("hang out", 5),
                 ("get it", 14), ("got the hang of it", 15)]
    bands, seen = {}, set()
    for name, idx in phrase_at:
        if idx in seen:
            continue
        seen.add(idx)
        yy, nl = ys[idx]
        bands[idx] = rect(s, MX + 0.11, yy - 0.012, 7.13, nl * lh + 0.024,
                          fill="E8EFFF", line=None, radius=0.06)

    for i, (who, ln) in enumerate(TRANSCRIPT):
        yy, nl = ys[i]
        text(s, MX + 0.26, yy, 6.82, nl * lh + 0.04,
             [{"runs": [(who + ":  ", {"bold": True, "color": ACCENT_TEXT}),
                        (ln, {"color": INK})], "line_spacing": 0.96}], size=size)

    allbadges = []
    for j, (name, idx) in enumerate(phrase_at):
        y2 = 1.62 + j * 0.74
        ch = rect(s, 8.32, y2, 4.40, 0.58, fill=CARD, line=ACCENT_LINE, lw=1.1,
                  radius=0.16)
        shadow(ch, blur=14, dist=3, opacity=7)
        text_in(ch, name, size=15, color=ACCENT_TEXT, bold=True,
                margin=(0.28, 0.04, 0.85, 0.04))
        tick = badge(s, 12.30, y2 + 0.29, 0.28, "ok")
        d.trigger(s, ch, [bands[idx], tick])
        allbadges += tick

    allb = button(s, 8.32, 5.58, 3.0, 0.50, "Show all", outline=True, size=12)
    d.trigger(s, allb, list(bands.values()) + allbadges)
    hint(s, 8.32, 6.24, 4.4, "Tap a phrase \u2014 it lights up in the script.", size=12)
    d.notes(s, "READ AGAIN \u2014 noticing (5 min).\n"
               "Students scan the script for the five phrases. They point, you tap: the "
               "line where the phrase appears lights up in the script.\n"
               "Ask what the phrase means *in that line* before moving on. \u201cShow "
               "all\u201d reveals every highlighted line at once.")
    return s


def s13_realtalk(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 13, "Real talk", "Real Talk (phrases)", size=30, tw=8.0)
    hint(s, MX, 1.58, 12.0,
         "Tap a sentence when the class has said what the blue phrase means.", size=12.5)

    for i, (sentence, gloss) in enumerate(REAL_TALK):
        y = 2.06 + i * 0.94
        c = card(s, MX, y, 12.09, 0.84, radius=0.18)
        text(s, MX + 0.30, y + 0.12, 6.80, 0.62, sentence, size=14,
             color=INK, line_spacing=0.98)
        g = rect(s, 7.62, y + 0.13, 4.36, 0.58, fill=ACCENT_SOFT, line=None,
                 radius=0.14)
        text_in(g, gloss, size=12.5, color=ACCENT_TEXT, bold=True,
                margin=(0.18, 0.02, 0.14, 0.02))
        d.trigger(s, c, [g])
    d.notes(s, "REAL TALK \u2014 6 min.\n"
               "Five phrases taken from the audio. Read each sentence with the class, ask "
               "what the blue phrase means there, then tap the card: the meaning fades in "
               "on the right.\n"
               "Drill the intonation: \u201cHow on earth\u2026\u201d rises; "
               "\u201cHold on\u201d is short and flat.")
    return s


def s14_practice(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 14, "Practice", "Practice Time", size=32, tw=6.4)
    tag = rect(s, 6.42, 0.96, 3.0, 0.50, fill=ACCENT_SOFT, line=None, radius=0.25)
    text_in(tag, "Real Talk (phrases)", size=12, color=ACCENT_TEXT, bold=True,
            align=PP_ALIGN.CENTER, margin=(0.08, 0, 0.08, 0))
    hint(s, MX, 1.58, 11.0, "Tap a frame when the class has finished it.", size=12.5)

    for i, fr in enumerate(FRAMES):
        y = 2.14 + i * 0.92
        c = card(s, MX, y, 12.09, 0.80, radius=0.18)
        text_in(c, fr, size=19.5, color=INK, margin=(0.35, 0.05, 1.0, 0.05))
        tick = badge(s, 11.85, y + 0.40, 0.40, "ok")
        d.trigger(s, c, [tick])
    d.notes(s, "PRACTICE \u2014 6 min.\n"
               "Students complete each frame with their own idea (a place, a reason, a "
               "name). Take three answers per frame, then tap the card to tick it off.\n"
               "The five frames recycle all five Real Talk phrases \u2014 insist on the "
               "full phrase, not just the gap word.")
    return s


def s15_tf(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 15, "Check understanding", "True or false ?", size=30, tw=6.4,
         ty=0.92)
    hint(s, 4.42, 1.06, 8.2,
         "Decide in pairs, then tap TRUE or FALSE \u2014 and read the line that proves it.",
         size=12.5)

    for i, (stmt, evidence) in enumerate(TF):
        y = 1.88 + i * 2.52
        card(s, MX, y, 12.09, 2.20, radius=0.22)
        text(s, MX + 0.34, y + 0.26, 7.55, 1.00, stmt, size=16, color=INK,
             line_spacing=0.98)

        tbtn = rect(s, 8.62, y + 0.24, 1.44, 0.50, fill=CARD, line="BFE3CB",
                    lw=1.5, radius=0.25)
        text_in(tbtn, "TRUE", size=12, color=OK, bold=True,
                align=PP_ALIGN.CENTER, margin=(0.04, 0, 0.04, 0))
        fbtn = rect(s, 8.62, y + 0.86, 1.44, 0.50, fill=CARD, line="F3C9C4",
                    lw=1.5, radius=0.25)
        text_in(fbtn, "FALSE", size=12, color=NO, bold=True,
                align=PP_ALIGN.CENTER, margin=(0.04, 0, 0.04, 0))
        tick = badge(s, 10.38, y + 0.49, 0.32, "ok")
        cross = badge(s, 10.38, y + 1.11, 0.32, "no")

        band = rect(s, MX + 0.30, y + 1.44, 11.49, 0.58, fill=OK_SOFT,
                    line=None, radius=0.12)
        text_in(band, "TRUE  \u2014  " + evidence, size=12.5, color=OK,
                bold=True, margin=(0.18, 0.02, 0.18, 0.02))
        nband = rect(s, MX + 0.30, y + 1.44, 11.49, 0.58, fill=NO_SOFT,
                     line=None, radius=0.12)
        text_in(nband, "Not quite \u2014 listen again.", size=12.5, color=NO,
                bold=True, margin=(0.18, 0.02, 0.18, 0.02))

        d.trigger(s, tbtn, [band, tick])
        d.trigger(s, fbtn, [nband, cross])
        d.notes(s, "CHECK UNDERSTANDING \u2014 4 min.\n"
                   "Both statements are true. Students decide in pairs and justify with a "
                   "line from the audio.\n"
                   "Tap TRUE to reveal the exact line that proves it; tap FALSE to see a "
                   "gentle \u201clisten again\u201d prompt.\n"
                   "Nothing is printed on the slide, so the class commits to an answer first.")
    return s


def s16_grid(d, idx):
    cfg = GRID[idx]
    s = d.slide()
    backdrop(s)
    q = cfg["q"]
    qs = 26 if len(q) > 62 else 28
    head(d, s, idx, cfg["kicker"], q, size=qs, tw=12.0, ty=0.90)
    for i, w in enumerate(WORKSHEET):
        col, row = i % 2, i // 2
        x = MX + col * 6.16
        y = 2.66 + row * 1.24
        c = rect(s, x, y, 5.96, 1.00, fill=CARD, line=LINE, lw=0.9, radius=0.18)
        shadow(c, blur=16, dist=3, opacity=7)
        text_in(c, w, size=19, color=INK, margin=(0.34, 0.06, 1.0, 0.06))
        b = badge(s, x + 5.46, y + 0.50, 0.42, cfg["marks"][i])
        d.trigger(s, c, [b])
    hint(s, MX, 6.40, 11.7, cfg["note"], size=12)
    d.notes(s, {
        16: "ASKING FOR CLARIFICATION \u2014 3 min.\n"
            "\u201cWhich things would you say when you don\u2019t understand something or "
            "want an explanation?\u201d\n"
            "Answers: I don\u2019t understand / Can you explain it? / How do you\u2026? / "
            "What does that mean? all ask for clarification. \u201cDoes that make "
            "sense?\u201d checks your own explanation and \u201cI get it\u201d confirms "
            "understanding, so those two are red.\n"
            "The original slide printed this worksheet with the ticks already filled in; "
            "here the ticks only appear when a student commits.",
        17: "CONFIRMING \u2014 3 min.\n"
            "\u201cWhat do you say when you understand something?\u201d Only \u201cI get "
            "it.\u201d confirms; every other phrase asks something.\n"
            "Push for the full sentence: \u201cAhaa, I get it now.\u201d",
        18: "CHECKING CLARITY \u2014 3 min.\n"
            "\u201cWhat do you say when you want to know if your explanation is "
            "clear?\u201d \u201cDoes that make sense?\u201d asks the listener to confirm. "
            "The others ask for information or confirmation of understanding.",
    }[idx])
    return s


def s19_speaking(d):
    s = d.slide(dark=True)
    backdrop(s, dark=True)
    progress(s, 19, True)
    page(s, 19, True)
    brand(s, True)
    d.kicker(s, "Speaking", dark=True)
    motif(d, s, 19, True)
    text(s, MX, 0.72, 9.0, 0.7, "Talk Time", size=36, bold=True, font=HEAD,
         color="FFFFFF")

    b = rect(s, MX, 1.56, 7.05, 1.52, fill=DARK_2, line=DARK_LINE, lw=1.0,
             prst=MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT)
    try:
        b.adjustments[0] = 0.2
        b.adjustments[1] = 0.14
        b.adjustments[2] = -0.5
    except Exception:
        pass
    text_in(b, [{"runs": [("Create conversations using", {})], "size": 16.5},
                {"runs": [("phrases from Real Talk", {"color": "AFC6E8"})], "size": 12.5},
                {"runs": [("phrases for asking for clarification & confirming",
                           {"color": "AFC6E8"})], "size": 12.5}],
            size=16, color="FFFFFF", margin=(0.34, 0.16, 0.34, 0.12),
            line_spacing=1.0)

    card(s, 8.10, 1.56, 4.55, 2.06, fill="FFFFFF", radius=0.16, sh=False)
    place(s, "image25.jpeg", 8.16, 1.62, 4.43, 1.94)

    tasks = [("\u2026 when you don\u2019t understand\nsomething",
              "I don\u2019t understand.\nCan you explain it?"),
             ("\u2026 when you understand\nsomething",
              "I get it.\nAhaa, I get it now"),
             ("\u2026 when you check your\nexplanation is clear",
              "Does that make sense?"),
             ("Use a Real Talk phrase", "How on earth?\nhold on / hang out")]
    for i, (label, hints) in enumerate(tasks):
        x = MX + i * 3.02
        c = rect(s, x, 3.54, 2.84, 2.80, fill=DARK_2, line=DARK_LINE, lw=1.0,
                 radius=0.2)
        text(s, x + 0.20, 3.70, 2.44, 1.22, label, size=12.5, color="FFFFFF",
             bold=True, line_spacing=0.98)
        h = rect(s, x + 0.16, 5.08, 2.52, 1.06, fill="1B3A66", line=None,
                 radius=0.12)
        text_in(h, hints, size=12, color="C6D6EE", bold=True,
                margin=(0.12, 0.06, 0.10, 0.06), line_spacing=0.94)
        d.trigger(s, c, [h])

    text(s, MX, 6.34, 11.6, 0.4,
         "Pick a card, then talk for one minute with your partner.",
         size=12, color=MUTED_DARKBG)
    d.notes(s, "SPEAKING \u2014 8 min.\n"
               "Pairs. One student taps a task card; the phrase hints appear inside it. "
               "They then build a one-minute conversation that must include the phrases.\n"
               "Circulate, note two strong conversations, and have those pairs perform at "
               "the end.")
    return s


def s20_quiz(d, idx):
    cfg = QUIZ[idx]
    s = d.slide()
    backdrop(s)
    head(d, s, idx, "Practice", "Choose the suitable phrase", size=28, tw=9.4,
         ty=0.94)

    y = 1.62
    for txt, kind in cfg["lines"]:
        size = 16
        w = 6.05 if kind == "in" else 6.35
        x = MX + (0.02 if kind == "in" else 0.72)
        nl = n_lines(txt, size, w - 0.9)
        h = max(0.92, nl * size * 1.3 / 72.0 + 0.40)
        bubble(s, x, y, w, h, txt, size=size, align=PP_ALIGN.LEFT,
               fill=(ACCENT_SOFT if kind == "in" else CARD),
               line_c=(ACCENT_LINE if kind == "in" else LINE),
               margin=(0.30, 0.16, 0.52, 0.16))
        y += h + 0.26

    card(s, MX, 4.72, 2.45, 1.78, fill=BG_TINT, radius=0.2, sh=False)
    place(s, "image25.jpeg" if cfg["art"] == "image25.jpeg" else cfg["art"],
          MX + 0.06, 4.78, 2.33, 1.66) if cfg["art"].endswith(".jpeg") else \
        place(s, cfg["art"], MX + 0.06, 4.78, 2.33, 1.66)

    badges, ring_ok = [], None
    for i, opt in enumerate(cfg["options"]):
        oy = 1.66 + i * 1.18
        c = rect(s, 7.05, oy, 5.66, 0.96, fill=CARD, line=LINE, lw=1.0, radius=0.18)
        shadow(c, blur=16, dist=3, opacity=7)
        text(s, 7.28, oy + 0.27, 0.4, 0.4, "ABC"[i], size=15, bold=True,
             color=MUTED, font=HEAD)
        text_in(c, opt, size=19, color=INK, margin=(0.66, 0.06, 0.85, 0.06))
        if i == cfg["correct"]:
            ring_ok = rect(s, 6.98, oy - 0.07, 5.80, 1.10, fill=None, line=OK,
                           lw=2.6, radius=0.2)
            tick = badge(s, 12.12, oy + 0.48, 0.40, "ok")
            badges += [ring_ok, tick]
            d.trigger(s, c, [ring_ok, tick])
        else:
            cross = badge(s, 12.12, oy + 0.48, 0.40, "no")
            badges.append(cross)
            d.trigger(s, c, [cross])

    why = rect(s, 7.05, 5.20, 5.66, 0.72, fill=OK_SOFT, line="BFE3CB", lw=1.0,
               radius=0.16)
    text_in(why, cfg["why"], size=12, color=OK, bold=True,
            margin=(0.18, 0.04, 0.16, 0.04))
    show = button(s, 7.05, 6.06, 2.70, 0.48, "Show answer", outline=True, size=12)
    d.trigger(s, show, badges)
    hint(s, 10.05, 6.14, 2.6, "Tap an option to answer.", size=12)
    d.notes(s, {
        20: "CHOOSE THE PHRASE \u2014 3 min.\n"
            "Read the exchange. Students choose the phrase that fills the gap, then tap an "
            "option: green ring + tick for the right one, a red cross for the others.\n"
            "Answer: \u201cHang out\u201d \u2014 \u201cLet\u2019s hang out in the "
            "mall\u201d means let\u2019s spend time there together.",
        21: "CHOOSE THE PHRASE \u2014 3 min.\n"
            "Answer: \u201cget it\u201d (\u201cYes, I get it\u201d = yes, I understand).\n"
            "\u201cgetting the long of it\u201d is deliberately close to \u201cgetting "
            "the hang of it\u201d \u2014 use it to contrast the two.",
        22: "CHOOSE THE PHRASE \u2014 3 min.\n"
            "Answer: \u201cCan you explain..\u201d \u2014 the request that fits "
            "\u201c\u2026 google maps to me, please?\u201d\n"
            "\u201cDoes that make sense?\u201d checks the listener\u2019s understanding; "
            "it does not ask for an explanation.",
    }[idx])
    return s


def s23_order(d):
    s = d.slide()
    backdrop(s)
    head(d, s, 23, "Put it in order", "Put the conversation in the correct order",
         size=27, tw=11.7, ty=0.92)

    boxes = [(0.98, 1.70, 5.30, 1.24, 0), (6.90, 1.70, 5.30, 1.24, 1),
             (0.98, 3.30, 5.30, 1.24, 2), (6.90, 3.30, 5.30, 1.24, 3)]
    nums = []
    for x, y, w, h, i in boxes:
        txt, n = REORDER[i]
        b = rect(s, x, y, w, h, fill=CARD, line=LINE, lw=1.0, radius=0.18)
        shadow(b, blur=18, dist=4, opacity=8)
        text(s, x + 0.92, y + 0.22, w - 1.14, h - 0.40, txt, size=17, color=INK,
             line_spacing=0.98)
        slot = rect(s, x + 0.20, y + 0.20, 0.52, 0.52, fill=BG_TINT,
                    line="CFDAEB", lw=1.2, radius=0.26)
        num = rect(s, x + 0.20, y + 0.20, 0.52, 0.52, fill=ACCENT, line=None,
                   prst=MSO_SHAPE.OVAL)
        text_in(num, str(n), size=15, color="FFFFFF", bold=True,
                align=PP_ALIGN.CENTER, margin=(0, 0, 0, 0))
        d.trigger(s, b, [num])
        nums.append(num)

    show = button(s, 0.98, 5.00, 3.0, 0.50, "Show the order", outline=True,
                  size=12)
    d.trigger(s, show, nums)
    hint(s, 0.98, 5.64, 8.0, "Tap each message to check its place in the conversation.",
         size=12)
    place(s, "image22.png", 10.95, 4.96, 1.30, 1.30)
    d.notes(s, "ORDER THE CONVERSATION \u2014 4 min.\n"
               "Pairs decide the order, then tap each bubble: its number fades in at the "
               "corner.\n"
               "Correct order: the question \u2192 the instruction \u2192 \u201cGreat\u201d "
               "\u2192 \u201cAhaa, I get it now\u201d.\n"
               "\u201cShow the order\u201d reveals all four numbers for the whole class.")
    return s


def s24_thanks(d):
    s = d.slide(dark=True)
    backdrop(s, dark=True)
    progress(s, 24, True)
    page(s, 24, True)
    brand(s, True)
    motif(d, s, 24, True)
    text(s, MX, 2.20, 11.0, 1.0, "Thank you!", size=58, bold=True, font=HEAD,
         color="FFFFFF")
    text(s, MX, 3.38, 11.0, 0.7, "You have a wonderful day", size=30,
         color="8FB0FF")
    rect(s, MX, 4.26, 1.1, 0.07, fill="5B8CFF")
    text(s, MX, 4.55, 9.0, 0.4,
         "Connected by Technology \u00b7 Conversation", size=12.5, color=MUTED_DARKBG)

    button(s, MX, 5.45, 2.95, 0.58, "Back to the start", outline=True,
           name="btn_back")
    brand(s, dark=True)
    d.notes(s, "CLOSING.\n"
               "Recap the five objectives in one breath, then send the class off.\n"
               "\u201cBack to the start\u201d jumps to the title slide, so the deck can be "
               "reused with the next class without closing it.")
    return s


# ---------------------------------------------------------------------------
def main():
    d = Deck(OUT, total=TOTAL)
    s01_title(d)
    s02_objectives(d)
    s03_warmup(d)
    s04_what(d)
    s05_language(d)
    s06_predict(d)
    s07_questions(d)
    s08_read(d)
    s09_task(d)
    s10_roleplay(d)
    s11_flash(d)
    s12_find(d)
    s13_realtalk(d)
    s14_practice(d)
    s15_tf(d)
    s16_grid(d, 16)
    s16_grid(d, 17)
    s16_grid(d, 18)
    s19_speaking(d)
    s20_quiz(d, 20)
    s20_quiz(d, 21)
    s20_quiz(d, 22)
    s23_order(d)
    s24_thanks(d)

    for sl in d.slides[1:]:
        d.morph(sl)

    slides = d.prs.slides
    d.hyperlink(find(slides[0], "Start the lesson"), slides[1])
    d.hyperlink(find(slides[-1], "Back to the start"), slides[0])

    d.save()
    print(f"wrote {OUT} \u2014 {len(d.slides)} slides, "
          f"{sum(1 for _ in slides)} checked")


def find(slide, label):
    for sh in slide.shapes:
        if sh.has_text_frame and label == sh.text_frame.text.strip():
            return sh
        if sh.has_text_frame and label in sh.text_frame.text:
            return sh
    raise KeyError(label)


if __name__ == "__main__":
    main()
