#!/usr/bin/env python3
"""ImageKit — make photos belong to the deck's theme.

Treatments (combinable, applied in this order): crop to a deck ratio,
duotone (luminance mapped ink→accent), brand tint (accent wash), and a
bottom shade gradient so white headlines survive busy photos. Colors come
from any generator theme by name (incl. database/data/themes.json ones).

Usage:
  python tools/imgkit.py photo.jpg --theme teal-orbit --duotone --ratio 16:9
  python tools/imgkit.py photo.jpg --accent FF6B4A --tint 0.25 --shade
  python tools/imgkit.py photo.jpg --ratio 9:16 --focus top --out hero.jpg
"""
import argparse
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
import generate_pptx as G  # noqa: E402

RATIOS = {"16:9": (1920, 1080), "9:16": (1080, 1920), "4:5": (1080, 1350),
          "1:1": (1350, 1350), "4:3": (1600, 1200)}


def hx(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def crop_ratio(img, ratio, focus="center"):
    tw, th = RATIOS[ratio]
    target = tw / th
    w, h = img.size
    cur = w / h
    if cur > target:      # too wide -> trim sides
        nw = int(h * target)
        x0 = {"left": 0, "right": w - nw}.get(focus, (w - nw) // 2)
        box = (x0, 0, x0 + nw, h)
    else:                 # too tall -> trim top/bottom
        nh = int(w / target)
        y0 = {"top": 0, "bottom": h - nh}.get(focus, (h - nh) // 2)
        box = (0, y0, w, y0 + nh)
    return img.crop(box).resize((tw, th))


def duotone(img, dark_hex, light_hex):
    from PIL import Image
    gray = img.convert("L")
    d, l = hx(dark_hex), hx(light_hex)
    lut = []
    for ch in range(3):
        lut += [round(d[ch] + (l[ch] - d[ch]) * i / 255) for i in range(256)]
    return gray.point(lut * 1).convert("RGB") if False else \
        Image.merge("RGB", [gray.point(lut[c * 256:(c + 1) * 256])
                            for c in range(3)])


def tint(img, accent_hex, amount):
    from PIL import Image
    overlay = Image.new("RGB", img.size, hx(accent_hex))
    return Image.blend(img.convert("RGB"), overlay, amount)


def shade(img, strength=0.68, coverage=0.5):
    from PIL import Image
    img = img.convert("RGBA")
    w, h = img.size
    grad = Image.new("L", (1, h), 0)
    start = int(h * (1 - coverage))
    for y in range(h):
        a = 0 if y < start else int(255 * strength * (y - start) /
                                    max(h - start, 1))
        grad.putpixel((0, y), a)
    black = Image.new("RGBA", (w, h), (0, 0, 0, 255))
    black.putalpha(grad.resize((w, h)))
    return Image.alpha_composite(img, black).convert("RGB")


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("image")
    ap.add_argument("--theme", default="", help="theme name for colors")
    ap.add_argument("--accent", default="", help="accent hex (overrides theme)")
    ap.add_argument("--ink", default="", help="duotone shadow hex")
    ap.add_argument("--ratio", choices=list(RATIOS), help="crop to ratio")
    ap.add_argument("--focus", default="center",
                    choices=["center", "top", "bottom", "left", "right"])
    ap.add_argument("--duotone", action="store_true")
    ap.add_argument("--tint", type=float, default=0.0,
                    help="accent wash 0..1 (e.g. 0.25)")
    ap.add_argument("--shade", action="store_true",
                    help="bottom gradient for text legibility")
    ap.add_argument("--out", help="output path (default <image>.treated.jpg)")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.image):
        print(f"no such file: {a.image}")
        return 2
    from PIL import Image
    theme = G.THEMES.get(a.theme or "", {})
    accent = (a.accent or theme.get("accent", "14C8B4")).lstrip("#")
    ink = (a.ink or theme.get("ink", "10222A")).lstrip("#")
    bg = theme.get("bg", "0B0B10")
    if a.theme and a.theme not in G.THEMES:
        print(f"unknown theme '{a.theme}' — see generate_pptx.py --themes")
        return 2

    img = Image.open(a.image)
    steps = []
    if a.ratio:
        img = crop_ratio(img, a.ratio, a.focus)
        steps.append(f"crop {a.ratio} ({a.focus})")
    if a.duotone:
        # dark themes: shadows to bg, highlights to accent; light: ink->accent
        dark = bg if theme.get("dark") else ink
        img = duotone(img, dark, accent)
        steps.append(f"duotone #{dark}→#{accent}")
    if a.tint:
        img = tint(img, accent, min(max(a.tint, 0.0), 1.0))
        steps.append(f"tint #{accent} @{a.tint}")
    if a.shade:
        img = shade(img)
        steps.append("bottom shade")
    out = a.out or os.path.splitext(a.image)[0] + ".treated.jpg"
    img.convert("RGB").save(out, quality=90)
    print(f"{out}: {img.size[0]}x{img.size[1]} — {', '.join(steps) or 'copied'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
