#!/usr/bin/env python3
"""Palette Lab — generate contrast-safe generator themes from one seed color.

Builds full 7-role themes (bg/card/ink/muted/accent/accent2/graybar + dark
flag) using classic color harmonies, then *repairs* them until they pass
WCAG floors (reference/10): ink:bg >= 10:1, muted:bg >= 4.5:1, accent:bg
>= 3:1 (large graphics). Themes land in database/data/themes.json, which the
generator merges over its built-ins — usable immediately by name in specs.

Usage:
  python tools/palette_lab.py --seed FF6B4A --name coral-pop [--harmony split] [--dark]
  python tools/palette_lab.py --pack           # curated set of 8 ready themes
  python tools/palette_lab.py --list           # what's in themes.json
"""
import argparse
import colorsys
import json
import os
import sys

THEMES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "..", "database", "data", "themes.json")


# ---------------------------------------------------------------- color math
def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def rgb_to_hex(r, g, b):
    return "".join(f"{round(max(0, min(1, v)) * 255):02X}" for v in (r, g, b))


def rel_lum(h):
    def f(c):
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = hex_to_rgb(h)
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a, b):
    la, lb = rel_lum(a), rel_lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def hsl(h):
    return colorsys.rgb_to_hls(*hex_to_rgb(h))  # h, l, s


def from_hsl(h, l, s):
    return rgb_to_hex(*colorsys.hls_to_rgb(h % 1.0, max(0.0, min(1.0, l)),
                                           max(0.0, min(1.0, s))))


def hue_rotate(h, degrees):
    hh, l, s = hsl(h)
    return from_hsl(hh + degrees / 360, l, s)


def _fix(fg, bg, target, darken=True):
    """Nudge lightness until fg:bg contrast meets target (24 steps max)."""
    hh, l, s = hsl(fg)
    step = -0.03 if darken else 0.03
    for _ in range(24):
        if contrast(from_hsl(hh, l, s), bg) >= target:
            break
        l = max(0.02, min(0.98, l + step))
    return from_hsl(hh, l, s)


# ------------------------------------------------------------------- builder
HARMONIES = {"complementary": 180, "analogous": 35, "triadic": 120,
             "split": 160, "mono": 0}


def build_theme(name, accent, accent2=None, dark=False):
    """Full 7-role theme from an accent (+optional second accent).
    Shared by brandkit.py and anatomy.py. Returns (theme, report_rows)."""
    accent = accent.upper().lstrip("#")
    if accent2 is None:
        hh, l, s = hsl(accent)
        accent2 = from_hsl(hh, min(0.82, l + 0.22), s)
    accent2 = accent2.upper().lstrip("#")
    hh, l, s = hsl(accent)

    if dark:
        bg = from_hsl(hh, 0.07, min(0.35, s * 0.45))
        card = from_hsl(hh, 0.135, min(0.35, s * 0.45))
        ink = from_hsl(hh, 0.96, 0.10)
        muted = from_hsl(hh, 0.68, 0.14)
        graybar = from_hsl(hh, 0.38, 0.12)
        accent = _fix(accent, bg, 3.0, darken=False)
        accent2 = _fix(accent2, bg, 2.6, darken=False)
        muted = _fix(muted, bg, 4.5, darken=False)
    else:
        bg = from_hsl(hh, 0.972, min(0.30, s * 0.18))
        card = "FFFFFF"
        ink = from_hsl(hh, 0.13, min(0.55, s * 0.5))
        muted = from_hsl(hh, 0.40, 0.13)
        graybar = from_hsl(hh, 0.80, 0.10)
        accent = _fix(accent, bg, 3.0, darken=True)
        muted = _fix(muted, bg, 4.5, darken=True)
    ink = _fix(ink, bg, 10.0, darken=not dark)

    theme = dict(bg=bg, card=card, ink=ink, muted=muted, accent=accent,
                 accent2=accent2, graybar=graybar, dark=dark)
    report = [(f"ink:bg {contrast(ink, bg):.1f}:1 (floor 10)", contrast(ink, bg) >= 10),
              (f"muted:bg {contrast(muted, bg):.1f}:1 (floor 4.5)", contrast(muted, bg) >= 4.5),
              (f"accent:bg {contrast(accent, bg):.1f}:1 (floor 3)", contrast(accent, bg) >= 3.0)]
    return theme, report


# ---------------------------------------------------------------------- io
def load_themes():
    if os.path.exists(THEMES_PATH):
        with open(THEMES_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_theme(name, theme):
    themes = load_themes()
    themes[name] = theme
    with open(THEMES_PATH, "w", encoding="utf-8") as f:
        json.dump(themes, f, indent=2, sort_keys=True)
        f.write("\n")
    return len(themes)


def show(name, theme, report):
    print(f"  {name:20s} bg=#{theme['bg']} card=#{theme['card']} "
          f"ink=#{theme['ink']} muted=#{theme['muted']} "
          f"accent=#{theme['accent']} accent2=#{theme['accent2']} "
          f"graybar=#{theme['graybar']} {'dark' if theme['dark'] else 'light'}")
    for line, ok in report:
        print(f"      {'PASS' if ok else 'WARN'}  {line}")


PACK = [  # name, seed, harmony degrees, dark
    ("coral-pop", "FF6B4A", 160, False),
    ("midnight-lime", "A3E635", 0, True),
    ("berry-cream", "C026D3", 35, False),
    ("deep-ocean", "0891B2", 180, True),
    ("amber-dusk", "D97706", 35, True),
    ("forest-paper", "15803D", 120, False),
    ("rose-noir", "F43F5E", 0, True),
    ("sky-mist", "0284C7", 35, False),
]


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--seed", help="base accent hex, e.g. FF6B4A")
    ap.add_argument("--name", help="theme name (kebab-case)")
    ap.add_argument("--harmony", default="split", choices=list(HARMONIES),
                    help="how accent2 is derived (default split)")
    ap.add_argument("--dark", action="store_true", help="dark theme")
    ap.add_argument("--pack", action="store_true", help="generate curated 8-pack")
    ap.add_argument("--list", action="store_true", help="list themes.json")
    a = ap.parse_args(argv[1:])

    if a.list:
        themes = load_themes()
        print(f"{len(themes)} external themes in database/data/themes.json")
        for n in sorted(themes):
            t = themes[n]
            print(f"  {n:20s} bg=#{t['bg']} accent=#{t['accent']} "
                  f"{'dark' if t['dark'] else 'light'}")
        return 0

    jobs = []
    if a.pack:
        jobs = [(n, seed, deg, dk) for n, seed, deg, dk in PACK]
    elif a.seed and a.name:
        jobs = [(a.name, a.seed, HARMONIES[a.harmony], a.dark)]
    else:
        ap.error("give --seed + --name, or --pack / --list")

    for name, seed, deg, dk in jobs:
        accent2 = None if deg == 0 else hue_rotate(seed, deg)
        theme, report = build_theme(name, seed, accent2, dark=dk)
        total = save_theme(name, theme)
        show(name, theme, report)
    print(f"\nthemes.json now holds {total} themes — use in specs by name, e.g. "
          f'"theme": "{jobs[0][0]}"')
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
