#!/usr/bin/env python3
"""BrandKit — turn a logo (any image) into a contrast-safe generator theme.

Extracts the dominant saturated colors from the image (Pillow median-cut
quantization), picks accent + accent2 by hue distance, then builds the
full 7-role theme with Palette Lab's WCAG repair. The theme lands in
database/data/themes.json and works by name in specs immediately.

Usage:
  python tools/brandkit.py logo.png --name acme-brand [--dark]
  python tools/brandkit.py logo.png --name acme-light --also-dark acme-dark
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from palette_lab import build_theme, hsl, rgb_to_hex, save_theme, show  # noqa: E402


def extract_palette(path, n=10):
    from PIL import Image
    img = Image.open(path).convert("RGBA")
    img.thumbnail((400, 400))
    # composite onto white so transparent PNGs read truthfully
    bgimg = Image.new("RGB", img.size, (255, 255, 255))
    bgimg.paste(img, mask=img.split()[3])
    q = bgimg.quantize(colors=n, method=Image.MEDIANCUT)
    counts = sorted(q.getcolors(), reverse=True)  # [(count, palette_idx)]
    pal = q.getpalette()
    total = sum(c for c, _ in counts) or 1
    out = []
    for count, idx in counts:
        r, g, b = (pal[idx * 3 + i] / 255 for i in range(3))
        out.append((count / total, rgb_to_hex(r, g, b)))
    return _vividify(bgimg, out)


def _vividify(img, palette):
    """Median-cut averages wash brand colors out (antialias edges). For each
    palette hue, swap in the most saturated *real* pixel near that hue."""
    small = img.copy()
    small.thumbnail((200, 200))
    raw = small.tobytes()
    pixels = []
    for i in range(0, len(raw) - 2, 3):
        hxs = rgb_to_hex(raw[i] / 255, raw[i + 1] / 255, raw[i + 2] / 255)
        h, l, s = hsl(hxs)
        pixels.append((h, l, s, hxs))
    out = []
    for share, hexs in palette:
        h, l, s = hsl(hexs)
        if s < 0.22 or l <= 0.08 or l >= 0.93:
            out.append((share, hexs))  # grays/blacks/whites pass through raw
            continue
        best, best_s = hexs, s
        for ph, pl, ps, px in pixels:
            dh = abs(ph - h) % 1.0
            dh = min(dh, 1 - dh)
            if dh < 0.035 and 0.18 < pl < 0.88 and ps > best_s:
                best, best_s = px, ps
        out.append((share, best))
    return out


def pick_accents(palette):
    """Accent = most common *saturated, mid-lightness* color; accent2 =
    the next saturated color whose hue is 25-200 degrees away."""
    scored = []
    for share, hexs in palette:
        h, l, s = hsl(hexs)
        if s < 0.22 or l < 0.08 or l > 0.93:
            continue  # grays / near-black / near-white never make accents
        window = 1.0 - min(1.0, abs(l - 0.52) / 0.45)  # favor highlight-usable L
        scored.append(((share ** 0.5) * (s ** 1.5) * (0.25 + window), hexs, h))
    scored.sort(reverse=True)
    if not scored:
        return None, None
    accent, ah = scored[0][1], scored[0][2]
    accent2 = None
    for _, hexs, h in scored[1:]:
        dist = abs(h - ah) % 1.0
        dist = min(dist, 1 - dist) * 360
        if 25 <= dist <= 200:
            accent2 = hexs
            break
    return accent, accent2


def dominant_is_dark(palette):
    top_share, top_hex = palette[0]
    _, l, _ = hsl(top_hex)
    return top_share > 0.35 and l < 0.25


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("image", help="path to logo/image (png/jpg/webp)")
    ap.add_argument("--name", required=True, help="theme name (kebab-case)")
    ap.add_argument("--dark", action="store_true",
                    help="force dark theme (default: auto from logo bg)")
    ap.add_argument("--light", action="store_true", help="force light theme")
    ap.add_argument("--also-dark", metavar="NAME",
                    help="also emit a dark sibling under this name")
    a = ap.parse_args(argv[1:])

    if not os.path.exists(a.image):
        print(f"no such file: {a.image}")
        return 2
    palette = extract_palette(a.image)
    print("extracted palette (share, color):")
    for share, hexs in palette[:6]:
        print(f"    {share:5.1%}  #{hexs}")
    accent, accent2 = pick_accents(palette)
    if not accent:
        print("no saturated brand color found — image is grayscale/monochrome?")
        return 1
    print(f"accent=#{accent}  accent2=#{accent2 or '(derived)'}")

    dark = a.dark or (dominant_is_dark(palette) and not a.light)
    theme, report = build_theme(a.name, accent, accent2, dark=dark)
    total = save_theme(a.name, theme)
    show(a.name, theme, report)
    if a.also_dark:
        t2, r2 = build_theme(a.also_dark, accent, accent2, dark=True)
        total = save_theme(a.also_dark, t2)
        show(a.also_dark, t2, r2)
    print(f"\nthemes.json now holds {total} themes — build with "
          f'"theme": "{a.name}"')
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
