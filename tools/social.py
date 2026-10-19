#!/usr/bin/env python3
"""Social Factory — one deck spec -> Reel (9:16) + Carousel (4:5) pack.

Converts a landscape generator spec into portrait slide types (or takes a
portrait spec directly), builds .pptx for each format via
generator/portrait.py, and writes the posting kit: caption.txt (hook line
+ hashtags) and shotlist.md (per-slide on-screen text, seconds, voiceover
line). Reels auto-advance (advTm) so File > Export > Video just works.

Usage:
  python tools/social.py generator/deck-spec.example.json --outdir /tmp/pack
  python tools/social.py spec.json --format reel --handle @Loopwise \
      --tags "#startups #Riyadh" --pick 1,4,5,12
"""
import argparse
import json
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
import portrait  # noqa: E402

PORTRAIT_TYPES = set(portrait.BUILDERS)


def _cap(text, n=80):
    """Social copy is terse — cap at a word boundary; full text lives in
    the voiceover (notes)."""
    text = str(text or "")
    if len(text) <= n:
        return text
    return text[:n].rsplit(" ", 1)[0] + "…"


def conv(sd):
    """Landscape slide dict -> portrait slide dict (or None to skip)."""
    t = sd.get("type", "content")
    base = {"notes": sd.get("notes", "")}
    if t == "title":
        lines = sd.get("title_lines") or str(sd.get("title", "")).split("\n")
        return dict(type="hook", badge=sd.get("badge") or sd.get("kicker", ""),
                    title_lines=[l for l in lines if l],
                    subtitle=sd.get("subtitle", ""), **base)
    if t == "section":
        return dict(type="hook", badge=f"PART {sd.get('number', '')}".strip(),
                    title_lines=[sd.get("title", "")],
                    subtitle=sd.get("subtitle", ""), **base)
    if t == "hero_stat":
        return dict(type="bigstat", stat=sd.get("stat", ""),
                    label=sd.get("label", ""), context=sd.get("context", ""),
                    title=sd.get("title", ""), kicker=sd.get("kicker", ""),
                    **base)
    if t == "stat_strip":
        s0 = (sd.get("stats") or [{}])[0]
        return dict(type="bigstat", stat=s0.get("value", ""),
                    label=s0.get("label", ""), context=s0.get("context", ""),
                    title=sd.get("title", ""), kicker=sd.get("kicker", ""),
                    **base)
    if t in ("chart", "donut"):
        return dict(type="bars", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""),
                    categories=sd.get("categories", []),
                    values=sd.get("values", []),
                    highlight=sd.get("highlight", 0),
                    insight=sd.get("insight", ""), **base)
    if t == "line_chart":
        srs = (sd.get("series") or [{}])[-1]
        return dict(type="bars", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""),
                    categories=sd.get("categories", []),
                    values=srs.get("values", []), highlight=0,
                    insight=sd.get("insight", ""), **base)
    if t in ("content", "team"):
        if t == "team":
            items = [dict(icon="star", head=p.get("name", ""),
                          body=_cap(f"{p.get('role','')} — {p.get('bio','')}"))
                     for p in sd.get("people", [])]
        else:
            items = [dict(icon="idea", head=b.get("head", ""),
                          body=_cap(b.get("body", "")))
                     for b in sd.get("bullets", [])]
        return dict(type="tips", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""), items=items, **base)
    if t == "icon_grid":
        items = [dict(it, body=_cap(it.get("body", "")))
                 for it in sd.get("items", [])]
        return dict(type="tips", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""), items=items, **base)
    if t in ("process", "timeline"):
        return dict(type="steps", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""), steps=sd.get("steps", []), **base)
    if t == "agenda":
        return dict(type="steps", kicker=sd.get("kicker", "ROADMAP"),
                    title=sd.get("title", ""),
                    steps=[dict(head=str(x), body="") for x in sd.get("items", [])],
                    **base)
    if t == "quote":
        return dict(type="quote", quote=sd.get("quote", ""),
                    attribution=sd.get("attribution", ""), **base)
    if t in ("before_after", "comparison"):
        if t == "before_after":
            b, af = sd.get("before", {}), sd.get("after", {})
        else:
            cols = sd.get("columns", [])
            b = next((c for c in cols if not c.get("highlight")), {})
            af = next((c for c in cols if c.get("highlight")), {})
        return dict(type="compare", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""), before=b, after=af, **base)
    if t in ("market", "tam_sam_som"):
        rings = sd.get("rings") or [{}]
        som = rings[-1]
        return dict(type="bigstat", kicker="MARKET", title=sd.get("title", ""),
                    stat=som.get("value", ""), label=som.get("label", ""),
                    context=som.get("note", ""), **base)
    if t == "closing":
        return dict(type="cta",
                    ask=sd.get("ask") or " ".join(sd.get("title_lines", []) or
                                                   [sd.get("title", "")]),
                    subtitle=sd.get("subtitle", ""),
                    contact=sd.get("contact", ""), **base)
    if t in ("waterfall", "funnel"):
        items = sd.get("steps") or sd.get("stages") or []
        return dict(type="bars", kicker=sd.get("kicker", ""),
                    title=sd.get("title", ""),
                    categories=[i.get("label", "") for i in items],
                    values=[abs(float(i.get("value", 0))) for i in items],
                    highlight=0, insight=sd.get("insight", ""), **base)
    if t == "matrix":
        hero = next((i for i in sd.get("items", []) if i.get("highlight")), {})
        return dict(type="bigstat", kicker="POSITION", title=sd.get("title", ""),
                    stat="★", label=hero.get("label", ""),
                    context=sd.get("insight", ""), **base)
    if t == "logo_wall":
        return dict(type="tips", kicker=sd.get("kicker", "PROOF"),
                    title=sd.get("title", ""),
                    items=[dict(icon="check", head=n, body="")
                           for n in sd.get("names", [])[:4]], **base)
    return None


def convert_spec(spec, fmt, pick=None, handle=""):
    src = spec.get("slides", [])
    if all(s.get("type") in PORTRAIT_TYPES for s in src):
        slides = [dict(s) for s in src]
        skipped = []
    else:
        slides, skipped = [], []
        for i, sd in enumerate(src, 1):
            if pick and i not in pick:
                continue
            c = conv(sd)
            if c:
                slides.append(c)
            else:
                skipped.append((i, sd.get("type")))
    if handle:
        for s in slides:
            if s.get("type") == "cta" and not s.get("contact"):
                s["handle"] = handle
    out = {"format": fmt, "theme": spec.get("theme", "teal-orbit"),
           "brand": spec.get("brand", ""), "rtl": bool(spec.get("rtl")),
           "title": spec.get("title", f"{fmt} pack"), "slides": slides}
    return out, skipped


def shotlist(fmt, pspec):
    lines = [f"# Shotlist — {fmt} ({len(pspec['slides'])} slides)", "",
             "| # | type | on-screen headline | seconds | voiceover (notes) |",
             "|---|---|---|---|---|"]
    total = 0
    for i, s in enumerate(pspec["slides"], 1):
        head = (s.get("title") or " ".join(s.get("title_lines", [])) or
                s.get("stat") or s.get("ask") or s.get("quote", "")[:40] or "")
        secs = s.get("seconds")
        if fmt == "reel" and not secs:
            w = len(str(s.get("notes", "")).split())
            secs = round(max(6.0, min(45.0, w / 130 * 60)), 1) if w else 4.0
        total += float(secs or 0)
        vo = str(s.get("notes", "")).replace("|", "/")[:110]
        lines.append(f"| {i} | {s.get('type')} | {str(head)[:44]} | "
                     f"{secs if secs else '—'} | {vo} |")
    lines.append("")
    if fmt == "reel":
        lines.append(f"Estimated runtime ≈ {total:.0f}s. Record VO per row, "
                     f"or export silent: PowerPoint File → Export → Create "
                     f"Video (timings are already in the file).")
    else:
        lines.append("Export slides as images (1080×1350 target) and upload "
                     "in order — max 20, never mix ratios in one post.")
    return "\n".join(lines) + "\n"


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("spec", help="landscape deck spec or portrait spec JSON")
    ap.add_argument("--format", choices=["reel", "carousel", "both"],
                    default="both")
    ap.add_argument("--outdir", default="social-out")
    ap.add_argument("--handle", default="", help="e.g. @Loopwise (CTA chip)")
    ap.add_argument("--tags", default="", help="hashtags for the caption")
    ap.add_argument("--pick", default="",
                    help="comma list of source slide numbers to convert")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.spec):
        print(f"no such file: {a.spec}")
        return 2
    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)
    pick = {int(x) for x in a.pick.split(",") if x.strip()} or None
    fmts = ["reel", "carousel"] if a.format == "both" else [a.format]
    os.makedirs(a.outdir, exist_ok=True)

    shots = []
    reel_spec = None
    for fmt in fmts:
        pspec, skipped = convert_spec(spec, fmt, pick, a.handle)
        sp = os.path.join(a.outdir, f"{fmt}.spec.json")
        dp = os.path.join(a.outdir, f"{fmt}.pptx")
        with open(sp, "w", encoding="utf-8") as f:
            json.dump(pspec, f, indent=2, ensure_ascii=False)
        n = portrait.build_portrait(pspec, dp)
        print(f"{fmt}: {n} portrait slides -> {dp} (+ {sp})")
        if skipped:
            print(f"   skipped (no portrait mapping): {skipped}")
        shots.append(shotlist(fmt, pspec))
        if fmt == "reel":
            reel_spec = pspec

    # caption: claim from first convertible hook, else spec title
    hook = ""
    for s in spec.get("slides", []):
        c = conv(s) if s.get("type") not in PORTRAIT_TYPES else s
        if c and c.get("type") == "hook":
            hook = " ".join(c.get("title_lines", []))
            sub = c.get("subtitle", "")
            hook = (hook + (" — " + sub if sub else "")).strip()
            break
    caption = (hook or spec.get("title", "")).strip()
    if a.handle:
        caption += f"\n\nFollow {a.handle} for more."
    if a.tags:
        caption += f"\n\n{a.tags}"
    with open(os.path.join(a.outdir, "caption.txt"), "w", encoding="utf-8") as f:
        f.write(caption + "\n")
    with open(os.path.join(a.outdir, "shotlist.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(shots))
    extra = ""
    if reel_spec:  # paste-ready voiceover script (record or any TTS)
        vo = ["VOICEOVER SCRIPT — one block per slide; bracket = on-screen seconds", ""]
        for i, s2 in enumerate(reel_spec["slides"], 1):
            secs = s2.get("seconds")
            if not secs:
                w = len(str(s2.get("notes", "")).split())
                secs = round(max(6.0, min(45.0, w / 130 * 60)), 1) if w else 4.0
            line = str(s2.get("notes", "")).strip() or \
                str(s2.get("title") or " ".join(s2.get("title_lines", [])))
            vo.append(f"[{i:02d} · {secs}s] {line}")
            vo.append("")
        with open(os.path.join(a.outdir, "voiceover.txt"), "w",
                  encoding="utf-8") as f:
            f.write("\n".join(vo))
        extra = " + voiceover.txt"
    print(f"posting kit: caption.txt + shotlist.md{extra} in {a.outdir}")
    print("QA them: python tools/qa_deck.py " +
          " ".join(os.path.join(a.outdir, f"{f}.pptx") for f in fmts))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
