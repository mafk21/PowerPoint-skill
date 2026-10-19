#!/usr/bin/env python3
"""Speaker Coach — rehearsal analytics for any .pptx.

Reads speaker notes, estimates per-slide talk time (~130 wpm, the same
model the generator's auto-timings use), flags slides you're probably
*reading* (on-slide text nearly as long as your script), slides with no
script at all, and runtime overruns vs a target. Can emit a rehearsal
cue-card sheet you can actually practice from.

Usage:
  python tools/coach.py deck.pptx [--target-min 10] [--wpm 130]
  python tools/coach.py deck.pptx --write-md rehearsal.md
"""
import argparse
import os
import re
import sys

from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from anatomy import walk  # noqa: E402


def slide_facts(slide):
    words, headline, best_pt = 0, "", 0.0
    vocab = set()
    for sh in walk(slide.shapes):
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                if not t:
                    continue
                words += len(t.split())
                vocab.update(w.lower().strip(".,:%—-–()[]") for w in t.split()
                             if len(w) > 3)
                pt = run.font.size.pt if run.font.size else 0.0
                if pt > best_pt and len(t) > 3:
                    best_pt, headline = pt, t
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    return headline, words, notes, vocab


def overlap(vocab, notes):
    if not vocab or not notes:
        return 0.0
    nvocab = {w.lower().strip(".,:%—-–()[]") for w in notes.split() if len(w) > 3}
    return len(vocab & nvocab) / len(vocab)


def fmt(sec):
    m, s = divmod(int(round(sec)), 60)
    return f"{m}:{s:02d}"


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck")
    ap.add_argument("--wpm", type=float, default=130,
                    help="speaking rate (default 130)")
    ap.add_argument("--target-min", type=float, default=0,
                    help="runtime target in minutes; flags overruns + trims")
    ap.add_argument("--write-md", metavar="FILE",
                    help="also write a rehearsal cue-card sheet")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.deck):
        print(f"no such file: {a.deck}")
        return 2

    prs = Presentation(a.deck)
    rows, total_sec, total_notes = [], 0.0, 0
    problems = 0
    for i, slide in enumerate(prs.slides, 1):
        headline, on_words, notes, vocab = slide_facts(slide)
        n_words = len(notes.split())
        sec = n_words / a.wpm * 60
        total_sec += sec
        total_notes += n_words
        flags = []
        if n_words == 0:
            flags.append("NO SCRIPT")
            problems += 1
        else:
            if overlap(vocab, notes) > 0.55 and on_words >= 15:
                flags.append("READING RISK (script copies the slide)")
            if n_words < 40 and on_words > 0 and on_words / max(n_words, 1) > 0.6:
                flags.append("THIN SCRIPT (expand notes, not the slide)")
        if sec > 90:
            flags.append("LONG (>90s on one slide)")
        if on_words > 60:
            flags.append(f"DENSE ({on_words} words on slide)")
        rows.append(dict(i=i, headline=headline or f"Slide {i}", on=on_words,
                         nw=n_words, sec=sec, flags=flags, notes=notes))

    print(f"Speaker Coach — {os.path.basename(a.deck)} "
          f"({len(rows)} slides, {a.wpm:.0f} wpm)")
    print(f"{'#':>3}  {'time':>5}  {'script':>6}  {'slide':>5}  headline / flags")
    for r in rows:
        h = r["headline"][:44]
        fl = "  ⚠ " + "; ".join(r["flags"]) if r["flags"] else ""
        print(f"{r['i']:>3}  {fmt(r['sec']):>5}  {r['nw']:>5}w  {r['on']:>4}w  {h}{fl}")
    print(f"\nTotal scripted time ≈ {fmt(total_sec)} from {total_notes} words "
          f"of notes; unscripted slides add talk time on top.")
    if a.target_min:
        tgt = a.target_min * 60
        if total_sec > tgt:
            over = total_sec - tgt
            need = int(round(over / 60 * a.wpm))
            print(f"OVER TARGET by {fmt(over)} — cut ≈{need} scripted words. "
                  f"Longest slides first:")
            for r in sorted(rows, key=lambda r: -r["sec"])[:3]:
                print(f"  - slide {r['i']} ({fmt(r['sec'])}, {r['nw']}w): "
                      f"{r['headline'][:50]}")
            problems += 1
        elif total_sec < tgt * 0.7:
            print(f"UNDER TARGET ({fmt(total_sec)} of {fmt(tgt)}) — room to "
                  f"add a story, a demo beat, or slower pacing.")
        else:
            print(f"On target ({fmt(total_sec)} of {fmt(tgt)}). Two full "
                  f"rehearsals at this pace before you present.")
    unscripted = sum(1 for r in rows if r["nw"] == 0)
    if unscripted:
        print(f"{unscripted} slide(s) have no speaker notes — write the script "
              f"before rehearsal; notes are the load-bearing half (ref/12).")

    if a.write_md:
        lines = [f"# Rehearsal sheet — {os.path.basename(a.deck)}", "",
                 f"Scripted runtime ≈ {fmt(total_sec)} at {a.wpm:.0f} wpm. "
                 f"Practice the first sentence of each slide out loud; the "
                 f"rest follows.", ""]
        clock = 0.0
        for r in rows:
            lines.append(f"## {r['i']}. {r['headline']} "
                         f"— {fmt(r['sec'])} (at {fmt(clock)})")
            clock += r["sec"]
            if r["flags"]:
                lines.append(f"⚠ {'; '.join(r['flags'])}")
            lines.append("")
            lines.append(r["notes"] or "_(no script yet — write one)_")
            lines.append("")
            lines.append("→ transition: _(your one-line bridge to the next "
                         "slide)_")
            lines.append("")
        with open(a.write_md, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"[cue cards written to {a.write_md}]")
    return 1 if problems and a.target_min and total_sec > a.target_min * 60 else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
