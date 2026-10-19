#!/usr/bin/env python3
"""Project site — generates site/index.html (single file, no assets).

Pulls live numbers from the repo itself: Monster Database census, theme
gallery (all palettes with swatches), slide-type catalog, tools grid,
example decks, reference word counts. Also emits a playable browser deck
(site/demo.html via htmldeck) and a contact sheet (site/demo.contact.html
via render) so visitors can *see* the system before reading about it.

Usage:  python tools/site.py [--outdir site]
Serve:  python -m http.server -d site 8080
"""
import argparse
import html
import os
import sqlite3
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
sys.path.insert(0, os.path.join(ROOT, "generator"))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import generate_pptx as G   # noqa: E402
import portrait             # noqa: E402
import htmldeck             # noqa: E402
import render               # noqa: E402

TOOLS = [
    ("qa_deck.py", "QA gate v2 — Morph map, overflow sentinel, WCAG contrast"),
    ("wizard.py", "Creator Mode — goal → DB structure → QA'd starter deck"),
    ("draft.py", "Raw notes → deck spec (prose becomes speaker notes)"),
    ("coach.py", "Speaker Coach — talk time, reading-risk, cue cards"),
    ("social.py", "Social Factory — Reel 9:16 + Carousel 4:5 + posting kit"),
    ("palette_lab.py", "Seed color → contrast-safe theme"),
    ("brandkit.py", "Logo image → theme"),
    ("anatomy.py", "Any .pptx → report + reusable theme"),
    ("render.py", "Contact-sheet proof renderer (geometry wireframes)"),
    ("htmldeck.py", "Spec → playable HTML deck (arrows, overview, notes)"),
    ("diff.py", "Deck A/B changelog"),
    ("db_lint.py", "Data linter — catches rot, drift, stale DB"),
    ("hijri.py", "Gregorian → Hijri converter"),
    ("site.py", "This site"),
]


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--outdir", default=os.path.join(ROOT, "site"))
    a = ap.parse_args(argv[1:])
    os.makedirs(a.outdir, exist_ok=True)

    # playable demo + contact sheet
    demo_spec = os.path.join(ROOT, "generator", "deck-spec.example.json")
    htmldeck.main(["htmldeck", demo_spec, "--out",
                   os.path.join(a.outdir, "demo.html")])
    render.main(["render", os.path.join(ROOT, "examples",
                                        "placeplate-style-demo.pptx"),
                 "--out", os.path.join(a.outdir, "demo.contact.html")])

    # live census
    db = os.path.join(ROOT, "database", "powerpoint.db")
    census, total_records = [], 0
    if os.path.exists(db):
        c = sqlite3.connect(db)
        tables = [r[0] for r in c.execute(
            "SELECT name FROM sqlite_master WHERE type='table'")]
        tables = [t for t in tables if not t.startswith(("search_index",
                                                          "sqlite_"))]
        for t in sorted(tables):
            n = c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            census.append((t, n))
            total_records += n

    refdir = os.path.join(ROOT, "reference")
    refs, words = [], 0
    for fn in sorted(os.listdir(refdir)):
        if fn.endswith(".md"):
            w = len(open(os.path.join(refdir, fn), encoding="utf-8")
                    .read().split())
            refs.append((fn, w))
            words += w

    examples = []
    exdir = os.path.join(ROOT, "examples")
    from pptx import Presentation
    for fn in sorted(os.listdir(exdir)):
        if fn.endswith(".pptx"):
            p = os.path.join(exdir, fn)
            try:
                n = len(Presentation(p).slides)
            except Exception:
                n = "?"
            examples.append((fn, n, os.path.getsize(p) // 1024))

    def swatches():
        cards = []
        for name, t in sorted(G.THEMES.items()):
            chips = "".join(
                f'<span class="chip" style="background:#{t[k]}" '
                f'title="{k} #{t[k]}"></span>'
                for k in ("bg", "card", "accent", "accent2", "ink"))
            cards.append(
                f'<div class="pal" style="background:#{t["bg"]};color:#{t["ink"]}">'
                f'<div class="palname">{html.escape(name)} '
                f'<small>{"dark" if t["dark"] else "light"}</small></div>'
                f'<div class="chips">{chips}</div></div>')
        return "".join(cards)

    census_rows = "".join(f"<tr><td>{html.escape(t)}</td><td>{n}</td></tr>"
                          for t, n in census)
    ref_rows = "".join(f"<tr><td>{html.escape(fn)}</td><td>{w:,}</td></tr>"
                       for fn, w in refs)
    ex_rows = "".join(f"<tr><td>{html.escape(fn)}</td><td>{n} slides</td>"
                      f"<td>{kb} KB</td></tr>" for fn, n, kb in examples)
    tool_cards = "".join(
        f'<div class="tool"><code>{html.escape(n)}</code><p>{html.escape(d)}</p></div>'
        for n, d in TOOLS)
    land = ", ".join(f"<code>{html.escape(b)}</code>" for b in sorted(G.BUILDERS))
    port = ", ".join(f"<code>{html.escape(b)}</code>" for b in sorted(portrait.BUILDERS))

    doc = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PowerPoint Skill — PlacePlate-style Morph system</title><style>
*{{box-sizing:border-box}} body{{margin:0;background:#0b0f12;color:#e9f2f0;
font-family:Inter,system-ui,sans-serif}} .wrap{{max-width:1060px;margin:0 auto;padding:28px 20px 60px}}
h1{{font-family:Montserrat,system-ui;font-size:clamp(30px,5vw,52px);line-height:1.05;margin:.2em 0}}
h2{{font-family:Montserrat,system-ui;margin-top:52px;font-size:26px}}
.accent{{color:#14C8B4}} .sub{{color:#9fc0bb;font-size:18px;line-height:1.55;max-width:760px}}
.stats{{display:flex;flex-wrap:wrap;gap:10px;margin:22px 0}}
.stat{{background:#12242a;border:1px solid #1d3a42;border-radius:12px;padding:12px 16px}}
.stat b{{display:block;font-size:22px;color:#14C8B4}} .stat span{{color:#9fc0bb;font-size:12px}}
a.btn{{display:inline-block;background:#14C8B4;color:#04333F;font-weight:800;border-radius:999px;
padding:12px 22px;text-decoration:none;margin:6px 8px 6px 0}}
a.ghost{{background:transparent;color:#14C8B4;border:1px solid #14C8B4}}
.pals{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px}}
.pal{{border-radius:12px;padding:14px;border:1px solid #26343a}}
.palname{{font-weight:700;margin-bottom:10px}} .palname small{{opacity:.65;font-weight:400}}
.chip{{display:inline-block;width:26px;height:26px;border-radius:7px;margin-right:6px;
border:1px solid rgba(127,127,127,.35)}}
.tools{{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px}}
.tool{{background:#101a1f;border:1px solid #1d3a42;border-radius:12px;padding:14px}}
.tool p{{color:#9fc0bb;font-size:13px;margin:.5em 0 0}}
code{{background:#16262c;border-radius:6px;padding:2px 7px;font-size:.86em;color:#7fe7d2}}
table{{border-collapse:collapse;min-width:340px}} td{{border-bottom:1px solid #1d3a42;padding:6px 14px 6px 0;color:#cfe0dd}}
.cols{{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:26px}}
pre{{background:#0e171b;border:1px solid #1d3a42;border-radius:12px;padding:16px;overflow:auto;color:#bfe}}
footer{{margin-top:60px;color:#6f8a86;font-size:13px;line-height:1.6}}
</style></head><body><div class="wrap">

<p class="accent" style="font-weight:800;letter-spacing:.14em">POWERPOINT SKILL · PLACEPLATE-STYLE</p>
<h1>Presentations that <span class="accent">move</span>,<br>built from data — not vibes.</h1>
<p class="sub">A complete Morph-first presentation system: a research-backed
reference library, a JSON→.pptx generator with {len(G.BUILDERS)} landscape and
{len(portrait.BUILDERS)} portrait slide builders, a {total_records}-record queryable
knowledge database, an automated QA gate, and a creator toolchain — in English
and Arabic (RTL), for the boardroom and for Reels.</p>

<div class="stats">
  <div class="stat"><b>{total_records}</b><span>DB records · {len(census)} tables + FTS5</span></div>
  <div class="stat"><b>{len(G.BUILDERS)} + {len(portrait.BUILDERS)}</b><span>slide builders (16:9 + portrait)</span></div>
  <div class="stat"><b>{len(G.THEMES)}</b><span>themes, all WCAG-checked</span></div>
  <div class="stat"><b>{len(TOOLS)}</b><span>tools in tools/</span></div>
  <div class="stat"><b>{words:,}</b><span>words of reference docs</span></div>
  <div class="stat"><b>18</b><span>automated tests + CI</span></div>
</div>

<a class="btn" href="demo.html">▶ Play the demo deck</a>
<a class="btn ghost" href="demo.contact.html">Contact sheet</a>

<h2>Try it</h2>
<pre>python tools/wizard.py --goal pitch --topic "Loopwise" --claim "Turn cafe waste into packaging"
python tools/draft.py my-notes.txt --build        # raw notes → deck
python generator/generate_pptx.py spec.json deck.pptx
python tools/qa_deck.py deck.pptx                  # must print GREEN
python tools/social.py spec.json --handle @you     # Reel + Carousel pack
python database/query.py recommend pitch           # structures from the DB</pre>

<h2>Themes ({len(G.THEMES)})</h2>
<p class="sub">Every palette passes contrast floors (ink 10:1 · muted 4.5:1 ·
accent 3:1). Make your own: <code>palette_lab.py --seed FF6B4A --name mine</code>
or from a logo: <code>brandkit.py logo.png --name brand</code>.</p>
<div class="pals">{swatches()}</div>

<h2>Slide catalog</h2>
<p><b>Landscape 16:9:</b> {land}</p>
<p><b>Portrait (Reel 9:16 · Carousel 4:5):</b> {port}</p>

<h2>Tools</h2>
<div class="tools">{tool_cards}</div>

<div class="cols">
<div><h2>Monster Database census</h2><table>{census_rows}</table></div>
<div><h2>Example decks (all QA-green)</h2><table>{ex_rows}</table></div>
</div>

<h2>Reference library</h2>
<table>{ref_rows}</table>

<footer>Generated by <code>tools/site.py</code> from live repo data ·
Morph-first · assertion headlines · one accent at a time · notes carry the script.
Arabic/RTL, bilingual slides, Arabic-Indic numerals and Hijri dates supported end-to-end.</footer>
</div></body></html>
"""
    out = os.path.join(a.outdir, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    print(f"site: {out} ({len(doc)//1024} KB) + demo.html + demo.contact.html")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
