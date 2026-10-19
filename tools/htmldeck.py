#!/usr/bin/env python3
"""HTML deck mode — any generator spec -> standalone playable HTML deck.

Semantic (not wireframe): each slide type becomes real HTML in the spec's
theme colors. One file, no assets, works offline. Present with arrow keys;
`o` toggles an overview grid, `n` toggles speaker notes. Ideal for sharing
a deck with people who don't have PowerPoint, embedding in docs/Notion,
or presenting from a phone browser. (For geometric proof sheets of an
existing .pptx, use tools/render.py.)

Usage:  python tools/htmldeck.py spec.json [--out deck.html]
"""
import argparse
import html
import json
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
import generate_pptx as G  # noqa: E402

INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")
INDIC_ON = False


def T(x):
    s = html.escape(str(x or ""))
    return s.translate(INDIC) if INDIC_ON else s


def bars_html(cats, vals, hi):
    vmax = max([abs(float(v)) for v in vals] + [1])
    rows = []
    for i, (c, v) in enumerate(zip(cats, vals)):
        w = abs(float(v)) / vmax * 100
        rows.append(
            f'<div class="bar-row"><span class="bar-cat">{T(c)}</span>'
            f'<span class="bar-track"><span class="bar-fill '
            f'{"hot" if i == hi else ""}" style="width:{w:.0f}%"></span></span>'
            f'<span class="bar-val">{T(f"{float(v):g}")}</span></div>')
    return '<div class="bars">' + "".join(rows) + "</div>"


def cards_html(items):
    out = []
    for it in items:
        out.append(
            f'<div class="card"><div class="chip">{T(G._icon_glyph(it))}</div>'
            f'<h3>{T(it.get("head") or it.get("title", ""))}</h3>'
            f'<p>{T(it.get("body", ""))}</p></div>')
    return '<div class="grid">' + "".join(out) + "</div>"


def steps_html(steps):
    out = []
    for i, st in enumerate(steps, 1):
        out.append(f'<div class="step"><span class="num">{i}</span><div>'
                   f'<h3>{T(st.get("head") or st.get("title", ""))}</h3>'
                   f'<p>{T(st.get("body", ""))}</p></div></div>')
    return '<div class="steps">' + "".join(out) + "</div>"


def bullets_html(items):
    out = []
    for b in items:
        out.append(f'<div class="bul"><h3>{T(b.get("head", ""))}</h3>'
                   f'<p>{T(b.get("body", ""))}</p></div>')
    return '<div class="bullets">' + "".join(out) + "</div>"


def slide_html(sd, rtl):
    t = sd.get("type", "content")
    k = f'<div class="kick">{T(sd["kicker"])}</div>' if sd.get("kicker") else ""
    title = f'<h2>{T(sd.get("title", ""))}</h2>' if sd.get("title") else ""
    notes = (f'<aside class="notes"><b>Speaker notes</b>'
             f'{T(sd["notes"])}</aside>') if sd.get("notes") else ""
    body = ""

    if t in ("title", "hook"):
        lines = "<br>".join(T(l) for l in
                             (sd.get("title_lines") or [sd.get("title", "")]))
        badge = f'<span class="badge">{T(sd["badge"])}</span>' if sd.get("badge") else ""
        body = (f'<div class="hero">{badge}<h1>{lines}</h1>'
                f'<p class="sub">{T(sd.get("subtitle", ""))}</p>'
                f'<p class="meta">{T(sd.get("meta", ""))}</p></div>')
    elif t == "section":
        body = (f'<div class="hero"><span class="secnum">{T(sd.get("number", ""))}</span>'
                f'<h1>{T(sd.get("title", ""))}</h1>'
                f'<p class="sub">{T(sd.get("subtitle", ""))}</p></div>')
    elif t == "agenda":
        lis = "".join(
            f'<li class="{"on" if i == int(sd.get("active", 0)) else ""}">'
            f'{T(x)}</li>' for i, x in enumerate(sd.get("items", [])))
        body = k + title + f'<ol class="agenda">{lis}</ol>'
    elif t in ("hero_stat", "bigstat"):
        body = (f'<div class="stat">{k}<div class="bignum">{T(sd.get("stat", ""))}</div>'
                f'<div class="statlabel">{T(sd.get("label", ""))}</div>'
                f'<p class="meta">{T(sd.get("context", ""))}</p></div>')
    elif t == "stat_strip":
        cells = "".join(
            f'<div class="statcell"><div class="bignum sm">{T(s.get("value",""))}</div>'
            f'<div class="statlabel">{T(s.get("label",""))}</div>'
            f'<p class="meta">{T(s.get("context",""))}</p></div>'
            for s in sd.get("stats", []))
        body = k + title + f'<div class="strip">{cells}</div>'
    elif t in ("chart", "bars", "donut", "line_chart", "waterfall", "funnel"):
        if t == "line_chart":
            srs = (sd.get("series") or [{}])[-1]
            cats, vals = sd.get("categories", []), srs.get("values", [])
        elif t in ("waterfall", "funnel"):
            items = sd.get("steps") or sd.get("stages") or []
            cats = [i.get("label", "") for i in items]
            vals = [abs(float(i.get("value", 0))) for i in items]
        else:
            cats, vals = sd.get("categories", []), sd.get("values", [])
        ins = (f'<div class="insight"><b>KEY INSIGHT</b>{T(sd["insight"])}</div>'
               if sd.get("insight") else "")
        body = k + title + bars_html(cats, vals, int(sd.get("highlight", 0))) + ins
    elif t in ("process", "timeline", "steps"):
        body = k + title + steps_html(sd.get("steps", []))
    elif t in ("icon_grid", "tips"):
        body = k + title + cards_html(sd.get("items", []))
    elif t == "content":
        vis = (f'<div class="visual"><div class="bignum sm">{T(sd["visual_big"])}</div>'
               f'<p class="meta">{T(sd.get("visual_caption", ""))}</p></div>'
               if sd.get("visual_big") else "")
        body = k + title + f'<div class="split"><div>{bullets_html(sd.get("bullets", []))}</div>{vis}</div>'
    elif t == "comparison":
        cols = ""
        for c in sd.get("columns", []):
            rows = "".join(f"<li>{T(str(r).lstrip('+- '))}</li>"
                           for r in c.get("rows", []))
            cols += (f'<div class="col {"hot-col" if c.get("highlight") else ""}">'
                     f'<h3>{T(c.get("head", ""))}</h3><ul>{rows}</ul></div>')
        body = k + title + f'<div class="cols">{cols}</div>'
    elif t in ("before_after", "compare"):
        cols = ""
        for key, cls in (("before", ""), ("after", "hot-col")):
            p = sd.get(key, {})
            rows = "".join(f"<li>{T(str(r).lstrip('+- '))}</li>"
                           for r in p.get("rows", []))
            cols += (f'<div class="col {cls}"><h3>{T(p.get("head", ""))}</h3>'
                     f'<ul>{rows}</ul></div>')
        body = k + title + f'<div class="cols">{cols}</div>'
    elif t in ("market", "tam_sam_som"):
        rows = "".join(
            f'<div class="ring"><span class="ringlab">{T(r.get("label",""))}</span>'
            f'<span class="ringval">{T(r.get("value",""))}</span>'
            f'<span class="meta">{T(r.get("note",""))}</span></div>'
            for r in sd.get("rings", []))
        body = k + title + f'<div class="rings">{rows}</div>'
    elif t == "team":
        cards = "".join(
            f'<div class="card"><div class="avatar">{T("".join(w[:1] for w in str(p.get("name","")).split()[:2]))}</div>'
            f'<h3>{T(p.get("name",""))}</h3><p class="role">{T(p.get("role",""))}</p>'
            f'<p>{T(p.get("bio",""))}</p></div>' for p in sd.get("people", []))
        body = k + title + f'<div class="grid">{cards}</div>'
    elif t == "quote":
        body = (f'<blockquote>“{T(sd.get("quote", ""))}”'
                f'<cite>— {T(sd.get("attribution", ""))}</cite></blockquote>')
    elif t == "logo_wall":
        names = "".join(f'<span class="logo">{T(n)}</span>'
                        for n in sd.get("names", []))
        body = k + title + f'<div class="wall">{names}</div><p class="meta">{T(sd.get("caption",""))}</p>'
    elif t == "matrix":
        items = "".join(
            f'<li class="{"hot" if i.get("highlight") else ""}">'
            f'{T(i.get("label", ""))} ({T(i.get("x", ""))}, {T(i.get("y", ""))})</li>'
            for i in sd.get("items", []))
        body = (k + title + f'<p class="meta">Quadrants: '
                f'{T(" · ".join(sd.get("quadrants", [])))}</p>'
                f'<ul class="mx">{items}</ul>'
                f'<div class="insight">{T(sd.get("insight", ""))}</div>')
    elif t == "bilingual":
        body = (k + f'<div class="cols">'
                f'<div class="col"><h3>{T(sd.get("en_title", ""))}</h3>'
                f'{bullets_html(sd.get("en_bullets", []))}'
                f'<p>{T(sd.get("en_body", ""))}</p></div>'
                f'<div class="col" dir="rtl"><h3>{T(sd.get("ar_title", ""))}</h3>'
                f'{bullets_html(sd.get("ar_bullets", []))}'
                f'<p>{T(sd.get("ar_body", ""))}</p></div></div>')
    elif t in ("closing", "cta"):
        ask = sd.get("ask") or " ".join(sd.get("title_lines", []) or
                                        [sd.get("title", "")])
        body = (f'<div class="hero"><h1>{T(ask)}</h1>'
                f'<p class="sub">{T(sd.get("subtitle", ""))}</p>'
                f'<p class="meta">{T(sd.get("contact") or sd.get("handle", ""))}</p></div>')
    else:
        body = k + title + f'<p>{T(sd.get("body") or sd.get("text", ""))}</p>'
    direction = ' dir="rtl"' if rtl else ""
    return f'<section class="slide"{direction}>{body}{notes}</section>'


CSS = """
*{box-sizing:border-box} html,body{margin:0;height:100%;background:#0c0e11}
body{font-family:Inter,Tajawal,system-ui,sans-serif;color:var(--ink)}
#deck{height:100vh;display:flex;align-items:center;justify-content:center}
.slide{display:none;width:min(96vw,150vh);aspect-ratio:16/9;background:var(--bg);
border-radius:14px;padding:4.5% 5.5%;position:relative;overflow:hidden;
box-shadow:0 18px 60px rgba(0,0,0,.55)}
.slide.on{display:block}
h1{font-family:Montserrat,Cairo,sans-serif;font-size:clamp(26px,5.2vw,64px);
line-height:1.08;margin:.25em 0}
h2{font-family:Montserrat,Cairo,sans-serif;font-size:clamp(20px,3vw,38px);margin:.1em 0 .6em}
h3{font-family:Montserrat,Cairo,sans-serif;font-size:clamp(13px,1.5vw,20px);margin:.2em 0}
p{font-size:clamp(12px,1.35vw,18px);line-height:1.45;color:var(--ink)}
.kick{color:var(--accent);font-weight:700;letter-spacing:.14em;font-size:clamp(10px,1.1vw,14px)}
.hero{height:100%;display:flex;flex-direction:column;justify-content:center}
.sub{font-size:clamp(16px,2.2vw,28px);color:var(--muted)}
.meta{color:var(--muted);font-size:clamp(11px,1.2vw,15px)}
.badge{background:var(--accent);color:var(--bg);border-radius:999px;padding:.35em 1em;
font-weight:700;font-size:clamp(10px,1.1vw,14px);width:max-content}
.secnum{font-size:clamp(40px,7vw,90px);color:var(--accent);font-weight:800;
font-family:Montserrat,Cairo,sans-serif}
.agenda{font-size:clamp(16px,2vw,26px);line-height:1.9}
.agenda .on{color:var(--accent);font-weight:800}
.stat{height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center}
.bignum{font-family:Montserrat,Cairo,sans-serif;font-weight:800;color:var(--accent);
font-size:clamp(60px,11vw,150px);line-height:1}
.bignum.sm{font-size:clamp(34px,5vw,64px)}
.statlabel{font-size:clamp(16px,2vw,26px);font-weight:700}
.strip{display:flex;gap:2%}
.statcell{background:var(--card);border-radius:12px;padding:18px;flex:1;text-align:center}
.bars{margin-top:2%}
.bar-row{display:flex;align-items:center;gap:12px;margin:10px 0}
.bar-cat{width:22%;font-weight:600;font-size:clamp(12px,1.3vw,17px)}
.bar-track{flex:1;background:var(--graybar);border-radius:99px;height:20px;overflow:hidden;opacity:.9}
.bar-fill{display:block;height:100%;background:#9aa7ad;border-radius:99px}
.bar-fill.hot{background:var(--accent)}
.bar-val{width:8%;font-weight:700}
.insight{background:var(--card);border-radius:12px;padding:14px 18px;margin-top:3%;
font-size:clamp(12px,1.4vw,18px)}
.insight b{display:block;color:var(--accent);font-size:.78em;letter-spacing:.12em}
.steps{display:flex;flex-direction:column;gap:14px}
.step{display:flex;gap:14px;align-items:flex-start}
.step .num{background:var(--accent);color:var(--bg);width:38px;height:38px;flex:0 0 38px;
border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px}
.card{background:var(--card);border-radius:14px;padding:18px}
.chip{background:var(--accent);color:var(--bg);width:40px;height:40px;border-radius:50%;
display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:800}
.avatar{background:var(--accent);color:var(--bg);width:52px;height:52px;border-radius:50%;
display:flex;align-items:center;justify-content:center;font-weight:800;font-size:20px}
.role{color:var(--accent);font-weight:700}
.split{display:grid;grid-template-columns:1.6fr 1fr;gap:4%}
.visual{background:var(--card);border-radius:14px;display:flex;flex-direction:column;
align-items:center;justify-content:center;padding:18px;text-align:center}
.bul{margin-bottom:14px}
.cols{display:flex;gap:2%}
.col{background:var(--card);border-radius:14px;padding:20px;flex:1}
.col.hot-col{background:var(--accent);color:var(--bg)}
.col.hot-col p,.col.hot-col li{color:var(--bg)}
.ring{display:flex;gap:16px;align-items:center;background:var(--card);border-radius:12px;
padding:14px 18px;margin:10px 0}
.ringlab{font-weight:800;color:var(--accent)}
.ringval{font-family:Montserrat,Cairo,sans-serif;font-weight:800;font-size:clamp(20px,2.6vw,34px)}
blockquote{font-family:Montserrat,Cairo,sans-serif;font-size:clamp(20px,3.2vw,40px);
font-weight:700;line-height:1.3;margin:8% 4%}
blockquote cite{display:block;margin-top:1em;font-size:.5em;color:var(--muted);font-style:normal}
.wall{display:flex;flex-wrap:wrap;gap:14px}
.logo{background:var(--card);border-radius:10px;padding:14px 22px;font-weight:800;
font-family:Montserrat,Cairo,sans-serif}
.mx li.hot{color:var(--accent);font-weight:800}
.notes{display:none;position:absolute;inset-inline:0;bottom:0;background:rgba(0,0,0,.82);
color:#fff;padding:14px 20px;font-size:14px;line-height:1.5}
body.shownotes .notes{display:block}
.notes b{display:block;color:var(--accent);font-size:11px;letter-spacing:.12em}
#hud{position:fixed;bottom:14px;inset-inline:0;display:flex;justify-content:center;gap:14px;
color:#9aa0a6;font-size:12px;font-family:Inter,system-ui}
#prog{position:fixed;top:0;left:0;height:3px;background:var(--accent);transition:width .25s}
body.overview #deck{display:block;height:auto;padding:20px}
body.overview .slide{display:block;width:640px;max-width:92vw;margin:0 auto 22px}
@media (prefers-color-scheme: light){ html,body{background:#dfe3e6} }
"""

JS = """
const slides=[...document.querySelectorAll('.slide')];let i=0;
function show(n){i=Math.max(0,Math.min(slides.length-1,n));
slides.forEach((s,k)=>s.classList.toggle('on',k===i));
document.getElementById('count').textContent=(i+1)+' / '+slides.length;
document.getElementById('prog').style.width=((i+1)/slides.length*100)+'%';}
addEventListener('keydown',e=>{if(e.key==='ArrowRight'||e.key===' ')show(i+1);
else if(e.key==='ArrowLeft')show(i-1);else if(e.key==='Home')show(0);
else if(e.key==='End')show(slides.length-1);
else if(e.key==='o'||e.key==='O')document.body.classList.toggle('overview');
else if(e.key==='n'||e.key==='N')document.body.classList.toggle('shownotes');});
addEventListener('click',e=>{if(!document.body.classList.contains('overview'))
show(i+(e.clientX>innerWidth/2?1:-1))});
show(0);
"""


def main(argv):
    global INDIC_ON
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("spec")
    ap.add_argument("--out", help="output HTML (default <spec>.html)")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.spec):
        print(f"no such file: {a.spec}")
        return 2
    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)
    t = dict(G.THEMES.get(spec.get("theme", "teal-orbit"), G.THEMES["teal-orbit"]))
    t.update(spec.get("theme_overrides", {}))
    rtl = bool(spec.get("rtl"))
    INDIC_ON = str(spec.get("numerals", "")).lower() == "arabic-indic"
    sections = []
    for sd in spec.get("slides", []):
        sections.append(slide_html(G._expand_tokens(sd, spec), rtl))
    vars_css = "".join(f"--{k}:#{t[k]};" for k in
                       ("bg", "card", "ink", "muted", "accent", "graybar"))
    title = html.escape(spec.get("title") or spec.get("brand") or "Deck")
    doc = f"""<!doctype html><html lang="{'ar' if rtl else 'en'}"{' dir="rtl"' if rtl else ''}>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><style>:root{{{vars_css}}}{CSS}</style></head>
<body><div id="prog"></div><main id="deck">{''.join(sections)}</main>
<div id="hud"><span id="count"></span><span>← → navigate · <b>o</b> overview · <b>n</b> notes · click halves to move</span>
<span>{html.escape(spec.get('brand', ''))}</span></div>
<script>{JS}</script></body></html>
"""
    out = a.out or os.path.splitext(a.spec)[0] + ".html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    print(f"html deck: {out} ({len(sections)} slides, {len(doc)//1024} KB, "
          f"theme={spec.get('theme', 'teal-orbit')})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
