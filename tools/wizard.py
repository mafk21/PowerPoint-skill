#!/usr/bin/env python3
"""Deck Wizard — answer a few questions, get a QA'd Morph deck.

Looks the goal up in the Monster Database (structures table: recipe
sequence + suggested themes), translates recipes into generator slide
types, fills each slide with structured starter copy woven from your
answers plus coaching notes, builds the deck, and runs the QA gate.

Usage (flags = non-interactive):
  python tools/wizard.py --goal pitch --topic "Loopwise" \
      --claim "Turn cafe waste into packaging" --stat "87%" \
      --stat-label "of pilot cafes renewed" --ask "Raising $450K" \
      --outdir /tmp/Loopwise
  python tools/wizard.py --interactive          # prompt me instead
"""
import argparse
import json
import os
import re
import sqlite3
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(ROOT, "..", "database", "powerpoint.db")

GOAL_KEYWORDS = {
    "pitch": "pitch", "investor": "pitch", "fundrais": "pitch",
    "thesis": "thesis", "academic": "thesis", "defense": "thesis",
    "business": "business", "sales": "sales", "seminar": "seminar",
    "qbr": "quarterly", "quarterly": "quarterly", "decision": "decision",
}

LABEL_TYPES = [
    ("title", "title"), ("outline", "agenda"), ("agenda", "agenda"),
    ("ask", "closing"), ("closing", "closing"), ("conclusion", "content"),
    ("market", "market"), ("tam", "tam_sam_som"), ("team", "team"),
    ("traction", "chart"), ("financial", "chart"), ("result", "chart"),
    ("data", "chart"), ("metric", "stat_strip"), ("kpi", "stat_strip"),
    ("method", "process"), ("how it works", "process"), ("gtm", "process"),
    ("go-to-market", "process"), ("roadmap", "timeline"),
    ("competition", "matrix"), ("competitive", "matrix"),
    ("comparison", "comparison"), ("vs", "comparison"),
    ("quote", "quote"), ("testimonial", "quote"), ("vision", "hero_stat"),
    ("problem", "content"), ("solution", "content"), ("product", "content"),
    ("intro", "content"), ("question", "icon_grid"), ("model", "icon_grid"),
    ("pricing", "comparison"), ("reference", "content"),
]

FALLBACK = ("Investor pitch deck",
            "R1 title > R4 problem > R4 solution > R10 market > R8 traction > "
            "R11 team > R13 ask", "teal-orbit, steel-blue")


def ask(q, default):
    try:
        v = input(f"{q} [{default}]: ").strip()
    except EOFError:
        return default
    return v or default


def find_structure(goal):
    key = GOAL_KEYWORDS.get(goal.lower(), goal.lower())
    try:
        c = sqlite3.connect(DB)
        rows = c.execute("SELECT name, sequence, suggested_themes "
                         "FROM structures").fetchall()
        for name, seq, themes in rows:
            if key in name.lower() or goal.lower() in name.lower():
                return name, seq, themes
    except Exception:
        pass
    return FALLBACK


def parse_sequence(seq):
    """'R1 title > R4 problem > R8/R9 results x2-3 > backup' -> [(type,label)]"""
    out = []
    for tok in re.split(r"\s*>\s*", seq or ""):
        tok = tok.strip()
        if not tok or "backup" in tok.lower():
            continue
        code = re.match(r"^R(\d+)", tok)
        label = re.sub(r"^R\d+(?:/R?\d+)?\s*", "", tok)
        label = re.sub(r"\(.*?\)", "", label).strip()
        if not label and code:  # bare recipe codes: R1 title, R2 agenda, R13 closing
            label = {"1": "title", "2": "outline", "13": "ask"}.get(
                code.group(1), f"section {code.group(1)}")
        m = re.search(r"\bx\s*(\d)(?:-\d)?\s*$", label)
        reps = int(m.group(1)) if m else 1
        if m:
            label = label[:m.start()].strip()
        stype = "content"
        low = label.lower()
        for kw, tname in LABEL_TYPES:
            if kw in low:
                stype = tname
                break
        for _ in range(reps):
            out.append((stype, label))
    return out


def starter(stype, label, a):
    """Structured placeholder copy per type — replace-values, not lorem."""
    T, C = a.topic, a.claim
    title = f"{label.capitalize()}: state the finding, not the topic"
    note = (f"[{label}] Open with the takeaway sentence. One idea only — "
            f"if you're saying 'and also', split the slide (ref/12).")
    if stype == "title":
        return dict(type="title", badge="START HERE", title_lines=[T],
                    subtitle=C, meta=a.meta,
                    notes="Say the claim out loud before anything else. "
                          "Big Idea = viewpoint + what's at stake (ref/12).")
    if stype == "agenda":
        return dict(type="agenda", title="Where this goes", items=a._agenda,
                    active=0, notes="Promise the route in 10 seconds; the "
                                    "agenda returns as a progress tracker.")
    if stype == "chart":
        return dict(type="chart", kicker=label.upper(), title=title,
                    categories=["Q1", "Q2", "Q3", "Q4"], values=[10, 16, 22, 31],
                    highlight=3, series_name="Your metric",
                    insight="Replace with your real numbers — the insight "
                            "card states what the data means, not what it is.",
                    notes=note)
    if stype == "hero_stat":
        return dict(type="hero_stat", stat=a.stat or "3x",
                    label=a.stat_label or "your proof point",
                    context="One number, said slowly. Let it land (ref/03).",
                    notes=note)
    if stype == "stat_strip":
        return dict(type="stat_strip", kicker=label.upper(),
                    title="The scoreboard so far",
                    stats=[dict(value=a.stat or "3x",
                                label=a.stat_label or "headline metric",
                                context="vs last period"),
                           dict(value="[N]", label="second metric",
                                context="replace me"),
                           dict(value="[N]", label="third metric",
                                context="replace me")], notes=note)
    if stype in ("market", "tam_sam_som"):
        return dict(type="tam_sam_som", kicker="MARKET",
                    title="Big enough to matter, focused enough to win",
                    rings=[dict(label="TAM", value="$[X]B",
                                note="total market, top-down OK here"),
                           dict(label="SAM", value="$[X]M",
                                note="who you can actually reach"),
                           dict(label="SOM", value="$[X]M",
                                note="bottom-up: customers × price")],
                    notes="SOM must be bottom-up or investors discount "
                          "the whole slide (ref/12).")
    if stype == "team":
        return dict(type="team", kicker="TEAM",
                    title="Operators who have done this before",
                    people=[dict(name="[Founder]", role="CEO",
                                 bio="The one credential that de-risks this."),
                            dict(name="[Founder]", role="CTO / Product",
                                 bio="Built [relevant thing] at [place].")],
                    notes=note)
    if stype == "process":
        return dict(type="process", kicker=label.upper(), title=title,
                    steps=[dict(head="Step one",
                                body="What happens, in ≤8 words."),
                           dict(head="Step two",
                                body="What happens, in ≤8 words."),
                           dict(head="Step three",
                                body="What happens, in ≤8 words.")],
                    notes=note)
    if stype == "icon_grid":
        return dict(type="icon_grid", kicker=label.upper(), title=title,
                    items=[dict(icon="idea", head="Point one", body="Support it in one line."),
                           dict(icon="target", head="Point two", body="Support it in one line."),
                           dict(icon="star", head="Point three", body="Support it in one line.")],
                    notes=note)
    if stype == "matrix":
        return dict(type="matrix", kicker="POSITION",
                    title="We win the quadrant that matters",
                    quadrants=["Niche", "Leaders", "Laggards", "Loud & shallow"],
                    x_label="REACH →",
                    items=[dict(label=T, x=0.8, y=0.78, highlight=True),
                           dict(label="[Rival]", x=0.6, y=0.4),
                           dict(label="[Rival]", x=0.3, y=0.82)],
                    insight="Name the axes so the top-right is obviously "
                            "where value lives — then be there.",
                    notes=note)
    if stype == "comparison":
        return dict(type="comparison", kicker=label.upper(), title=title,
                    columns=[dict(head="Status quo", highlight=False,
                                  rows=["- Slow", "- Manual", "- Costly"]),
                             dict(head=T, highlight=True,
                                  rows=["+ Fast", "+ Automated", "+ Fair price"])],
                    notes=note)
    if stype == "timeline":
        return dict(type="timeline", kicker=label.upper(), title=title,
                    steps=[dict(head="Now", body="Where you are."),
                           dict(head="Next", body="The committed milestone."),
                           dict(head="Later", body="What it unlocks.")],
                    notes=note)
    if stype == "quote":
        return dict(type="quote",
                    quote="Replace with a real customer's sentence.",
                    attribution="[Name], [role], [company]", notes=note)
    if stype == "closing":
        return dict(type="closing",
                    title_lines=[a.ask or "Join us"],
                    subtitle=C,
                    contact=a.contact,
                    notes="End on the ask + the STAR moment, not 'thank you' "
                          "— end with what they should feel/do (ref/12).")
    return dict(type="content", kicker=label.upper(), title=title,
                bullets=[dict(head="First support",
                              body="One line of evidence, not a paragraph."),
                         dict(head="Second support",
                              body="Each row earns its place or gets cut."),
                         dict(head="So what",
                              body="The implication for this audience.")],
                visual_big="1 idea", visual_caption="per slide", notes=note)


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--goal", default="pitch",
                    help="pitch|thesis|business|sales|seminar|qbr|decision")
    ap.add_argument("--topic", default="My Project")
    ap.add_argument("--claim", default="One sentence: what changes because of us")
    ap.add_argument("--brand", default="")
    ap.add_argument("--theme", default="", help="default: DB-recommended")
    ap.add_argument("--stat", default="")
    ap.add_argument("--stat-label", default="")
    ap.add_argument("--ask", default="")
    ap.add_argument("--contact", default="")
    ap.add_argument("--meta", default="")
    ap.add_argument("--outdir", default="wizard-out")
    ap.add_argument("--interactive", action="store_true")
    ap.add_argument("--no-qa", action="store_true")
    a = ap.parse_args(argv[1:])

    if a.interactive:
        a.goal = ask("Goal (pitch/thesis/business/sales/seminar/qbr/decision)", a.goal)
        a.topic = ask("Topic / company name", a.topic)
        a.claim = ask("One-line claim (the Big Idea)", a.claim)
        a.stat = ask("Hero stat (e.g. 87%, $1.2M, 3x)", a.stat or "3x")
        a.stat_label = ask("What the stat means", a.stat_label or "our proof point")
        a.ask = ask("The ask (closing line)", a.ask or "Join us")
        a.brand = ask("Footer brand text", a.brand or a.topic.upper())
        a.theme = ask("Theme (blank = DB recommendation)", a.theme)
    if not a.brand:
        a.brand = a.topic.upper()

    name, seq, themes = find_structure(a.goal)
    theme = a.theme or (themes.split(",")[0].strip() if themes else "teal-orbit")
    outline = parse_sequence(seq)
    a._agenda = [lbl.capitalize() for t, lbl in outline
                 if t not in ("title", "agenda", "closing")][:6]
    slides = [starter(t, lbl, a) for t, lbl in outline]
    spec = {"theme": theme, "brand": a.brand, "timings": True, "slides": slides}

    os.makedirs(a.outdir, exist_ok=True)
    spec_path = os.path.join(a.outdir, "spec.json")
    deck_path = os.path.join(a.outdir, "deck.pptx")
    with open(spec_path, "w", encoding="utf-8") as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)

    print(f"Wizard — {name}")
    print(f"theme: {theme} (from DB: {themes})")
    for i, (t, lbl) in enumerate(outline, 1):
        print(f"  {i:2d}. {t:12s} {lbl}")

    sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
    import generate_pptx
    generate_pptx.build(spec, deck_path)
    print(f"\nbuilt {deck_path} ({len(slides)} slides) + {spec_path}")
    if not a.no_qa:
        r = subprocess.run([sys.executable,
                            os.path.join(ROOT, "qa_deck.py"), deck_path],
                           capture_output=True, text=True)
        tail = (r.stdout or "").strip().splitlines()[-1:]
        print("QA:", tail[0] if tail else "(qa failed to run)")
        print(f"full report: python tools/qa_deck.py {deck_path}")
    print(f"next: edit {spec_path} (replace [brackets]), rebuild with "
          f"python generator/generate_pptx.py {spec_path} {deck_path}, "
          f"rehearse with python tools/coach.py {deck_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
