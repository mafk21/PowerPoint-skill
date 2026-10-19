#!/usr/bin/env python3
"""Notes-to-Deck Drafter — raw notes text -> generator spec (+ optional build).

Heuristic, deterministic, no AI: it recognizes the shapes speakers naturally
write (headings, bullet lists, number series, quotes, steps, asks) and maps
each block to the right slide type. The original prose is preserved as
speaker notes — the slide carries the assertion, the script goes where
scripts belong (reference/10: redundancy effect).

Conventions it understands:
  # Heading            -> slide title (or kicker if a longer sentence follows)
  - bullet / * bullet  -> content rows ("Head: body" or "Head — body" splits)
  Label: 12            -> chart data (>=3 such lines = bar chart, max highlighted)
  42%  /  $1.2M        -> hero stat line, next sentence = its context
  > quoted text        -> quote slide (add "— Name" line for attribution)
  1. step 2. step      -> process (3-5 numbered lines)
  last block with ask/raising/contact -> closing

Usage:
  python tools/draft.py notes.txt --brand ACME --theme teal-orbit
  python tools/draft.py notes.txt --build --out spec.json --pptx deck.pptx
"""
import argparse
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
ICON_WORDS = {
    "growth": ("grow", "scale", "revenue", "sales", "traction", "expand"),
    "bolt": ("fast", "speed", "instant", "quick", "performance"),
    "shield": ("secur", "trust", "safe", "privacy", "complian"),
    "money": ("price", "cost", "margin", "profit", "monetiz", "pay"),
    "time": ("time", "schedule", "deadline", "deliver"),
    "globe": ("market", "global", "region", "world", "international"),
    "idea": ("idea", "insight", "research", "discover", "learn"),
    "gear": ("automat", "platform", "engine", "workflow", "process"),
    "star": ("quality", "best", "award", "premium", "leader"),
    "target": ("goal", "focus", "target", "objective"),
    "heart": ("love", "customer", "user", "community", "people"),
    "layers": ("stack", "layer", "infrastructure", "data"),
}

NUM = r"[-+]?\$?\d[\d,]*\.?\d*\s*(?:%|x|k|m|bn|b|m)?"
DATA_LINE = re.compile(r"^[-*•]?\s*([^:=]{2,28})[:=]\s*\$?(\d[\d,]*\.?\d*)\s*(%|k|m|bn)?\s*$", re.I)
STAT_LINE = re.compile(rf"^\s*({NUM})\s*$", re.I)
STEP_LINE = re.compile(r"^\s*(?:step\s*)?(\d+)[.)]\s+(.*)$", re.I)


def split_blocks(text):
    blocks, cur = [], []
    for line in text.splitlines():
        if line.strip() == "" or line.startswith("# "):
            if cur:
                blocks.append(cur)
                cur = []
            if line.startswith("# "):
                cur = [line]
        else:
            cur.append(line)
    if cur:
        blocks.append(cur)
    return [[l.rstrip() for l in b] for b in blocks if any(l.strip() for l in b)]


def icon_for(text):
    t = text.lower()
    for icon, keys in ICON_WORDS.items():
        if any(k in t for k in keys):
            return icon
    return "diamond"


def split_head_body(s):
    for sep in (" — ", " – ", ": ", " - "):
        if sep in s:
            h, b = s.split(sep, 1)
            if 1 <= len(h.split()) <= 6:
                return h.strip(), b.strip()
    words = s.split()
    if len(words) > 7:
        return " ".join(words[:4]), s
    return s.strip(), ""


def first_sentence(lines):
    txt = " ".join(l.strip() for l in lines)
    m = re.split(r"(?<=[.!?])\s+", txt)
    return m[0] if m else txt


def classify(lines, is_first, is_last):
    """Return (slide_type, fields_dict, reason)."""
    title = None
    if lines and lines[0].startswith("# "):
        title = lines[0][2:].strip()
        lines = lines[1:]
    body = [l.strip() for l in lines if l.strip()]
    joined = " ".join(body)

    if is_first and (title or len(body) <= 3):
        tl = [title] if title else ([body[0]] if body else ["Untitled"])
        rest = body[1:] if not title else body
        return "title", dict(title_lines=tl,
                             subtitle=rest[0] if rest else "",
                             meta=rest[1] if len(rest) > 1 else ""), \
            "first block"

    if body and body[0].startswith(">"):
        q = " ".join(l.lstrip("> ").strip() for l in body
                     if not l.startswith("—") and not l.startswith("--"))
        attr = next((l.lstrip("—- ").strip() for l in body
                     if l.startswith("—") or l.startswith("--")), "")
        return "quote", dict(quote=q.strip('“”"'), attribution=attr), "'> ' quote"

    data = [DATA_LINE.match(l) for l in body]
    data = [m for m in data if m]
    if len(data) >= 3:
        cats = [m.group(1).strip() for m in data]
        vals = [float(m.group(2).replace(",", "")) for m in data]
        unit = next((m.group(3) for m in data if m.group(3)), "")
        return "chart", dict(categories=cats, values=vals,
                             highlight=vals.index(max(vals)),
                             series_name=(title or "Value") +
                             (f" ({unit})" if unit else ""),
                             insight=first_sentence([l for l in body
                                                     if not DATA_LINE.match(l)])
                             if any(not DATA_LINE.match(l) for l in body) else ""), \
            f"{len(data)} label:value lines"

    steps = [STEP_LINE.match(l) for l in body]
    steps = [m for m in steps if m]
    if 3 <= len(steps) <= 5:
        return "process", dict(steps=[dict(zip(("head", "body"),
                                               split_head_body(m.group(2))))
                                      for m in steps]), "numbered steps"

    if body and STAT_LINE.match(body[0]) and len(body) >= 2:
        return "hero_stat", dict(stat=body[0].strip(),
                                 label=body[1] if len(body) > 1 else "",
                                 context=body[2] if len(body) > 2 else ""), \
            "standalone number line"

    low = joined.lower()
    if is_last and any(k in low for k in ("raising", "we ask", "the ask",
                                          "seeking", "contact", "join us",
                                          "thank")):
        return "closing", dict(title_lines=[title or "Thank you"],
                               subtitle=first_sentence(body)), "closing language"

    bullets = [re.sub(r"^[-*•]\s*", "", l) for l in body
               if re.match(r"^\s*[-*•]\s+", l)]
    if bullets:
        items = [split_head_body(b) for b in bullets]
        avg = sum(len(b.split()) for b in bullets) / len(bullets)
        if len(bullets) >= 5 and avg <= 9:
            return "icon_grid", dict(items=[dict(icon=icon_for(h + " " + b),
                                                 head=h, body=b)
                                            for h, b in items[:6]]), \
                f"{len(bullets)} short bullets"
        return "content", dict(bullets=[dict(head=h, body=b)
                                        for h, b in items[:4]]), \
            f"{len(bullets)} bullets"

    m = re.findall(rf"\b({NUM})\b", joined)
    if title and len(joined.split()) <= 14 and len(m) >= 1 and len(body) <= 2:
        return "hero_stat", dict(stat=m[0], label=title,
                                 context=joined), "short block with big number"

    sent = first_sentence(body) if body else ""
    bullets_txt = [dict(zip(("head", "body"), split_head_body(x)))
                   for x in re.split(r"(?<=[.!?])\s+", joined) if x.strip()][:4]
    return "content", dict(bullets=bullets_txt or
                           [dict(head=title or "Key point", body=joined[:140])]), \
        "prose fallback"


def draft_spec(text, theme="teal-orbit", brand="", goal="pitch"):
    blocks = split_blocks(text)
    slides, report = [], []
    for i, blk in enumerate(blocks):
        stype, fields, why = classify(blk, i == 0, i == len(blocks) - 1)
        title = None
        if blk and blk[0].startswith("# "):
            title = blk[0][2:].strip()
        sd = {"type": stype}
        sd.update(fields)
        if title and stype not in ("title", "closing"):
            sent = first_sentence([l for l in blk[1:] if l.strip()])
            if len(title.split()) <= 3 and len(sent.split()) >= 6 and stype in (
                    "content", "chart", "process", "icon_grid"):
                sd["kicker"] = title.upper()
                sd["title"] = " ".join(sent.split()[:10]).rstrip(".")
            else:
                sd.setdefault("title", title)
        sd["notes"] = " ".join(l.strip() for l in blk
                               if not l.startswith("# "))[:900]
        slides.append(sd)
        report.append((i + 1, stype, why))
    spec = {"theme": theme, "brand": brand, "timings": True, "slides": slides}
    return spec, report


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("notes", help="text/markdown file with raw notes")
    ap.add_argument("--theme", default="teal-orbit")
    ap.add_argument("--brand", default="")
    ap.add_argument("--goal", default="pitch")
    ap.add_argument("--out", help="spec JSON path (default: <notes>.spec.json)")
    ap.add_argument("--build", action="store_true", help="also build the .pptx")
    ap.add_argument("--pptx", help="output deck path (default: <notes>.pptx)")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.notes):
        print(f"no such file: {a.notes}")
        return 2
    with open(a.notes, encoding="utf-8") as f:
        text = f.read()
    spec, report = draft_spec(text, a.theme, a.brand, a.goal)
    out = a.out or os.path.splitext(a.notes)[0] + ".spec.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)
    print(f"drafted {len(spec['slides'])} slides -> {out}")
    for n, stype, why in report:
        print(f"  {n:2d}. {stype:12s} ({why})")
    if a.build:
        sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
        import generate_pptx
        deck = a.pptx or os.path.splitext(a.notes)[0] + ".pptx"
        generate_pptx.build(spec, deck)
        print(f"built {deck} — now run: python tools/qa_deck.py {deck}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
