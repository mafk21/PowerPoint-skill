#!/usr/bin/env python3
"""DataFeed — refresh a spec's charts from CSV (the monthly-deck autopilot).

Mark data slides in your spec once:

  {"type": "chart", "title": "...findings...", "csv": "data/q3.csv",
   "csv_label": "Month", "csv_values": ["Revenue"]}

then every month: drop the new CSV in place and run
`python tools/datafeed.py qbr.spec.json --build deck.pptx`.
Works for chart / donut / funnel / bars (single value column) and
line_chart (csv_values = one column per series). Highlight moves to the
new maximum unless --keep-highlight. One-off mode can target a slide
directly: `--slide 5 --csv file.csv --label Month --values Revenue,Costs`.

Numbers may carry $, %, thousands separators, and (parenthesized)
negatives; delimiters , ; and tab are sniffed.
"""
import argparse
import csv as csvmod
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
CHARTY = ("chart", "donut", "funnel", "bars", "line_chart")


def num(x):
    s = str(x).strip()
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[$,%\s]", "", s.strip("()"))
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csvmod.Sniffer().sniff(sample, delimiters=",;\t")
        except Exception:
            dialect = csvmod.excel
        return list(csvmod.DictReader(f, dialect=dialect))


def feed_slide(sd, rows, label_col, value_cols, keep_highlight=False):
    labels = [str(r.get(label_col, "")).strip() for r in rows]
    matrix = []
    for col in value_cols:
        matrix.append([num(r.get(col)) for r in rows])
    keep = [i for i in range(len(rows))
            if labels[i] and all(m[i] is not None for m in matrix)]
    labels = [labels[i] for i in keep]
    matrix = [[m[i] for i in keep] for m in matrix]
    t = sd.get("type")
    if t == "line_chart":
        sd["categories"] = labels
        sd["series"] = [{"name": c, "values": matrix[k]}
                        for k, c in enumerate(value_cols)]
    elif t == "funnel":
        sd["stages"] = [{"label": l, "value": v}
                        for l, v in zip(labels, matrix[0])]
    else:
        sd["categories"] = labels
        sd["values"] = matrix[0]
        if not keep_highlight and matrix[0]:
            sd["highlight"] = matrix[0].index(max(matrix[0]))
    return len(labels)


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("spec")
    ap.add_argument("--slide", type=int, help="one-off: slide number (1-based)")
    ap.add_argument("--csv", help="one-off: CSV path")
    ap.add_argument("--label", help="one-off: label column")
    ap.add_argument("--values", help="one-off: value column(s), comma list")
    ap.add_argument("--keep-highlight", action="store_true")
    ap.add_argument("--out", help="output spec (default: overwrite input)")
    ap.add_argument("--build", metavar="DECK.pptx", help="also build the deck")
    a = ap.parse_args(argv[1:])
    if not os.path.exists(a.spec):
        print(f"no such file: {a.spec}")
        return 2
    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)
    base = os.path.dirname(os.path.abspath(a.spec))
    jobs = []
    if a.slide:
        if not (a.csv and a.label and a.values):
            print("one-off mode needs --csv, --label and --values")
            return 2
        jobs = [(a.slide, a.csv, a.label,
                 [v.strip() for v in a.values.split(",")])]
    else:
        for i, sd in enumerate(spec.get("slides", []), 1):
            if sd.get("csv"):
                jobs.append((i, sd["csv"], sd.get("csv_label", ""),
                             sd.get("csv_values") or []))
    if not jobs:
        print('no slides carry "csv" — tag chart slides with '
              '"csv"/"csv_label"/"csv_values" or use --slide mode')
        return 1

    for n, path, label_col, value_cols in jobs:
        sd = spec["slides"][n - 1]
        if sd.get("type") not in CHARTY:
            print(f"slide {n}: type {sd.get('type')} not chart-like, skipped")
            continue
        cpath = path if os.path.isabs(path) else os.path.join(base, path)
        if not os.path.exists(cpath):
            print(f"slide {n}: CSV missing: {cpath}")
            return 1
        rows = read_csv(cpath)
        if not rows:
            print(f"slide {n}: empty CSV {cpath}")
            return 1
        if not label_col:
            label_col = next(iter(rows[0].keys()))
        if not value_cols:
            value_cols = [k for k in rows[0] if k != label_col][:1]
        got = feed_slide(sd, rows, label_col, value_cols, a.keep_highlight)
        print(f"slide {n} ({sd.get('type')}): {got} rows "
              f"<- {os.path.basename(cpath)} [{label_col} / {', '.join(value_cols)}]")

    out = a.out or a.spec
    with open(out, "w", encoding="utf-8") as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)
    print(f"spec updated -> {out}")
    if a.build:
        sys.path.insert(0, os.path.join(ROOT, "..", "generator"))
        if spec.get("format") in ("reel", "carousel"):
            import portrait
            portrait.build_portrait(spec, a.build)
        else:
            import generate_pptx
            generate_pptx.build(spec, a.build)
        print(f"built {a.build} — QA it: python tools/qa_deck.py {a.build}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
