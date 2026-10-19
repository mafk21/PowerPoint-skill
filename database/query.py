#!/usr/bin/env python3
"""
Query the Monster Database (database/powerpoint.db).

    python database/query.py search "morph 3d"          # full-text everything
    python database/query.py stats
    python database/query.py creators --platform tiktok
    python database/query.py creators --archetype template-shop
    python database/query.py templates --category academic --max-price 5
    python database/query.py templates --free
    python database/query.py techniques --category morph-advanced
    python database/query.py palettes --mode dark
    python database/query.py prompt 6                   # full prompt text
    python database/query.py recommend pitch            # structure+themes+SKUs
    python database/query.py recommend thesis
    python database/query.py sql "SELECT name,price_usd FROM templates WHERE price_usd=0"

Uses only the stdlib. Rebuild the DB with database/build_db.py.
"""
import argparse
import os
import sqlite3
import sys

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "powerpoint.db")


def rows(cur, q, args=()):
    cur.execute(q, args)
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def show(records, fields=None):
    if not records:
        print("(no matches)")
        return
    for r in records:
        keys = fields or list(r.keys())
        print(" • " + " | ".join(f"{r.get(k, '')}" for k in keys if r.get(k) not in ("", None)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Monster Database query CLI")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("search"); p.add_argument("term"); p.add_argument("--limit", type=int, default=15)
    sub.add_parser("stats")
    p = sub.add_parser("creators"); p.add_argument("--platform"); p.add_argument("--archetype")
    p = sub.add_parser("templates"); p.add_argument("--category"); p.add_argument("--max-price", type=float); p.add_argument("--free", action="store_true")
    p = sub.add_parser("techniques"); p.add_argument("--category"); p.add_argument("--generator")
    p = sub.add_parser("palettes"); p.add_argument("--mode", choices=["dark", "light"])
    p = sub.add_parser("prompt"); p.add_argument("id", nargs="?", type=int)
    p = sub.add_parser("recommend"); p.add_argument("use_case")
    p = sub.add_parser("principles"); p.add_argument("--myths", action="store_true")
    p = sub.add_parser("sql"); p.add_argument("query")
    a = ap.parse_args(argv)

    if not os.path.exists(DB):
        print("Database not built yet — run: python database/build_db.py")
        return 2
    db = sqlite3.connect(DB)
    cur = db.cursor()

    if a.cmd == "search":
        q = "SELECT kind, name, snippet(search_index,2,'[',']',' … ',18) AS hit " \
            "FROM search_index WHERE search_index MATCH ? ORDER BY rank LIMIT ?"
        fts = " ".join(a.term.split())
        try:
            recs = rows(cur, q, (fts, a.limit))
        except sqlite3.OperationalError:
            recs = rows(cur, q, (f'"{a.term}"', a.limit))
        for r in recs:
            print(f"[{r['kind']}] {r['name']}\n    {r['hit']}")
        if not recs:
            print("(no matches)")

    elif a.cmd == "stats":
        tables = rows(cur, "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'search%' AND name NOT LIKE 'sqlite%' ORDER BY name")
        total = 0
        for t in tables:
            n = rows(cur, f"SELECT COUNT(*) c FROM \"{t['name']}\"")[0]["c"]
            total += n
            print(f"{t['name']:14s} {n}")
        print(f"{'TOTAL':14s} {total}")

    elif a.cmd == "creators":
        q, args = "SELECT * FROM creators WHERE 1=1", []
        if a.platform: q += " AND platforms LIKE ?"; args.append(f"%{a.platform}%")
        if a.archetype: q += " AND archetype LIKE ?"; args.append(f"%{a.archetype}%")
        show(rows(cur, q, args), ["name", "handle", "platforms", "followers", "specialty"])

    elif a.cmd == "templates":
        q, args = "SELECT * FROM templates WHERE 1=1", []
        if a.category: q += " AND category=?"; args.append(a.category)
        if a.max_price is not None: q += " AND price_usd<=?"; args.append(a.max_price)
        if a.free: q += " AND price_usd=0"
        show(rows(cur, q + " ORDER BY price_usd, code", args),
             ["code", "name", "price_usd", "category", "layouts", "style_family", "url"])

    elif a.cmd == "techniques":
        q, args = "SELECT * FROM techniques WHERE 1=1", []
        if a.category: q += " AND category=?"; args.append(a.category)
        if a.generator: q += " AND generator_support LIKE ?"; args.append(f"%{a.generator}%")
        show(rows(cur, q, args), ["name", "category", "difficulty", "generator_support", "learned_from"])

    elif a.cmd == "palettes":
        q, args = "SELECT * FROM palettes WHERE kind='generator-theme'", []
        if a.mode: q += " AND mode=?"; args.append(a.mode)
        show(rows(cur, q, args), ["name", "mode", "bg", "accent", "ink"])

    elif a.cmd == "prompt":
        if a.id is None:
            show(rows(cur, "SELECT id,name,stage FROM prompts ORDER BY id"), ["id", "name", "stage"])
        else:
            recs = rows(cur, "SELECT * FROM prompts WHERE id=?", (a.id,))
            for r in recs:
                print(f"#{r['id']} {r['name']} ({r['stage']})\n\n{r['text']}")
            if not recs: print("(no such prompt)")

    elif a.cmd == "principles":
        q = "SELECT * FROM principles" + (" WHERE kind='MYTH'" if a.myths else " WHERE kind!='MYTH'")
        show(rows(cur, q), ["name", "kind", "generation_rule"])

    elif a.cmd == "recommend":
        keys = {"pitch": "%pitch%", "investor": "%pitch%", "thesis": "%thesis%",
                "academic": "%thesis%", "business": "%Business / company%",
                "sales": "%Sales%", "proposal": "%Sales%", "seminar": "%Seminar%",
                "training": "%Seminar%", "qbr": "%review%", "decision": "%decision%",
                "exec": "%decision%"}
        like = keys.get(a.use_case.lower(), f"%{a.use_case}%")
        recs = rows(cur, "SELECT * FROM structures WHERE name LIKE ?", (like,))
        if not recs:
            print("Use cases: pitch, thesis, business, sales, seminar, qbr, decision")
            return 1
        for r in recs:
            print(f"== {r['name']}  ({r['slide_count']} slides, {r['minutes']} min)")
            print(f"Framework : {r['primary_framework']}")
            print(f"Themes    : {r['suggested_themes']}")
            print(f"Buy orig. : {r['placeplate_skus']}  (reference/07)")
            print(f"Sequence  : {r['sequence']}")
        pals = rows(cur, "SELECT name,bg,accent FROM palettes WHERE kind='generator-theme' AND name IN (%s)"
                    % ",".join("?" * 3), tuple(t.strip() for t in recs[0]["suggested_themes"].split(",")[:3]))
        for p in pals:
            print(f"  palette {p['name']}: bg #{p['bg']} accent #{p['accent']}")

    elif a.cmd == "sql":
        if not a.query.strip().lower().startswith("select"):
            print("Read-only: SELECT only.")
            return 2
        for r in rows(cur, a.query):
            print(r)
    db.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
