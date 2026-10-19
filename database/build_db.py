#!/usr/bin/env python3
"""
Build the Monster Database: database/powerpoint.db (SQLite + FTS5).

Single source of truth = JSON files in database/data/ + generator themes
(imported live from generator/generate_pptx.py so palettes can never drift).
Re-run any time after editing the JSONs:

    python database/build_db.py

Produces:
  * powerpoint.db  — tables: creators, templates, bundles, services,
    techniques, recipes, structures, principles, tools, fonts, prompts,
    palettes + FTS5 `search_index` over everything.
  * INDEX.md       — human-readable census of the database.
"""
import io
import json
import os
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, "generator"))

DB_PATH = os.path.join(HERE, "powerpoint.db")


def load(name):
    with io.open(os.path.join(HERE, "data", name), encoding="utf-8") as f:
        return json.load(f)


def table_from_records(cur, table, records):
    """Create a table whose columns are the union of record keys (all TEXT
    except numeric-looking prices kept as REAL when every value is numeric)."""
    keys = []
    for r in records:
        for k in r:
            if k not in keys:
                keys.append(k)
    cols = []
    for k in keys:
        vals = [r.get(k) for r in records if r.get(k) is not None]
        numeric = vals and all(isinstance(v, (int, float)) for v in vals)
        cols.append(f'"{k}" {"REAL" if numeric else "TEXT"}')
    cur.execute(f'DROP TABLE IF EXISTS "{table}"')
    cur.execute(f'CREATE TABLE "{table}" ({", ".join(cols)})')
    ph = ", ".join("?" * len(keys))
    for r in records:
        cur.execute(f'INSERT INTO "{table}" VALUES ({ph})',
                    [json.dumps(r.get(k), ensure_ascii=False)
                     if isinstance(r.get(k), (list, dict)) else r.get(k)
                     for k in keys])
    return len(records)


def main():
    catalog = load("catalog.json")
    knowledge = load("knowledge.json")
    arsenal = load("arsenal.json")

    # Palettes: generator themes (live import = single source of truth) + extras
    import generate_pptx as gen
    palettes = []
    for name, t in gen.THEMES.items():
        palettes.append({
            "name": name, "kind": "generator-theme",
            "mode": "dark" if t["dark"] else "light",
            "bg": t["bg"], "card": t["card"], "ink": t["ink"],
            "muted": t["muted"], "accent": t["accent"],
            "accent2": t["accent2"], "graybar": t["graybar"],
            "use": "Deck theme in generator/generate_pptx.py (see reference/02 for best-for)",
            "source": "generator THEMES",
        })
    for p in arsenal["palettes_extra"]:
        palettes.append({"name": p["name"], "kind": "chart-palette", "mode": "any",
                         "bg": "", "card": "", "ink": "", "muted": "",
                         "accent": " ".join(p["colors"]), "accent2": "",
                         "graybar": "", "use": p["use"], "source": p["source"]})

    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    db = sqlite3.connect(DB_PATH)
    cur = db.cursor()
    counts = {}
    datasets = [
        ("creators", catalog["creators"]),
        ("templates", catalog["templates"]),
        ("bundles", catalog["bundles"]),
        ("services", catalog["services"]),
        ("techniques", knowledge["techniques"]),
        ("recipes", knowledge["recipes"]),
        ("structures", knowledge["structures"]),
        ("principles", knowledge["principles"]),
        ("tools", arsenal["tools"]),
        ("fonts", arsenal["fonts"]),
        ("prompts", arsenal["prompts"]),
        ("palettes", palettes),
    ]
    for table, records in datasets:
        counts[table] = table_from_records(cur, table, records)

    # Full-text search over everything (kind, name, body)
    cur.execute("CREATE VIRTUAL TABLE search_index USING fts5(kind, name, body)")
    name_keys = {"creators": "name", "templates": "name", "bundles": "name",
                 "services": "name", "techniques": "name", "recipes": "name",
                 "structures": "name", "principles": "name", "tools": "name",
                 "fonts": "name", "prompts": "name", "palettes": "name"}
    for table, records in datasets:
        for r in records:
            body = " | ".join(f"{k}: {v}" for k, v in r.items())
            cur.execute("INSERT INTO search_index VALUES (?,?,?)",
                        (table, str(r.get(name_keys[table], "")), body))
    total = sum(counts.values())
    db.commit()

    # Human-readable census
    lines = ["# Monster Database — Census", "",
             f"**{total} records** across {len(datasets)} tables in `powerpoint.db` "
             "(SQLite + FTS5 full-text). Rebuilt by `build_db.py` from "
             "`data/*.json` + generator themes.", "",
             "| Table | Records | What |", "|---|---:|---|"]
    what = {
        "creators": "Creators/agencies like PlacePlate + what to steal",
        "templates": "PlacePlate catalog SKUs (PP-1…PP-49)",
        "bundles": "PlacePlate bundle packs",
        "services": "PlacePlate custom-service tiers",
        "techniques": "Morph + advanced technique library",
        "recipes": "Slide layout recipes R1–R16 (generator builders)",
        "structures": "Deck story arcs per use case",
        "principles": "Research principles + debunked myths",
        "tools": "Add-ins, AI tools, marketplaces, asset libraries",
        "fonts": "Latin + Arabic font system",
        "prompts": "Copy-paste AI prompt playbook",
        "palettes": "Generator themes + accessible chart palettes",
    }
    for table, _ in datasets:
        lines.append(f"| `{table}` | {counts[table]} | {what[table]} |")
    lines += ["", "Query it: `python database/query.py --help` · "
              "Rebuild: `python database/build_db.py`"]
    with io.open(os.path.join(HERE, "INDEX.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    db.close()
    print(f"Built {DB_PATH}: {total} records in {len(datasets)} tables + FTS index")
    for t, c in counts.items():
        print(f"  {t:12s} {c}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
