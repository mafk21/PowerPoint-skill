# 🐲 The Monster Database

A real, queryable SQLite database of everything this repo knows — built from
the JSON sources in `data/` so it can grow forever without hand-editing SQL.

- **238 records · 12 tables · FTS5 full-text search** (census: [`INDEX.md`](INDEX.md))
- Rebuild anytime: `python database/build_db.py` (stdlib only)
- Query CLI: `python database/query.py …` (stdlib only)

## Schema

| Table | Source JSON | Contents |
|---|---|---|
| `creators` (36) | `data/catalog.json` | PlacePlate-like creators/agencies: handle, platforms, followers snapshot, archetype, specialty, **what to steal** |
| `templates` (40) | 〃 | PlacePlate SKUs PP-1…PP-49: price, category, layouts, style family, product URL |
| `bundles` (7) · `services` (2) | 〃 | Bundle packs + custom-service tiers |
| `techniques` (25) | `data/knowledge.json` | Morph core/recipes/advanced, interactive, craft — steps, difficulty, learned-from, generator support |
| `recipes` (16) | 〃 | Slide recipes R1–R16 ↔ generator builder types, purposes, content limits |
| `structures` (7) | 〃 | Deck arcs: slide sequence (as recipe IDs), framework, themes, matching PlacePlate SKUs |
| `principles` (30) | 〃 | Research principles (Richard Mayer/CLT/memory/dataviz/accessibility) + `kind='MYTH'` debunks |
| `tools` (30) | `data/arsenal.json` | Add-ins, AI tools (with honest verdicts), marketplaces, asset libraries |
| `fonts` (12) | 〃 | Latin + Arabic type system and pairings |
| `prompts` (12) | 〃 | The prompt playbook, full text, by stage |
| `palettes` (10) | **live import** | 8 generator themes (imported from `generate_pptx.THEMES` — can never drift) + Okabe-Ito + strategic-gray scheme |
| `search_index` (FTS5) | generated | Full-text over every field of every table |

## Query examples

```bash
python database/query.py stats                       # census
python database/query.py search "logo reveal"       # anything, everywhere
python database/query.py search "colorblind"         # → Okabe-Ito + WCAG rules
python database/query.py creators --platform tiktok --archetype template-shop
python database/query.py templates --category academic --max-price 5
python database/query.py templates --free
python database/query.py techniques --category morph-advanced
python database/query.py palettes --mode dark
python database/query.py principles --myths          # the debunk list
python database/query.py prompt 6                    # full JSON-spec prompt
python database/query.py recommend thesis            # structure+themes+SKUs+palettes joined
python database/query.py sql "SELECT name, followers FROM creators WHERE archetype LIKE '%shop%'"
```

`recommend` accepts: `pitch · thesis · business · sales · seminar · qbr · decision`.

## Extending the monster

1. Edit the relevant `data/*.json` (keep records flat: strings/numbers; lists
   are JSON-encoded into a TEXT column automatically).
2. New table? Add its records list to a JSON, register `(table, records)` in
   `build_db.py`'s `datasets` + a line in `what` — FTS picks it up for free.
3. `python database/build_db.py` → commit JSONs **and** the rebuilt
   `powerpoint.db` + `INDEX.md`.

Freshness discipline: follower counts/prices are snapshots (Oct 2026) — the
`followers`/`price_usd` fields are for triage, not citation. Verify live
before quoting externally; update the JSON + rebuild when you do.
