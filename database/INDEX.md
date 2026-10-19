# Monster Database — Census

**238 records** across 12 tables in `powerpoint.db` (SQLite + FTS5 full-text). Rebuilt by `build_db.py` from `data/*.json` + generator themes.

| Table | Records | What |
|---|---:|---|
| `creators` | 36 | Creators/agencies like PlacePlate + what to steal |
| `templates` | 40 | PlacePlate catalog SKUs (PP-1…PP-49) |
| `bundles` | 7 | PlacePlate bundle packs |
| `services` | 2 | PlacePlate custom-service tiers |
| `techniques` | 25 | Morph + advanced technique library |
| `recipes` | 16 | Slide layout recipes R1–R16 (generator builders) |
| `structures` | 7 | Deck story arcs per use case |
| `principles` | 30 | Research principles + debunked myths |
| `tools` | 30 | Add-ins, AI tools, marketplaces, asset libraries |
| `fonts` | 12 | Latin + Arabic font system |
| `prompts` | 12 | Copy-paste AI prompt playbook |
| `palettes` | 21 | Generator themes + accessible chart palettes |

Query it: `python database/query.py --help` · Rebuild: `python database/build_db.py`
