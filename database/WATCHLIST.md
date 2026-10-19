# Data watchlist — keeping the Monster Database honest

The database is a snapshot of a moving ecosystem. Numbers below were
verified **2026-10-19**. Anything older than a quarter should be treated
as "last known", not current. Refresh protocol at the bottom.

## Fast-moving (re-check quarterly)

| What | Where it lives | Why it rots |
|---|---|---|
| Creator follower counts (36 creators) | `data/catalog.json` → `creators`; atlas in `reference/09` | Social counts move monthly; two accounts were already ±20% between checks |
| PlacePlate catalog & prices (47 products) | `data/catalog.json` → `templates`/`bundles`; `reference/07` | New SKUs land as PP-5x; bundle pricing shifts; Gumroad may 403 scrapers — verify via a real browser |
| AI presentation-tool landscape (Copilot, Gamma, Beautiful.ai, Canva…) | `data/arsenal.json` → `tools`; `reference/13` | PPTX-export fidelity rankings changed twice in 2026 alone; re-run the export test, don't trust listicles |
| Add-in pricing (think-cell, PPT Productivity, Power-user…) | `data/arsenal.json`; `reference/13` | Annual price hikes; BrightSlide free-tier scope |

## Slow-moving (re-check yearly)

| What | Where | Notes |
|---|---|---|
| Research effect sizes (Mayer, Alley…) | `reference/10` | Stable literature; watch for replication/meta-analysis updates |
| Platform specs (IG/TikTok formats, safe zones) | `reference/13 §6`; `generator/portrait.py` FORMATS | Carousel max (20) and 4:5 preference changed before — check before big launches |
| Font availability (Google Fonts Arabic set) | `reference/14` | New Arabic display faces appear ~yearly |
| PowerPoint Morph/Designer changes | `reference/01`, `11` | Microsoft ships silently; re-test `!!` + byWord behavior after major Office updates |

## Never store here

Prices paid, crypto/Binance payment details, license keys, or anything
from a private transaction. Public list prices only.

## Refresh protocol (15 minutes)

1. Re-verify the fast-moving rows (browser, logged sessions where needed).
2. Edit `database/data/*.json` (the only source of truth) + matching
   `reference/` pages.
3. `python database/build_db.py` — rebuilds `powerpoint.db` + `INDEX.md`.
4. `python tools/db_lint.py` — must print CLEAN (catches dupes, bad hex,
   contrast drift, stale DB, undocumented builders).
5. `python -m unittest discover -s tests` — full suite green.
6. Update the "verified" date above; commit data + db + docs together.
