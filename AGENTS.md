# Agent instructions for this repo

This repository is a **PowerPoint generation skill** built from research on
PlacePlate (@placeplate12345). For ANY request to create, design, restyle, or
review a PowerPoint / .pptx / slide deck — or questions about Morph transitions
or presentation design:

1. **Read `SKILL.md` first** and follow its operating procedure exactly.
2. Pull depth from `reference/*.md` (see the map inside `SKILL.md`) instead of
   relying on generic presentation habits.
3. Build decks with `generator/generate_pptx.py` (JSON spec in → Morph-enabled
   .pptx out). Extend that generator for new layouts/themes; keep `SKILL.md`,
   `reference/03-slide-recipes.md` and the code in sync when you do.
4. Always run the QA gate before presenting work as done:
   `python tools/qa_deck.py <file.pptx>` must report GREEN (it verifies Morph
   XML, `!!` anchors, word budgets, font floors, notes), plus the human boxes
   in `reference/06-qa-checklist-and-mistakes.md`.
4b. Query `database/powerpoint.db` first for creators/templates/techniques/
   research (`python database/query.py search "…"` / `recommend <use-case>`)
   before reaching for the web; rebuild with `database/build_db.py` after
   editing `database/data/*.json`.
5. Keep attribution intact: style & techniques are PlacePlate's (linked in
   `reference/08-source-research.md`); never reproduce their paid files or
   present output as an official PlacePlate template.
6. Python env: `pip install python-pptx`. A local `.venv/` may exist — it is
   gitignored; recreate freely.
