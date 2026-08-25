# pptconv

Convert a PowerPoint deck from an old template to a new one, with a verification
loop that renders every slide and compares old vs new visually.

Pipeline: `inspect → fonts → extract → mapping.json → apply → render → qa`.
See `CLAUDE.md` for the working rules and `docs/REFERENCE.md` for the engineering
notes. Drop the source deck and the new template into `input/`, then use the
Makefile targets.

Custom toolkit (python-pptx + Pillow + user-space LibreOffice render stack):

| tool | job |
|---|---|
| `tools/inspect_deck.py` | dump layouts, placeholders, theme fonts/colors of any deck |
| `tools/extract.py` | deck → neutral content manifest (`work/content.json` + assets) |
| `tools/apply.py` | rebuild manifest onto the new template via `mapping.json`, with review flags |
| `tools/render.sh` | pptx → PDF → per-slide PNGs (LibreOffice headless + pdftoppm) |
| `tools/qa.py` | side-by-side old/new pair images + contact sheet for visual review |
| `tools/install-fonts.sh` | extract embedded fonts, verify theme typefaces resolve |
