# pptconv — PowerPoint template conversion

Mission: convert Dawn's PowerPoint deck from its old template to the new corporate
template, at the highest quality bar. The result must genuinely look like it was
authored in the new template: its fonts, colors, layouts, and design rule.

## The iron rule

**No slide ships unseen.** Every converted slide gets rendered to PNG and visually
inspected (Read the image) side-by-side with the original before the deck is called
done. Renders can differ from PowerPoint if fonts are missing — run the font check
first and trust `output/qa/pair-NN.png` only after fonts resolve correctly.

## Workflow (documents land in input/)

1. `make inspect-src inspect-tpl` — understand both decks: layouts, placeholder
   inventories, theme typefaces, slide size/aspect. Read docs/REFERENCE.md first.
2. `make fonts` — extract embedded fonts, check theme typefaces resolve in
   fontconfig. Missing corporate fonts: get the .ttf/.otf into
   `~/.local/share/fonts/pptconv/` + `fc-cache -f` (ask JB if the family is unknown).
3. `make extract` — content manifest to work/content.json + work/assets/.
4. Author `mapping.json` (old layout name → new layout name; see tools/apply.py
   docstring). This is a judgment step: match by role, not by name similarity.
5. `make apply` — rebuild onto the new template. Read the printed REVIEW FLAGS and
   `output/converted.review.txt`.
6. `make render-src render-out qa` — then **Read every `output/qa/pair-NN.png`**
   and fix: overflow, wrong layout choice, lost content, off-brand colors/fonts.
   Iterate 5–6 until clean. `output/qa/contact.png` gives the whole-deck view.
7. Charts/SmartArt/groups are flagged, not silently converted — resolve each flag
   explicitly (recreate, rasterize from the source render, or ask).

## Stack

- Python venv `.venv` (via uv): python-pptx 1.0.2, Pillow, lxml. Run as
  `.venv/bin/python tools/<tool>.py`.
- LibreOffice 25.8 lives UNTRACKED in `vendor/libreoffice/` (user-space, no sudo);
  only `tools/render.sh` calls it. If vendor/ is missing (fresh clone), re-download:
  see docs/REFERENCE.md appendix or the render.sh header.
- `tools/`: inspect.py · extract.py · apply.py · qa.py · render.sh · install-fonts.sh
- docs/REFERENCE.md: the deep engineering notes (pptx anatomy, rebuild-vs-retrofit,
  python-pptx pitfalls). Read before touching apply.py.

## Conventions

- Rebuild strategy, never retrofit (REFERENCE.md §2) unless JB overrides.
- Old formatting dies; only structure + semantic emphasis (bold/italic/links) and
  speaker notes survive. That is the point of retemplating.
- input/ and output decks contain Dawn's work content — never publish or share
  outside this repo; repo stays private.
