# pptconv

Convert a PowerPoint deck from an old template to a new one, with a verification
loop that renders every slide and compares old vs new visually.

Built with [Claude Code](https://claude.com/claude-code): the first fully
converted deck — a real corporate roadmap moved onto the company's new
presentation template — was delivered eight minutes after the source files
arrived. The decks
themselves are not in this repo (they are work content and stay untracked in
`input/`/`output/`); what is here is the reusable workbench and the
project-specific conversion script.

## How it works

Pipeline: `inspect → fonts → extract → mapping → apply/graft → render → visual QA`.

| tool | job |
|---|---|
| `tools/inspect_deck.py` | dump layouts, placeholders, theme fonts/colors of any deck |
| `tools/extract.py` | deck → neutral content manifest (`work/content.json` + assets) |
| `tools/apply.py` | generic rebuild of a manifest onto a new template via a layout mapping, with review flags |
| `tools/graft_roadmap.py` | the project-specific conversion: XML-grafts diagram slides into the new template |
| `tools/render.sh` | pptx → PDF → per-slide PNGs (headless LibreOffice + pdftoppm, both user-space) |
| `tools/qa.py` | side-by-side old/new pair images + contact sheet for visual review |
| `tools/install-fonts.sh` | extract embedded fonts, verify theme typefaces resolve in fontconfig |

The deliverable is always native OOXML written by python-pptx on top of the new
template's own masters and theme, so PowerPoint opens it as a first-class file.
LibreOffice is only used to render preview images.

The key idea for diagram-heavy slides (`graft_roadmap.py`): the source diagrams
were drawn in *theme* colors and fonts, so transplanting their shape XML into the
new deck makes them restyle themselves — every scheme color and theme font
reference re-resolves against the new template. Three subtleties made it work:

1. Transplanted placeholder shapes must be de-placeholdered (the new layout has no
   matching placeholders), which silently drops their inherited list styles — so
   the source layout's `lstStyle` is baked into each shape first.
2. The new theme font (Nunito) is wider than the old one, so short fixed-box
   labels ("2027") started wrapping inside the timeline circles; short single-token
   labels get `wrap="none"`.
3. Empty placeholders (fill-in-later text boxes) become invisible once
   de-placeholdered — each is refilled with the scaffold text of the filled box in
   the same row, so every column stays editable.

Every one of those was caught by the visual QA loop, not by reading code: render
both decks, build side-by-side pair images, and actually look at them (Claude
reads the images natively) before anything ships. House rule: **no slide ships
unseen.**

## Why chat AI tools couldn't do this — and what did

This conversion was first attempted with Claude Cowork and other AI assistants,
and the results weren't usable. That's not about model quality; the same family
of models did the work here. It's about what the AI is allowed to do around the
model. A chat/workspace assistant receives a file, transforms it in one pass,
and hands back the result without ever seeing what it produced. It can read and
write documents, but it can't install software, run a renderer, or look at its
own output — and PowerPoint punishes exactly that: a deck like this one is
dozens of vector shapes with inherited styles, and blind edits quietly break
them (or the "conversion" comes back as a lossy AI re-creation instead of a
real, editable deck).

The setup used here is different: **Claude Code running in a terminal on a
plain Linux server**, where the AI has a shell, a filesystem, and a development
loop. Concretely, that meant it could:

1. **Install its own toolchain.** Python + `python-pptx` for surgical file
   manipulation, and headless LibreOffice + poppler for rendering — installed
   user-space, no admin rights needed (see `tools/render.sh`; the LibreOffice
   deb tree is simply extracted into `vendor/`).
2. **Write custom programs against the file instead of editing it blind.**
   The diagrams were transplanted at XML level into the new template
   (`tools/graft_roadmap.py`), so every theme color and font reference
   re-resolved to the new design by itself. Native OOXML in, native OOXML out —
   the deliverable opens in PowerPoint as a first-class, fully editable file.
3. **Check its own work visually.** Every slide is rendered to an image and
   compared side by side with the original (`tools/qa.py`) before anything is
   sent — the AI reads the images with its own vision. That loop caught three
   bugs a one-shot tool would have shipped: inherited styles lost in transfer,
   year labels wrapping because the new font runs wider, and empty text
   placeholders turning invisible.
4. **Operate as a colleague, not a vending machine.** The work lives in a real
   project under version control, so every feedback round is a small fix to an
   existing, tested pipeline rather than a fresh roll of the dice — with the
   before/after renders as evidence for each change.

None of this requires a special server, by the way. Claude Code runs in a
terminal on an ordinary laptop (macOS/Linux/Windows) and can set up everything
it needs in user space, exactly as it did here — so the workflow in this repo is
reproducible wherever a terminal is available.

## The prompts that built this

This project was driven end-to-end by natural-language prompts to Claude Code
(model: Claude Fable 5, with subagents fanned out for parallel setup work).
Paraphrased, in order:

1. *"Set up a new project, `pptconv`, for converting a PowerPoint presentation
   from an old corporate template to a new one. Install the tooling the job
   needs, and parallelize the setup work with subagents."*
   → repo scaffold, Python environment (uv + python-pptx/Pillow/lxml),
   user-space LibreOffice install (no root required), the generic pipeline
   tools, and an experimentally verified engineering reference
   (`docs/REFERENCE.md`) written by a parallel subagent.
2. *"The documents arrive later — build and prove the full pipeline on test data
   before they do."* → end-to-end dress rehearsal on a generated sample deck,
   which already caught one real bug (template example slides leaking into the
   output).
3. *"Don't stop at off-the-shelf tools: write custom scripts where the generic
   approach falls short, and verify the result visually — render the slides and
   check that fonts, colors, and design carry over correctly."* → the render +
   side-by-side QA loop became the backbone of the workflow, and the generic
   rebuild tool was complemented by the project-specific XML graft.
4. Feedback-driven iteration once the real deck was converted: a font-metrics
   bug (the new template's font runs wider, wrapping short labels — fixed with
   a no-wrap rule) and a reviewer-reported issue (*"the text boxes next to each
   year are missing"* — empty placeholders had turned invisible, fixed with the
   scaffold refill). Each fix was verified visually against the original before
   delivery.

## Reusing it

For a new conversion: drop the source deck and target template in `input/`, then
`make inspect-src inspect-tpl fonts extract` and either use the generic
`tools/apply.py` with a `mapping.json`, or write a small project-specific graft
like `graft_roadmap.py` when the slides are diagram-heavy. `make render-src
render-out qa` gives you the pair images to judge by. See `CLAUDE.md` for the
working rules and `docs/REFERENCE.md` for the pptx internals (placeholder
mechanics, rebuild vs retrofit, chart/SmartArt grafting).

## License

MIT — see `LICENSE`.
