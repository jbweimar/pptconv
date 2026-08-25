# Template Conversion Reference (python-pptx 1.0.2)

Engineering reference for converting decks from an OLD corporate template to a NEW one.
Environment: `/home/jbweimar/projects/pptconv/.venv/bin/python`, python-pptx **1.0.2** (API claims below verified against the installed package). Repo dirs: `input/` (source decks + templates), `work/` (manifests, mappings), `output/` (converted decks + renders), `tools/` (scripts).

## 1. PPTX anatomy relevant to conversion

A `.pptx` is a zip (OPC package) of XML parts wired together by relationship files (`_rels/*.rels`):

- `ppt/presentation.xml` — slide id list (order), slide size (`p:sldSz`), refs to masters and notesMaster.
- `ppt/slideMasters/slideMaster1.xml` — top of the inheritance chain: background, default text styles per placeholder type and outline level (`p:txStyles`), color map (`p:clrMap`), list of its layouts.
- `ppt/slideLayouts/slideLayoutN.xml` — one per layout, child of exactly one master. Defines placeholder geometry/formatting that slides inherit.
- `ppt/slides/slideN.xml` — actual content; related to exactly one layout.
- `ppt/theme/theme1.xml` — one per master. `a:clrScheme` (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink), `a:fontScheme` (`a:majorFont`/`a:minorFont` = headings/body, referenced in XML as `+mj-lt`/`+mn-lt`), `a:fmtScheme` (fill/line/effect defaults).
- `ppt/notesMasters/notesMaster1.xml` + `ppt/notesSlides/notesSlideN.xml` — speaker notes.
- Binary parts: `ppt/media/*` (images/video), `ppt/embeddings/*` (chart xlsx, OLE), `ppt/charts/chartN.xml`, `ppt/diagrams/*` (SmartArt).

**Inheritance**: slide → layout → master → theme. Any property set at a lower level overrides the level above; anything *unset* (element absent) inherits. This is why re-templating works at all — and why old decks full of direct formatting don't pick up a new theme.

**Placeholder mechanics**: a placeholder is a shape whose `p:spPr`-sibling `p:nvSpPr/p:nvPr` contains `<p:ph type="..." idx="..."/>`. Matching between slide and layout is by **`idx`** (plus `type` for idx-0 title variants), *not* by name or position. A slide placeholder with `idx="1"` inherits geometry and formatting from the layout placeholder with `idx="1"`. In python-pptx:

```python
ph.placeholder_format.idx    # int, e.g. 0 (title), 1, 2, 10...
ph.placeholder_format.type   # PP_PLACEHOLDER member: TITLE, BODY, OBJECT, PICTURE, DATE...
slide.placeholders[1]        # keyed by idx, NOT position; KeyError:
                             #   "no placeholder on this slide with idx == 1" if absent
```

Verified: `Slides.add_slide(layout)` calls `shapes.clone_layout_placeholders(layout)`, which clones every layout placeholder **except** `DATE`, `FOOTER`, `SLIDE_NUMBER` (those stay "latent" and render from the layout when enabled). Cloned placeholders carry `type/orient/sz/idx` but **no position or size** — `ph.left` etc. read through to the layout, so an unedited placeholder always sits exactly where the new template puts it. Don't set `left/top/width/height` on placeholders unless you deliberately want to detach them from the layout geometry.

If you ever reassign a slide's layout in place (retrofit), PowerPoint matches placeholders by idx; slide placeholders whose idx doesn't exist in the new layout become **orphans** — they keep their last inherited geometry frozen as direct values and stop inheriting.

## 2. Two strategies

**(a) REBUILD (default — use this).** `Presentation("new_template.pptx")` as base, delete its stub slides, create fresh slides from the *new* layouts, copy content over placeholder-by-placeholder. Output contains only the new master/layouts/theme; zero old formatting survives unless you deliberately copy it. Clean, predictable, reviewable.

**(b) RETROFIT.** Open the old deck, graft the new master+layouts+theme parts in, repoint each slide's layout relationship, delete old masters. Fast and preserves exotic content (SmartArt, animations, embedded objects) untouched — but inherits all junk: run-level font/color overrides on nearly every shape (so the new theme barely shows), orphaned placeholders, stale color maps, old layouts lingering in the package. Acceptable only when (1) decks are huge and content is overwhelmingly non-placeholder (charts/SmartArt-heavy), and (2) the old decks are known to be clean of direct formatting. Otherwise rebuild. python-pptx has no API for retrofit; it is raw part surgery — budget for it accordingly.

The rest of this document assumes **rebuild**.

## 3. Rebuild mechanics

```python
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

old = Presentation("input/old_deck.pptx")
new = Presentation("input/new_template.pptx")   # the new template IS the base

# Remove any stub slides shipped inside the template (keep layouts/master):
for sldId in list(new.slides._sldIdLst):        # no public delete API in 1.0.2
    rId = sldId.rId
    new.part.drop_rel(rId)
    new.slides._sldIdLst.remove(sldId)

layout = new.slide_layouts.get_by_name("Title and Content")   # verified API; returns
assert layout is not None                                     # default (None) if missing
slide = new.slides.add_slide(layout)
```

`prs.slide_layouts` is the *first* master's layouts. Multi-master templates: iterate `prs.slide_masters` and each `master.slide_layouts`. Match layouts **by name**, never index — layout order differs between templates.

**Placeholder copy — text.** The point of re-templating is to DROP old overrides and inherit the new look. So: copy plain text + paragraph structure (paragraph breaks, indent `level`), preserve only *semantic* emphasis (bold/italic/underline, hyperlinks), and drop fonts, sizes, and colors so theme styles rule. Unset font attributes read as `None` in python-pptx (verified) — that is exactly the "inherit" state you want to reproduce:

```python
def copy_text_frame(src_tf, dst_tf, keep_emphasis=True):
    dst_tf.clear()                       # leaves one empty paragraph
    for i, sp in enumerate(src_tf.paragraphs):
        dp = dst_tf.paragraphs[0] if i == 0 else dst_tf.add_paragraph()
        dp.level = sp.level              # outline level drives new template's bullet styles
        for sr in sp.runs:
            dr = dp.add_run()
            dr.text = sr.text
            if keep_emphasis:
                # copy ONLY explicitly-set semantic emphasis; leave None as None
                for attr in ("bold", "italic", "underline"):
                    v = getattr(sr.font, attr)
                    if v is not None:
                        setattr(dr.font, attr, v)
            if sr.hyperlink.address:
                dr.hyperlink.address = sr.hyperlink.address
            # deliberately NOT copied: font.name, font.size, font.color
```

Escape hatch: if a run's color is a scheme color carrying meaning (e.g. `ACCENT_1` used as "our highlight"), `run.font.color.type == MSO_THEME_COLOR_TYPE`-style checks let you re-apply `dr.font.color.theme_color = sr.font.color.theme_color` — the *new* theme then resolves it. Never copy `.rgb` values.

**Placeholder copy — driver loop** for one slide, given a resolved target layout and idx map (section 6):

```python
new_slide = new.slides.add_slide(target_layout)
for src_ph in old_slide.placeholders:
    dst_idx = idx_map.get(src_ph.placeholder_format.idx)   # from mapping file
    if dst_idx is None:
        report_orphan(src_ph); continue                    # content w/o a home -> §4 or review
    dst_ph = new_slide.placeholders[dst_idx]
    if src_ph.has_text_frame and src_ph.text_frame.text.strip():
        copy_text_frame(src_ph.text_frame, dst_ph.text_frame)
```

Empty cloned placeholders left over ("Click to add text" ghosts) are harmless — PowerPoint doesn't print/render prompt text — but you can delete them: `sp = ph._element; sp.getparent().remove(sp)`.

**Title convenience**: `slide.shapes.title` returns the idx-0 TITLE/CENTER_TITLE placeholder or `None`.

## 4. Non-placeholder content

Iterate `old_slide.shapes`, skip `shape.is_placeholder`, dispatch on `shape.shape_type` (`MSO_SHAPE_TYPE`).

**Pictures** (`PICTURE`, 13) — round-trip via blob (verified):

```python
import io
blob = src_pic.image.blob            # original bytes; .ext, .size (px), .dpi also available
new_pic = new_slide.shapes.add_picture(
    io.BytesIO(blob), left, top, width=w)   # give width OR height -> native aspect kept
```

`add_picture` with only one dimension preserves aspect ratio. Repositioning: scale the old (left, top, width, height) by `new.slide_width / old.slide_width` (and height ratio) as a first pass — see §7 for aspect-change math. If the new layout has a `PICTURE` placeholder, prefer `placeholder.insert_picture(io.BytesIO(blob))` (verified: returns `PlaceholderPicture`, auto-crops to fill the placeholder; inspect/adjust `.crop_left` etc. if the crop is bad).

**Tables** — recreate cell-by-cell so the new template's table style applies:

```python
src_tbl = shape.table                 # shape.has_table
gf = new_slide.shapes.add_table(len(src_tbl.rows), len(src_tbl.columns), l, t, w, h)
dst_tbl = gf.table
for r in range(len(src_tbl.rows)):
    for c in range(len(src_tbl.columns)):
        copy_text_frame(src_tbl.cell(r, c).text_frame, dst_tbl.cell(r, c).text_frame)
# merged cells: src cell.is_merge_origin / .span_height/.span_width ->
#   dst_tbl.cell(r,c).merge(dst_tbl.cell(r2,c2)); skip cells with .is_spanned
dst_tbl.first_row = src_tbl.first_row  # header-row banding flags
dst_tbl.first_col = src_tbl.first_col
```

Column widths: copy proportionally (`dst_tbl.columns[i].width = int(src_w * scale)`).

**Charts** — `shape.has_chart`. Two paths, both verified working in 1.0.2:

*Recreate from data* (preferred — chart picks up new theme colors) for the category-chart families python-pptx can read AND write (bar/column, line, pie, doughnut, area, radar):

```python
from pptx.chart.data import CategoryChartData
ch = shape.chart
cd = CategoryChartData()
cd.categories = list(ch.plots[0].categories)
for ser in ch.plots[0].series:
    cd.add_series(ser.name, tuple(ser.values))
new_slide.shapes.add_chart(ch.chart_type, l, t, w, h, cd)
```

Limits: XY/bubble need `XyChartData`/`BubbleChartData` and per-point reads; stock/surface/3-D and heavily customized charts (data labels, secondary axes, custom colors) lose their tweaks. For those, *graft the chart part* (verified end-to-end, including the embedded xlsx):

```python
import copy
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart

cp = shape.chart_part
pkg = new.part.package
new_cp = ChartPart(pkg.next_partname("/ppt/charts/chart%d.xml"), cp.content_type,
                   pkg, element=copy.deepcopy(cp._element))
for rel in cp.rels.values():                       # embedded workbook etc.
    if not rel.is_external:
        tp = rel.target_part
        new_cp.relate_to(Part(PackURI(str(tp.partname)), tp.content_type, pkg, tp.blob),
                         rel.reltype)
gf_el = copy.deepcopy(shape._element)
rId = new_slide.part.relate_to(new_cp, RT.CHART)
gf_el.find(qn("a:graphic") + "/" + qn("a:graphicData") + "/" + qn("c:chart")).set(qn("r:id"), rId)
new_slide.shapes._spTree.append(gf_el)
```

Grafted charts keep OLD theme colors baked in — flag for review.

**Grouped shapes** (`GROUP`, 6) — `shape.shapes` iterates members. Simple approach: flatten (recurse and copy members individually, offsetting by the group's transform). `group.shapes` members report coordinates in group-local space; for exact placement read the group's `a:xfrm` (`off`/`ext` vs `chOff`/`chExt`) and scale. Alternative: deep-copy the whole `grpSp` element into the new spTree (works because groups rarely reference rels; if members include pictures, their `r:embed` rIds must be re-related to image parts on the new slide first).

**SmartArt** — `GraphicFrame` whose `shape_type` returns `None` (verified: not chart/table/OLE URI). python-pptx has **no** SmartArt API. Options, in order: (1) graft the four diagram parts (`data`, `layout`, `quickStyle`, `colors` under `ppt/diagrams/`) plus the graphicFrame element with re-related rIds — same pattern as the chart graft but with 4 relationships; keeps old styling; (2) rasterize (render old slide, crop, insert as picture); (3) recreate as text/shapes. Always **flag for manual review** in the report.

**Video / audio** (`MEDIA`, 16) — `add_movie(io.BytesIO(part_blob), l, t, w, h, mime_type=...)` exists; get the source bytes from the shape's media relationship. Poster frame is regenerated. **OLE objects** — `shape.ole_format.blob/.prog_id`; re-add via `shapes.add_ole_object(...)` or graft the element + embedding part. Both: flag for review.

**Plain textboxes / autoshapes** — `add_textbox(l, t, w, h)` + `copy_text_frame`; for autoshapes, `add_shape(src.auto_shape_type, ...)`. Do not copy fill/line colors unless they carry meaning — let defaults come from the new theme's `fmtScheme`.

## 5. Speaker notes

```python
if old_slide.has_notes_slide:                       # check BEFORE .notes_slide —
    src_tf = old_slide.notes_slide.notes_text_frame # accessing it creates one
    if src_tf is not None and src_tf.text.strip():
        dst_tf = new_slide.notes_slide.notes_text_frame   # auto-creates from new notesMaster
        copy_text_frame(src_tf, dst_tf, keep_emphasis=False)  # notes = plain text
```

`notes_text_frame` can be `None` if the notes slide lacks a BODY placeholder. Only the notes body transfers; other notes-slide shapes (rare) are dropped.

## 6. Layout mapping methodology

Never guess per-slide. First **dump both templates**:

```python
def dump_layouts(prs):
    return [
        {"master": m.name, "layout": l.name,
         "placeholders": [{"idx": p.placeholder_format.idx,
                           "type": str(p.placeholder_format.type),
                           "name": p.name} for p in l.placeholders]}
        for m in prs.slide_masters for l in m.slide_layouts
    ]
```

Also dump the *old deck's actual usage*: for each slide, `slide.slide_layout.name` + which placeholder idxs hold content. Write both to `work/inventory_{old,new}.json`.

Then author `work/layout_map.json` by hand (this is a judgment call, not derivable):

```json
{
  "default_layout": "Title and Content",
  "map": {
    "Titel en object": {
      "new_layout": "Title and Content",
      "idx_map": {"0": 0, "1": 1},
      "notes": ""
    },
    "Twee objecten oud": {
      "new_layout": "Two Content",
      "idx_map": {"0": 0, "1": 1, "2": 2}
    },
    "Weird legacy layout": {
      "new_layout": "Title and Content",
      "idx_map": {"0": 0, "1": 1},
      "review": true
    }
  }
}
```

Rules: every layout name that appears in the old deck MUST have an entry (fail fast on a missing key rather than silently defaulting). Layouts with no reasonable counterpart map to `default_layout` with `"review": true`; the converter records slide numbers of every reviewed/fallback slide in `output/report.json`. Source placeholder idxs missing from `idx_map` are orphans: route their content to the body placeholder (appended paragraphs) or a textbox, and report them.

## 7. Theme / style gotchas

- **Direct color vs scheme color**: `font.color.type` is `MSO_COLOR_TYPE.RGB` (hard override — drop it) vs `.SCHEME` (theme slot — safe to carry, resolves against the new `clrScheme`). Same logic for shape `fill.fore_color`. Copying `.rgb` values defeats the conversion.
- **Fonts**: theme fonts appear in XML as `+mj-lt` (major/headings) and `+mn-lt` (minor/body); python-pptx reads an explicit `font.name` only — `None` means "inherits theme font". Drop every explicit `font.name` that equals the OLD theme's fonts (read them from old `theme1.xml`: `a:majorFont/a:latin/@typeface`); keep genuinely intentional odd fonts (code samples in a mono face) and flag them.
- **Bullets**: python-pptx has no bullet API. Bullet glyph/color/indent come from the master's `p:txStyles` per level — so setting `paragraph.level` correctly is the whole job. If an old deck used manual `a:buChar`/`a:buNone` overrides inside `a:pPr`, they hide in paragraph XML; strip them by not copying `pPr` (the `copy_text_frame` above never does). Exception worth honoring: `buNone` ("no bullet") may be semantic.
- **Slide size**: `prs.slide_width/height` are EMU (914400/inch); default 4:3 = 9144000 × 6858000, 16:9 = 12192000 × 6858000. Never set the size on the new deck — the new template's size is authoritative. Repositioning free shapes across a 4:3→16:9 change: uniform-scale by `min(w_ratio, h_ratio)` to preserve aspect, then re-center: `new_left = (new_w - shape_w*s)/2 + (old_left - old_w/2)*s` keeps horizontal composition; simple per-axis stretching distorts pictures. Placeholder content needs none of this — it lands where the new layout says.
- **Footer / date / slide-number**: not cloned onto slides (verified §1); they display when the layout/master shows them and the deck enables them. python-pptx 1.0.2 has **no headers/footers API** — fixed footer text must be correct in the new template's master/layouts themselves. If the old deck carried per-slide footer text in a cloned FOOTER placeholder, harvest it into the report; don't try to recreate it per-slide.
- **Multi-master templates**: `prs.slide_layouts` only sees master #1. Always resolve layouts via the inventory (master name + layout name).

## 8. QA checklist

Automate what you can into `tools/qa.py`; renders land in `output/`.

1. **Content completeness diff** (hard fail if violated): for each slide pair, compare (a) concatenated normalized text (all text frames incl. tables, whitespace-collapsed) — every source string must appear in the target; (b) picture count and blob hashes; (c) chart count; (d) notes text equality.
2. **Structure**: slide count equal; every slide's layout is the mapped one; `output/report.json` review-flags (SmartArt, grafted charts, fallback layouts, orphaned placeholders, dropped formatting) all triaged by a human.
3. **Overflow detection** (heuristic — python-pptx cannot measure rendered text): for each text placeholder, estimate `lines = ceil(len(text) / (width_emu / (font_size_pt * 0.55 * 12700)))` per paragraph and compare `lines * font_size_pt * 1.25 * 12700` against `ph.height` (resolve inherited size/geometry from the layout when the slide value is `None`; font size from master `txStyles` when runs are `None`). Flag anything > 90% full. `text_frame.fit_text()` exists but rewrites font sizes (needs font files) — prefer flagging over auto-shrinking.
4. **Visual compare**: render old and new decks to PNGs (LibreOffice headless: `soffice --headless --convert-to pdf` then `pdftoppm -png`, or the repo's existing render tooling — renders live under `output/`). Eyeball side-by-side pairs; per-slide image-diff scores (PIL `ImageChops.difference` on downscaled grayscale) are only good for *ranking* which slides to look at first — old vs new templates legitimately differ everywhere.
5. **Round-trip sanity**: reopen the output with python-pptx and with PowerPoint/LibreOffice once — a deck that opens clean in both has valid part wiring (graft steps are where corruption risk lives).

## 9. End-to-end workflow for this repo

1. **Collect** — old deck(s) + new template into `input/`. Confirm the new template's masters/layouts/theme are final (footers, logos, slide size).
2. **Inspect** — `tools/dump_inventory.py` → `work/inventory_old.json`, `work/inventory_new.json`, plus per-slide usage of the old deck (layout name, filled placeholder idxs, shape-type census incl. SmartArt/chart/OLE counts). Read both; note slide-size difference.
3. **Extract manifest** — `tools/extract_content.py` → `work/manifest.json`: per slide, all text (structured: placeholder idx → paragraphs → runs w/ emphasis), notes, media blob hashes. This is the QA baseline and survives independent of the source file.
4. **Author mapping** — write `work/layout_map.json` (§6) by hand from the two inventories. Every old layout gets an entry; ambiguous ones get `"review": true`.
5. **Rebuild** — `tools/convert.py input/old.pptx work/layout_map.json output/new.pptx`: new template as base, strip stub slides, per-slide add_slide + placeholder copy (§3), non-placeholder dispatch (§4), notes (§5); write `output/report.json` with every flag.
6. **Render** — old and new decks to `output/renders/{old,new}/NN.png`.
7. **QA loop** — run `tools/qa.py` (§8); fix converter or mapping; re-run 5–7 until the completeness diff is clean and every review flag is resolved (manually fixed in PowerPoint, or accepted). Manual post-edits happen ONLY after the last automated rebuild, on a copy — rerunning the converter overwrites `output/new.pptx`.

Keep every script idempotent (regenerate outputs from inputs; never mutate `input/`), and keep `work/layout_map.json` in git — it is the one hand-authored artifact.
