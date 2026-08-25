#!/usr/bin/env python
"""Visual QA: build side-by-side comparison images (old deck vs converted deck).

Produces, in <outdir>:
  pair-NN.png     old slide N (left) | new slide N (right) — for eyeball review
  contact.png     grid of all new slides — quick whole-deck impression

The intended loop: render both decks with tools/render.sh, run qa.py, then LOOK at
every pair image (Claude reads these natively) and fix what's off. No slide ships
unseen.

Usage: qa.py <old_render_dir> <new_render_dir> <outdir> [--offset N]
  --offset N  new deck slide N+i pairs with old slide i (when slides were dropped/added)
"""
from __future__ import annotations
import sys
from pathlib import Path

from PIL import Image, ImageDraw

GAP = 24
BG = (245, 242, 235)
LABEL_H = 34


def slides_in(d):
    return sorted(Path(d).glob("slide-*.png"))


def pair(old_png, new_png, out_png, n):
    imgs = []
    for p in (old_png, new_png):
        imgs.append(Image.open(p).convert("RGB") if p and p.exists() else None)
    h = max(i.height for i in imgs if i) if any(imgs) else 400
    canvases = []
    for i in imgs:
        if i is None:
            ph = Image.new("RGB", (int(h * 16 / 9), h), (220, 210, 210))
            ImageDraw.Draw(ph).text((30, 30), "MISSING", fill=(120, 0, 0))
            canvases.append(ph)
        else:
            canvases.append(i.resize((int(i.width * h / i.height), h)))
    w = sum(c.width for c in canvases) + GAP * 3
    sheet = Image.new("RGB", (w, h + LABEL_H + GAP * 2), BG)
    d = ImageDraw.Draw(sheet)
    d.text((GAP, 8), f"slide {n}   OLD (left)  vs  NEW (right)", fill=(60, 50, 40))
    x = GAP
    for c in canvases:
        sheet.paste(c, (x, LABEL_H + GAP))
        x += c.width + GAP
    sheet.save(out_png)


def contact(pngs, out_png, cols=4, thumb_w=420):
    if not pngs:
        return
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        thumbs.append(im.resize((thumb_w, int(im.height * thumb_w / im.width))))
    th = thumbs[0].height
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (thumb_w + GAP) + GAP, rows * (th + GAP + 20) + GAP), BG)
    d = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        x = GAP + (i % cols) * (thumb_w + GAP)
        y = GAP + (i // cols) * (th + GAP + 20)
        sheet.paste(t, (x, y + 18))
        d.text((x, y), f"{i + 1}", fill=(60, 50, 40))
    sheet.save(out_png)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 3:
        sys.exit(__doc__)
    offset = 0
    if "--offset" in sys.argv:
        offset = int(sys.argv[sys.argv.index("--offset") + 1])
    old_dir, new_dir, outdir = args
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    old, new = slides_in(old_dir), slides_in(new_dir)
    n = max(len(old), len(new) + offset)
    for i in range(1, n + 1):
        o = old[i - 1] if i <= len(old) else None
        ni = i - offset
        nw = new[ni - 1] if 1 <= ni <= len(new) else None
        pair(o, nw, out / f"pair-{i:02}.png", i)
    contact(new, out / "contact.png")
    print(f"{n} pair images + contact sheet -> {out}/  (now LOOK at every pair-NN.png)")


if __name__ == "__main__":
    main()
