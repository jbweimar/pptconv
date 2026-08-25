#!/usr/bin/env bash
# render.sh <deck.pptx> <outdir>
# Converts a .pptx to PDF via headless LibreOffice, then rasterizes each
# page to <outdir>/slide-NN.png with pdftoppm at 110 dpi.
set -euo pipefail

SOFFICE=/home/jbweimar/projects/pptconv/vendor/libreoffice/opt/libreoffice25.8/program/soffice
PDFTOPPM=/usr/bin/pdftoppm
LO_PROFILE="file:///tmp/lo-pptconv"

if [[ $# -ne 2 ]]; then
  echo "usage: $(basename "$0") <deck.pptx> <outdir>" >&2
  exit 2
fi

deck=$(readlink -f "$1")
outdir=$2
[[ -f $deck ]] || { echo "error: no such file: $deck" >&2; exit 1; }
mkdir -p "$outdir"
outdir=$(readlink -f "$outdir")

# Convert to PDF in an isolated temp dir so parallel runs / stale outputs
# cannot collide; dedicated UserInstallation avoids profile lock issues.
tmpdir=$(mktemp -d)
trap 'rm -rf "$tmpdir"' EXIT

"$SOFFICE" --headless --norestore \
  -env:UserInstallation="$LO_PROFILE" \
  --convert-to pdf --outdir "$tmpdir" "$deck" >/dev/null

base=$(basename "$deck")
pdf="$tmpdir/${base%.*}.pdf"
[[ -f $pdf ]] || { echo "error: PDF conversion failed for $deck" >&2; exit 1; }

"$PDFTOPPM" -r 110 -png "$pdf" "$outdir/slide"

# pdftoppm names files slide-1.png / slide-01.png depending on page count;
# normalize to zero-padded two-digit slide-NN.png.
shopt -s nullglob
for f in "$outdir"/slide-*.png; do
  n=${f##*/slide-}; n=${n%.png}
  printf -v padded 'slide-%02d.png' "$((10#$n))"
  [[ $(basename "$f") == "$padded" ]] || mv "$f" "$outdir/$padded"
done

ls "$outdir"/slide-*.png
