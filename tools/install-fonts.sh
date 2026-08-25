#!/usr/bin/env bash
# Install fonts user-space so slide renders are faithful (no sudo needed).
# 1. Extracts any fonts EMBEDDED in the given .pptx files (ppt/fonts/*.fntdata are
#    plain TTF/OTF payloads) into ~/.local/share/fonts/pptconv/.
# 2. Reports theme typefaces each deck asks for vs what fontconfig can resolve,
#    so missing families are explicit (drop matching .ttf/.otf into the same dir).
set -euo pipefail
DEST="$HOME/.local/share/fonts/pptconv"
mkdir -p "$DEST"

for deck in "$@"; do
  echo "== $deck"
  tmp=$(mktemp -d)
  unzip -q -o "$deck" 'ppt/fonts/*' -d "$tmp" 2>/dev/null || true
  found=0
  for f in "$tmp"/ppt/fonts/*; do
    [ -e "$f" ] || continue
    base=$(basename "${f%.fntdata}")
    cp "$f" "$DEST/${base}.ttf"
    found=$((found+1))
  done
  echo "  embedded fonts extracted: $found"
  echo "  theme typefaces requested:"
  unzip -p "$deck" 'ppt/theme/theme1.xml' 2>/dev/null \
    | grep -o 'typeface="[^"+][^"]*"' | sort -u | sed 's/typeface=/    /; s/"//g' \
    | while read -r face; do
        if fc-match -q "$face" 2>/dev/null && [ "$(fc-match "$face" | cut -d: -f2 | xargs)" != "" ]; then
          match=$(fc-match "$face" | sed 's/:.*//')
          echo "    $face -> $match"
        fi
      done
  rm -rf "$tmp"
done

fc-cache -f "$DEST" >/dev/null 2>&1 || true
echo
echo "fontconfig now knows $(fc-list | wc -l) fonts. If a theme face maps to a wrong"
echo "substitute above, drop the real .ttf/.otf into $DEST and re-run fc-cache -f."
