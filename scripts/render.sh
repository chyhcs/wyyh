#!/bin/bash
# Render the deck to preview-32x9.png via LibreOffice + PyMuPDF.
# LibreOffice chokes on the deck's non-ASCII filename, so convert through an ASCII copy.
set -euo pipefail

cd "$(dirname "$0")/.."

DECK=$(ls -- *.pptx | head -1)
[ -n "$DECK" ] || { echo "no .pptx found — run 'npm run build' first" >&2; exit 1; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
cp -- "$DECK" "$TMP/deck.pptx"

soffice --headless -env:UserInstallation="file://$TMP/lo" \
  --convert-to pdf --outdir "$TMP" "$TMP/deck.pptx" >/dev/null 2>&1

python3 - "$TMP/deck.pdf" <<'PY'
import sys, pymupdf
doc = pymupdf.open(sys.argv[1])
for i, page in enumerate(doc):
    out = "preview-32x9.png" if len(doc) == 1 else f"preview-32x9-{i+1}.png"
    page.get_pixmap(dpi=150).save(out)
    print(f"wrote {out}  ({page.rect.width:.0f}x{page.rect.height:.0f}pt)")
PY
