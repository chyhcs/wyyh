#!/bin/bash
# Build the deck, then run the three checks that catch real defects:
#   1. OOXML validation  — PowerPoint refuses some XML that every other tool accepts
#   2. content dump      — verifies nothing was dropped or left as placeholder text
#   3. preview render    — the only way to see overflow and overlap
set -euo pipefail

cd "$(dirname "$0")/.."

echo "==> build"
node build_deck.js

DECK=$(ls -- *.pptx | head -1)

echo "==> validate"
VALIDATE=/root/.claude/skills/pptx/scripts/office/validate.py
if [ -f "$VALIDATE" ]; then
  python3 "$VALIDATE" "$DECK"
else
  echo "skipped: pptx skill validator not present in this environment"
fi

echo "==> content"
markitdown "$DECK" | grep -iE '\bx{3,}\b|lorem|ipsum|\bTODO|\[insert' \
  && { echo "placeholder text found"; exit 1; } || echo "no placeholder text"

echo "==> render"
./scripts/render.sh

echo "==> QA complete — inspect preview-32x9.png for overflow and overlap"
