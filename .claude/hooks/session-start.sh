#!/bin/bash
# Installs everything needed to build and QA the deck.
# Idempotent: every step is a no-op once its dependency is present.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-$(dirname "$0")/../..}"

# 1. pptxgenjs — the deck generator
npm install --no-audit --no-fund

# 2. Python side: OOXML validation, content dump, PDF rasterizing
pip install --quiet --disable-pip-version-check \
  defusedxml lxml Pillow pymupdf "markitdown[pptx]"

# 3. LibreOffice Impress + a CJK font, for rendering previews.
# libreoffice-core ships without the Impress filters, so conversion fails
# with "source file could not be loaded" until impress is installed.
# Optional: the build and validation steps work without it.
if [ ! -f /usr/lib/libreoffice/program/simpress ]; then
  if command -v apt-get >/dev/null; then
    export DEBIAN_FRONTEND=noninteractive
    apt-get update -qq || true
    apt-get install -y -qq libreoffice-impress fonts-wqy-zenhei \
      || echo "warning: LibreOffice Impress unavailable — 'npm run preview' will not work"
  fi
fi

echo "environment ready: npm run build | npm run qa | npm run preview"
