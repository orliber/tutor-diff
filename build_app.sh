#!/usr/bin/env bash
# Build tutor-diff as a macOS .app bundle using PyInstaller.
# Usage: bash build_app.sh
set -euo pipefail

APP_NAME="דוח פערים מתרגלים"
VENV_DIR=".venv"

# Create/reuse virtualenv
if [ ! -d "$VENV_DIR" ]; then
    echo "→ Creating virtualenv…"
    python3 -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

echo "→ Installing dependencies…"
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt pyinstaller

echo "→ Building .app bundle…"
pyinstaller \
    --name        "$APP_NAME" \
    --windowed \
    --onedir \
    --noconfirm \
    --add-data    "build_excel.py:." \
    --hidden-import "openpyxl.cell._writer" \
    --hidden-import "openpyxl.styles.differential" \
    app.py

echo ""
echo "✓ Done — bundle is at:  dist/$APP_NAME.app"
echo "  Drag it to /Applications to install."
