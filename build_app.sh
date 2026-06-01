#!/usr/bin/env bash
# Build tutor-diff as a macOS .app bundle using PyInstaller.
# Usage: bash build_app.sh
set -euo pipefail

APP_NAME="דוח פערים מתרגלים"
ZIP_NAME="דוח_פערים_מתרגלים.zip"
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

# Package as zip for sharing (ditto preserves macOS .app structure correctly)
echo "→ Creating zip for sharing…"
rm -f "dist/$ZIP_NAME"
ditto -c -k --keepParent "dist/$APP_NAME.app" "dist/$ZIP_NAME"

echo ""
echo "✓ הושלם!"
echo ""
echo "  קובץ לשליחה:  dist/$ZIP_NAME"
echo ""
echo "  הוראות לנמענת:"
echo "  1. פתחי את ה-zip"
echo "  2. לחצי פעמיים על האפליקציה"
echo "  3. אם Mac מבקש אישור: לחצי קליק ימני → פתח → פתח"
