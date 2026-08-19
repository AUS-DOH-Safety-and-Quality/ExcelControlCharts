#!/bin/bash
# Unregisters the ExcelControlCharts add-in and removes its manifest.
# Double-click in Finder, or pipe from the web:
#   curl -fsSL https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/uninstall.command | bash

set -euo pipefail

DATA_DIR="$HOME/Library/Application Support/ExcelControlCharts"
MANIFEST_PATH="$DATA_DIR/manifest.xml"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

if [ -f "$MANIFEST_PATH" ]; then
  ADDIN_ID=$(sed -n 's/.*<Id>\(.*\)<\/Id>.*/\1/p' "$MANIFEST_PATH" | head -1)
  if [ -n "$ADDIN_ID" ]; then
    rm -f "$WEF_DIR/$ADDIN_ID."*
  fi
fi

rm -rf "$DATA_DIR"

echo "Unregistered the add-in. Restart Excel to finish removing it."
