#!/bin/bash
# Registers the deployed ExcelControlCharts add-in for sideloading in Excel.
# Double-click in Finder, or pipe from the web:
#   curl -fsSL https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/install.command | bash

set -euo pipefail

BASE_URL="https://aus-doh-safety-and-quality.github.io/ExcelControlCharts"

DATA_DIR="$HOME/Library/Application Support/ExcelControlCharts"
MANIFEST_PATH="$DATA_DIR/manifest.xml"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

mkdir -p "$DATA_DIR"

# The published manifest already points at the deployment, so no rewriting is needed.
curl -fsSL "$BASE_URL/manifest.xml" -o "$MANIFEST_PATH"

ADDIN_ID=$(sed -n 's/.*<Id>\(.*\)<\/Id>.*/\1/p' "$MANIFEST_PATH" | head -1)
if [ -z "$ADDIN_ID" ]; then
  echo "Could not read <Id> from $MANIFEST_PATH." >&2
  exit 1
fi

# Excel picks up any manifest in its sideload directory, named "<id>.<filename>".
mkdir -p "$WEF_DIR"
SIDELOAD_PATH="$WEF_DIR/$ADDIN_ID.manifest.xml"
rm -f "$SIDELOAD_PATH"
ln "$MANIFEST_PATH" "$SIDELOAD_PATH" 2>/dev/null || cp "$MANIFEST_PATH" "$SIDELOAD_PATH"

echo "Registered the add-in from $BASE_URL/"
echo "Restart Excel and choose it from the Home tab."
