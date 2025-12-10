#!/usr/bin/env bash
set -euo pipefail

CSV_FILE="${1:-util/print-certificates.csv}"

pick_browser() {
  if [ -n "${BROWSER:-}" ] && command -v "$BROWSER" >/dev/null 2>&1; then
    echo "$BROWSER"
    return
  fi

  for candidate in chromium chromium-browser google-chrome-stable google-chrome; do
    if command -v "$candidate" >/dev/null 2>&1; then
      echo "$candidate"
      return
    fi
  done

  echo ""
}

browser_bin="$(pick_browser)"
if [ -z "$browser_bin" ]; then
  echo "No headless Chromium/Chrome found. Install chromium or set BROWSER=<path>." >&2
  exit 1
fi

if [ ! -f "$CSV_FILE" ]; then
  echo "CSV file not found: $CSV_FILE" >&2
  exit 1
fi

realpath_f() {
  if command -v realpath >/dev/null 2>&1; then
    realpath "$1" 2>/dev/null || python - "$1" <<'PY'
import os, sys
print(os.path.abspath(sys.argv[1]))
PY
  else
    python - "$1" <<'PY'
import os, sys
print(os.path.abspath(sys.argv[1]))
PY
  fi
}

while IFS=, read -r html_path pdf_path; do
  # Skip empty lines and comments
  if [ -z "${html_path// }" ] || [[ "$html_path" =~ ^# ]]; then
    continue
  fi

  # Strip surrounding quotes if present
  html_path="${html_path%\"}"
  html_path="${html_path#\"}"
  pdf_path="${pdf_path%\"}"
  pdf_path="${pdf_path#\"}"

  abs_html="$(realpath_f "$html_path")"
  abs_pdf="$(realpath_f "$pdf_path")"
  out_dir="$(dirname "$abs_pdf")"

  if [ ! -f "$abs_html" ]; then
    echo "HTML not found, skipping: $abs_html" >&2
    continue
  fi

  mkdir -p "$out_dir"

  echo "Printing $abs_html -> $abs_pdf"
  "$browser_bin" \
    --headless \
    --disable-gpu \
    --no-sandbox \
    --print-to-pdf="$abs_pdf" \
    "file://$abs_html"
done < "$CSV_FILE"
