#!/usr/bin/env bash
# Wrapper for validate_pdf.py — validates page count and basic integrity of a PDF.
# Usage: ./validate_pdf.sh <pdf_path> [expected_pages]
set -euo pipefail

PDF_PATH="${1:?Usage: validate_pdf.sh <pdf_path> [expected_pages]}"
EXPECTED_PAGES="${2:-16}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python "$SCRIPT_DIR/validate_pdf.py" "$PDF_PATH" "$EXPECTED_PAGES"