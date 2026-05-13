#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

PYTHON_BIN="${JITP_PYTHON:-python3}"
OUTPUT_DIR="${JITP_OUTPUT_DIR:-$REPO_ROOT/outputs/jitp_manifold_metadata}"
EXISTING_JSON="${JITP_EXISTING_JSON:-$REPO_ROOT/data/jitp_manifold_metadata.json}"

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "Python runtime not found: $PYTHON_BIN" >&2
  exit 1
fi

mkdir -p "$OUTPUT_DIR"

"$PYTHON_BIN" "$REPO_ROOT/scrape_jitp_manifold.py" \
  --output-dir "$OUTPUT_DIR" \
  --existing-json "$EXISTING_JSON" \
  "$@"

"$PYTHON_BIN" "$REPO_ROOT/build_jitp_metadata_workbook.py" \
  --input-json "$OUTPUT_DIR/jitp_manifold_metadata.json" \
  --output-dir "$OUTPUT_DIR"

echo "Updated JSON: $OUTPUT_DIR/jitp_manifold_metadata.json"
echo "Updated workbook: $OUTPUT_DIR/jitp_manifold_article_metadata.xlsx"
