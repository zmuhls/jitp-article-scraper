#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

PYTHON_BIN="${JITP_PYTHON:-/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3}"
NODE_BIN="${JITP_NODE:-/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node}"
NODE_MODULES="${JITP_NODE_MODULES:-/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules}"
OUTPUT_DIR="${JITP_OUTPUT_DIR:-$REPO_ROOT/outputs/jitp_manifold_metadata}"
EXISTING_JSON="${JITP_EXISTING_JSON:-$REPO_ROOT/data/jitp_manifold_metadata.json}"

if [[ ! -x "$PYTHON_BIN" ]]; then
  echo "Python runtime not found or not executable: $PYTHON_BIN" >&2
  exit 1
fi

if [[ ! -x "$NODE_BIN" ]]; then
  echo "Node runtime not found or not executable: $NODE_BIN" >&2
  exit 1
fi

if [[ ! -e "$REPO_ROOT/node_modules" ]]; then
  if [[ -d "$NODE_MODULES" ]]; then
    ln -s "$NODE_MODULES" "$REPO_ROOT/node_modules"
  else
    echo "Node modules directory not found: $NODE_MODULES" >&2
    exit 1
  fi
fi

mkdir -p "$OUTPUT_DIR"

"$PYTHON_BIN" "$REPO_ROOT/scrape_jitp_manifold.py" \
  --output-dir "$OUTPUT_DIR" \
  --existing-json "$EXISTING_JSON" \
  "$@"

"$NODE_BIN" "$REPO_ROOT/build_jitp_metadata_workbook.mjs" \
  --input-json "$OUTPUT_DIR/jitp_manifold_metadata.json" \
  --output-dir "$OUTPUT_DIR"

echo "Updated JSON: $OUTPUT_DIR/jitp_manifold_metadata.json"
echo "Updated workbook: $OUTPUT_DIR/jitp_manifold_article_metadata.xlsx"
