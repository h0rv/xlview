#!/usr/bin/env bash
set -euo pipefail

# Tight feedback loop for the scroll/header regression:
# - optionally rebuild WASM
# - run the headless browser regression test
# - pause for edits, then rerun
#
# Usage:
#   bash scripts/scroll_loop.sh
#   BUILD=0 bash scripts/scroll_loop.sh        # skip wasm build
#   TEST_CMD="node tests/browser/run_scroll_test.js" bash scripts/scroll_loop.sh

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
BUILD="${BUILD:-1}"
TEST_CMD="${TEST_CMD:-node tests/browser/run_scroll_test.js}"

cd "$ROOT"

while true; do
  if [[ "$BUILD" == "1" ]]; then
    npm run build:wasm:dev
  fi

  # Run the regression test; don't exit the loop on failure.
  set +e
  eval "$TEST_CMD"
  set -e

  echo
  read -r -p "Fix code, then press Enter to rerun (q to quit): " ans
  if [[ "${ans}" == "q" || "${ans}" == "Q" ]]; then
    break
  fi
done
