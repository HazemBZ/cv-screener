#!/usr/bin/env bash
set -euo pipefail

# Quick test: process sample PDFs without Docker.
# This auto-detects the project dir and runs via uv.
#
# Usage:
#   ./test.sh                     # test samples/ → /tmp/results.xlsx
#   ./test.sh ./my_cvs out.xlsx   # test custom dir

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INPUT="${1:-$SCRIPT_DIR/samples}"
OUTPUT="${2:-/tmp/results.xlsx}"

if [ ! -d "$INPUT" ]; then
    echo "❌ Input directory not found: $INPUT"
    echo "Usage: $0 [input_dir] [output_path]"
    echo ""
    echo "  If no arguments given, tests samples/ directory."
    echo "  Create sample CVs by placing PDFs in samples/"
    exit 1
fi

echo "📄 Testing with CVs from: $INPUT"
echo "📊 Output: $OUTPUT"
echo ""

cd "$SCRIPT_DIR"

# Use uv run if available, fallback to pip
if command -v uv &>/dev/null; then
    uv run python cv_screener.py \
        --input "$INPUT" \
        --output "$OUTPUT" \
        --criteria "$SCRIPT_DIR/criteria.yaml" \
        --workers "$(nproc)"
else
    source .venv/bin/activate 2>/dev/null || true
    python cv_screener.py \
        --input "$INPUT" \
        --output "$OUTPUT" \
        --criteria "$SCRIPT_DIR/criteria.yaml" \
        --workers "$(nproc)"
fi
