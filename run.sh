#!/usr/bin/env bash
set -euo pipefail

# Convenience wrapper for running the CV screener.
#
# CLI mode (process PDFs → Excel):
#   ./run.sh ./path/to/cvs ./output.xlsx [criteria.yaml]
#
# Web mode (start web UI):
#   ./run.sh --web [port]
#   ./run.sh --web 9090

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ---- Web mode ----
if [ "${1:-}" = "--web" ]; then
    PORT="${2:-8080}"
    echo "🔨 Building Docker image..."
    docker build -t cv-screener-web -f "$SCRIPT_DIR/Dockerfile.web" "$SCRIPT_DIR" > /dev/null
    echo "🚀 Starting web UI on http://0.0.0.0:$PORT"
    echo ""
    docker run --rm -p "$PORT:8080" cv-screener-web
    exit 0
fi

# ---- CLI mode ----
INPUT_DIR="${1:-}"
OUTPUT="${2:-/output/results.xlsx}"
CRITERIA="${3:-$SCRIPT_DIR/criteria.yaml}"

if [ -z "$INPUT_DIR" ]; then
    echo "Usage:"
    echo "  $0 <input_dir> [output_path] [criteria_yaml]    # CLI mode"
    echo "  $0 --web [port]                                  # Web UI mode"
    echo ""
    echo "Examples:"
    echo "  $0 ./cvs ./results.xlsx"
    echo "  $0 /absolute/path/to/cvs /tmp/results.xlsx my_criteria.yaml"
    echo "  $0 --web 8080"
    exit 1
fi

# Resolve to absolute paths
INPUT_ABS="$(realpath "$INPUT_DIR")"
OUTPUT_ABS="$(realpath "$(dirname "$OUTPUT")")/$(basename "$OUTPUT")"
CRITERIA_ABS="$(realpath "$CRITERIA")"

echo "🔨 Building Docker image..."
docker build -t cv-screener "$SCRIPT_DIR" > /dev/null

echo "🚀 Running CV screener..."
echo "   Input:   $INPUT_ABS"
echo "   Output:  $OUTPUT_ABS"
echo "   Criteria: $CRITERIA_ABS"
echo ""

docker run --rm \
    -v "$INPUT_ABS":/input:ro \
    -v "$(dirname "$OUTPUT_ABS")":/output \
    -v "$CRITERIA_ABS":/app/criteria.yaml:ro \
    cv-screener \
    --input /input --output "$OUTPUT" --criteria /app/criteria.yaml --workers "$(nproc)"
