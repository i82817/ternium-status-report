#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

LOG="$SCRIPT_DIR/rebuild.log"
echo "=== $(date) ===" >> "$LOG"

uv run --with openpyxl python3 -I rebuild.py >> "$LOG" 2>&1

if ! git rev-parse --is-inside-work-tree > /dev/null 2>&1; then
  echo "Not a git repo — skipping push" | tee -a "$LOG"
  exit 0
fi

git add index.html
if git diff --cached --quiet; then
  echo "No changes to commit" | tee -a "$LOG"
  exit 0
fi

git commit -m "chore: auto-refresh report $(date +%Y-%m-%d)"
git push
echo "Pushed ✓" | tee -a "$LOG"
