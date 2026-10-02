#!/usr/bin/env bash
# Ask Word where widow control and keepLines move lines (keeps_docx(),
# docs/reference/acceptance-real-docx.md): writes the fixture, has Word save it as PDF
# (through word-pagination-probe.sh, which owns the Office automation
# traps) and prints, per page, the line labels Word put there, as the JSON
# fixtures/keeps.word.json records.
#
#   bash scripts/word-keeps-probe.sh <out-dir>
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/keeps.docx" keeps
bash scripts/word-pagination-probe.sh "$OUT/keeps.docx" "$OUT/keeps.pdf" >&2
pages=$(pdfinfo "$OUT/keeps.pdf" | awk '/^Pages:/{print $2}')
{
  echo '{'
  echo ' "source": "Microsoft Word 16 (macOS), keeps_docx() saved as PDF by scripts/word-keeps-probe.sh; per page, its lines top to bottom",'
  echo ' "pages": ['
  for p in $(seq 1 "$pages"); do
    sep=$([ "$p" -lt "$pages" ] && echo , || true)
    pdftotext -f "$p" -l "$p" -layout "$OUT/keeps.pdf" - | python3 -c '
import json, sys
print("  " + json.dumps([l.strip() for l in sys.stdin.read().split("\n") if l.strip()]), end="")'
    echo "$sep"
  done
  echo ' ]'
  echo '}'
}
