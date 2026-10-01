#!/usr/bin/env bash
# Ask Word what symbols (<w:sym>), optional hyphens (<w:softHyphen/>),
# absolute-position tabs (<w:ptab>) and empty table cells look like.
# Writes symbols_docx() and the edited save the conformance test makes
# (docx-conformance/tests/symbols.rs, EDITS), has Word save BOTH as PDF
# (through word-pagination-probe.sh, which owns the Office automation traps
# -- a repair prompt shows up as a hang and no PDF), and prints, per line,
# its baseline y and every word with where it starts and ends (pt from the
# page's top-left). The text area is x = 36..324.
#
#   bash scripts/word-symbols-probe.sh <out-dir>
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/symbols.docx" symbols
SYMBOLS_EDITED_OUT="$OUT/symbols-edited.docx" \
  cargo test -q -p docx-conformance --test symbols an_edited_save >&2
for f in symbols symbols-edited; do
  bash scripts/word-pagination-probe.sh "$OUT/$f.docx" "$OUT/$f.pdf" >&2
  pdftotext -bbox "$OUT/$f.pdf" "$OUT/$f.bbox.html"
  echo "== $f"
  python3 - "$OUT/$f.bbox.html" <<'PY'
import re, sys
html = open(sys.argv[1]).read()
words = sorted(
    (round(float(y1), 1), float(x0), float(x1), t)
    for x0, x1, y1, t in re.findall(
        r'xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', html
    )
)
lines = {}
for y, x0, x1, t in words:
    lines.setdefault(round(y), []).append(f"{t}@{x0:.1f}-{x1:.1f}")
for y in sorted(lines):
    print(f"  y={y}: " + "  ".join(lines[y]))
PY
done
