#!/usr/bin/env bash
# Ask Word what an EDITED save-back of tabs, non-breaking hyphens and line
# breaks looks like. Writes the original run_specials_docx() and the edited
# save the conformance test makes (docx-conformance/tests/run_specials.rs,
# EDITS), has Word save BOTH as PDF (through word-pagination-probe.sh, which
# owns the Office automation traps — a repair prompt shows up as a hang and
# no PDF), and prints, per line, every word and where it starts (pt from the
# page's left edge). Tab stops sit at 72 pt and 144 pt from the left margin
# (36 pt), i.e. at x = 108 and 180.
#
#   bash scripts/word-run-specials-probe.sh <out-dir>
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/run-specials.docx" run-specials
RUN_SPECIALS_EDITED_OUT="$OUT/run-specials-edited.docx" \
  cargo test -q -p docx-conformance --test run_specials an_edited_save >&2
for f in run-specials run-specials-edited; do
  bash scripts/word-pagination-probe.sh "$OUT/$f.docx" "$OUT/$f.pdf" >&2
  pdftotext -bbox "$OUT/$f.pdf" "$OUT/$f.bbox.html"
  echo "== $f"
  python3 - "$OUT/$f.bbox.html" <<'PY'
import re, sys
html = open(sys.argv[1]).read()
words = sorted(
    (round(float(y1)), float(x0), t)
    for x0, y1, t in re.findall(
        r'xMin="([\d.]+)" yMin="[\d.]+" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)</word>', html
    )
)
lines = {}
for y, x, t in words:
    lines.setdefault(y, []).append(f"{t}@{x:.1f}")
for y in sorted(lines):
    print("  " + "  ".join(lines[y]))
PY
done
