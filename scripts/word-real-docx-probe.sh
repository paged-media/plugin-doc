#!/usr/bin/env bash
# Ask Word how it lays the constructs the real corpus documents exposed
# (docs/reference/acceptance-real-docx.md): merged tab stops past the margin, a style
# based on a character style, a style's own list, an unstyled paragraph and
# a legacy VML inline picture. Writes real_docx_docx(), has Word save it as
# PDF (through word-pagination-probe.sh, which owns the Office automation
# traps), and prints every word with its box (pt from the page's top-left)
# as the JSON fixtures/real-docx.word.json records. The text area is
# x = 36..324.
#
#   bash scripts/word-real-docx-probe.sh <out-dir>
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/real-docx.docx" real-docx
bash scripts/word-pagination-probe.sh "$OUT/real-docx.docx" "$OUT/real-docx.pdf" >&2
pdftotext -bbox "$OUT/real-docx.pdf" "$OUT/real-docx.bbox.html"
python3 - "$OUT/real-docx.bbox.html" <<'PY'
import json, re, sys
html = open(sys.argv[1]).read()
words = [
    {"t": t, "x0": round(float(x0), 2), "y0": round(float(y0), 2),
     "x1": round(float(x1), 2), "y1": round(float(y1), 2)}
    for x0, y0, x1, y1, t in re.findall(
        r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', html
    )
]
print("{\n \"source\": \"Microsoft Word 16 (macOS), real_docx_docx() saved as PDF by scripts/word-real-docx-probe.sh; word boxes in pt from the page top-left (pdftotext -bbox)\",\n \"words\": [\n" + ",\n".join("  " + json.dumps(w, ensure_ascii=False) for w in words) + "\n ]\n}")
PY
