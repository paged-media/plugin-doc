#!/usr/bin/env bash
# Ask Word where every kind of break puts the text that follows it (thoughts
# ADR 028 / 029). Builds docx_conformance::breaks_docx(), has Word save it as
# PDF (through word-pagination-probe.sh, which owns the Office automation
# traps) and prints Word's page map: per page its size, and per column the
# labels of the lines on it (every fixture paragraph's first word is its
# label), plus how far below the top margin the first line sits (a blank
# line at the top of a page shows there, not in the labels).
#
#   bash scripts/word-breaks-probe.sh <out-dir>
#
# fixtures/breaks.word.json is that output plus provenance.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/breaks.docx" breaks
bash scripts/word-pagination-probe.sh "$OUT/breaks.docx" "$OUT/breaks.pdf"
echo "fonts: $(pdffonts "$OUT/breaks.pdf" | awk 'NR>2 {sub(/^[A-Z]+\+/, "", $1); print $1}' | sort -u | tr '\n' ' ')" >&2
pdftotext -bbox "$OUT/breaks.pdf" "$OUT/breaks.bbox.html"
python3 - "$OUT/breaks.bbox.html" <<'PY'
import json, re, sys
TOP_MARGIN = 36.0  # breaks_docx(): 0.5 in on every page
html = open(sys.argv[1]).read()
pages = []
for m in re.finditer(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', html, re.S):
    w, h, body = float(m.group(1)), float(m.group(2)), m.group(3)
    words = [
        (float(x0), float(y0), float(y1), t)
        for x0, y0, y1, t in re.findall(
            r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)</word>', body
        )
    ]
    # A line = consecutive words on one baseline; its column is where its
    # FIRST word starts (left or right of the page's middle).
    lines = []
    for x0, y0, y1, t in words:
        if lines and abs(lines[-1]["y1"] - y1) < 0.5 and x0 > lines[-1]["x0"]:
            continue
        lines.append({"x0": x0, "y0": y0, "y1": y1, "label": t, "col": int(x0 > w / 2)})
    cols = []
    for c in sorted({l["col"] for l in lines}):
        cols.append([l["label"] for l in lines if l["col"] == c])
    page = {"page": len(pages) + 1, "size_pt": [round(w, 2), round(h, 2)]}
    if not lines:
        page["blank"] = True
    else:
        page["columns" if len(cols) > 1 else "labels"] = cols if len(cols) > 1 else cols[0]
        page["first_line_top_pt"] = round(lines[0]["y0"] - TOP_MARGIN, 2)
    pages.append(page)
print(json.dumps({"page_count": len(pages), "pages": pages}, indent=1))
PY
