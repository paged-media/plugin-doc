#!/usr/bin/env bash
# Ask Word how it lays out plain line breaks and blank lines (core ab383b1 /
# 65cf615). Builds docx_conformance::line_breaks_docx(), has Word save it as
# PDF (through word-pagination-probe.sh, which owns the Office automation
# traps) and prints, per page, every visible line: its label (first word)
# and how far its top sits below the top margin. Blank lines (an empty
# paragraph, an empty line between two <w:br/>) carry no words; they show as
# the gaps between those positions.
#
#   bash scripts/word-line-breaks-probe.sh <out-dir>
#
# fixtures/line-breaks.word.json is that output plus provenance.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/line-breaks.docx" line-breaks
bash scripts/word-pagination-probe.sh "$OUT/line-breaks.docx" "$OUT/line-breaks.pdf" >&2
echo "fonts: $(pdffonts "$OUT/line-breaks.pdf" | awk 'NR>2 {sub(/^[A-Z]+\+/, "", $1); print $1}' | sort -u | tr '\n' ' ')" >&2
pdftotext -bbox "$OUT/line-breaks.pdf" "$OUT/line-breaks.bbox.html"
python3 - "$OUT/line-breaks.bbox.html" <<'PY'
import json, re, sys
TOP_MARGIN = 36.0  # line_breaks_docx(): 0.5 in
html = open(sys.argv[1]).read()
pages = []
for m in re.finditer(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', html, re.S):
    w, h, body = float(m.group(1)), float(m.group(2)), m.group(3)
    lines = []
    for x0, y0, y1, t in re.findall(
        r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)</word>', body
    ):
        x0, y0, y1 = float(x0), float(y0), float(y1)
        if lines and abs(lines[-1]["y1"] - y1) < 0.5 and x0 > lines[-1]["x0"]:
            continue
        lines.append({"x0": x0, "y1": y1, "label": t, "top_pt": round(y0 - TOP_MARGIN, 2)})
    pages.append({
        "page": len(pages) + 1,
        "size_pt": [round(w, 2), round(h, 2)],
        "lines": [{"label": l["label"], "top_pt": l["top_pt"]} for l in lines],
    })
print(json.dumps({"page_count": len(pages), "pages": pages}, indent=1))
PY
