#!/usr/bin/env bash
# Ask Word where a CONTINUOUS section (and a nextColumn one) puts its text
# (thoughts ADR 029). Builds docx_conformance::continuous_docx(), has Word
# save it as PDF (through word-pagination-probe.sh, which owns the Office
# automation traps) and prints Word's page map: per page its size and, per
# line, its label (every fixture paragraph's first word) with the left edge
# and top of its first word in points from the page's top-left corner
# (pdftotext -bbox), so a line's column and the margins in force show.
#
#   bash scripts/word-continuous-probe.sh <out-dir>
#
# fixtures/continuous.word.json is that output plus provenance.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/continuous.docx" continuous
bash scripts/word-pagination-probe.sh "$OUT/continuous.docx" "$OUT/continuous.pdf"
echo "fonts: $(pdffonts "$OUT/continuous.pdf" | awk 'NR>2 {sub(/^[A-Z]+\+/, "", $1); print $1}' | sort -u | tr '\n' ' ')" >&2
pdftotext -bbox "$OUT/continuous.pdf" "$OUT/continuous.bbox.html"
python3 - "$OUT/continuous.bbox.html" <<'PY'
import json, re, sys
html = open(sys.argv[1]).read()
pages = []
for m in re.finditer(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', html, re.S):
    w, h, body = float(m.group(1)), float(m.group(2)), m.group(3)
    lines = []
    for x0, y0, y1, t in re.findall(
        r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)</word>', body
    ):
        x0, y0, y1 = float(x0), float(y0), float(y1)
        # A line = consecutive words on one baseline; keep its first word.
        if lines and abs(lines[-1]["_y1"] - y1) < 0.5 and x0 > lines[-1]["x"]:
            continue
        lines.append({"label": t, "x": round(x0, 2), "top": round(y0, 2), "_y1": y1})
    for l in lines:
        del l["_y1"]
    page = {"page": len(pages) + 1, "size_pt": [round(w, 2), round(h, 2)], "lines": lines}
    if not lines:
        page["blank"] = True
    pages.append(page)
print(json.dumps({"page_count": len(pages), "pages": pages}, indent=1))
PY
