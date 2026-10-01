#!/usr/bin/env bash
# Ask Word which header and footer each page shows (thoughts ADR 033, RFI
# DOC-05). Writes headers_docx() and a copy without w:evenAndOddHeaders,
# has Word save BOTH as PDF (through word-pagination-probe.sh, which owns
# the Office automation traps) and prints, per page, the text Word draws in
# the header band (above the 54 pt top margin), the body, and the footer
# band (below the bottom margin) -- the JSON docx-conformance/fixtures/
# headers.word.json records.
#
#   bash scripts/word-headers-probe.sh <out-dir>
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/headers.docx" headers
python3 - "$OUT" <<'PY'
import sys, zipfile
out = sys.argv[1]
src = zipfile.ZipFile(f"{out}/headers.docx")
dst = zipfile.ZipFile(f"{out}/headers-no-even.docx", "w", zipfile.ZIP_DEFLATED)
for i in src.infolist():
    d = src.read(i.filename)
    if i.filename == "word/settings.xml":
        d = d.replace(b"<w:evenAndOddHeaders/>", b"")
    dst.writestr(i.filename, d)
dst.close()
PY
for f in headers headers-no-even; do
  bash scripts/word-pagination-probe.sh "$OUT/$f.docx" "$OUT/$f.pdf" >&2
  pdftotext -bbox "$OUT/$f.pdf" "$OUT/$f.bbox.html"
done
python3 - "$OUT" <<'PY'
import json, re, sys
out = sys.argv[1]
def pages(path):
    html = open(path).read()
    res = []
    for page in re.findall(r'<page[^>]*>(.*?)</page>', html, re.S):
        bands = {"header": [], "body": [], "footer": []}
        for x0, y0, t in re.findall(
            r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">([^<]*)</word>', page
        ):
            y = float(y0)
            band = "header" if y < 54 else "footer" if y > 192 - 54 else "body"
            bands[band].append((y, float(x0), t))
        res.append({k: " ".join(t for _, _, t in sorted(v)) for k, v in bands.items()})
    return res
print(json.dumps({f: pages(f"{out}/{f}.bbox.html") for f in ["headers", "headers-no-even"]}, indent=1))
PY
