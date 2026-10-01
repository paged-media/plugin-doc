#!/usr/bin/env bash
# Ask Word how it numbers footnotes (thoughts ADR 034, RFI DOC-06). Writes
# footnote_numbering_docx() and two copies -- one with S1's numStart 10
# instead of 3, one with every section's w:footnotePr removed (only
# settings.xml's lowerRoman left) -- has Word save each as PDF (through
# word-pagination-probe.sh, which owns the Office automation traps) and
# prints, per note, the mark Word prints in front of its text: the JSON
# docx-conformance/fixtures/footnote-numbering.word.json records.
#
#   bash scripts/word-footnotes-probe.sh <out-dir>
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/fn.docx" footnote-numbering
python3 - "$OUT" <<'PY'
import re, sys, zipfile
out = sys.argv[1]
def variant(name, fn):
    src = zipfile.ZipFile(f"{out}/fn.docx")
    dst = zipfile.ZipFile(f"{out}/{name}.docx", "w", zipfile.ZIP_DEFLATED)
    for i in src.infolist():
        d = src.read(i.filename)
        if i.filename == "word/document.xml":
            d = fn(d)
        dst.writestr(i.filename, d)
    dst.close()
variant("fn-start10", lambda d: d.replace(b'<w:numStart w:val="3"/>', b'<w:numStart w:val="10"/>'))
variant("fn-no-section-props", lambda d: re.sub(rb"<w:footnotePr>.*?</w:footnotePr>", b"", d))
PY
for f in fn fn-start10 fn-no-section-props; do
  bash scripts/word-pagination-probe.sh "$OUT/$f.docx" "$OUT/$f.pdf" >&2
done
python3 - "$OUT" <<'PY'
import json, re, subprocess, sys
out = sys.argv[1]
res = {}
for f in ["fn", "fn-start10", "fn-no-section-props"]:
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", f"{out}/{f}.pdf"], capture_output=True, text=True).stdout).group(1))
    marks = []
    for p in range(1, n + 1):
        lines = [l.strip() for l in subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", f"{out}/{f}.pdf", "-"], capture_output=True, text=True).stdout.splitlines() if l.strip()]
        # The note area: a mark line, then "note <name>".
        for a, b in zip(lines, lines[1:]):
            m = re.fullmatch(r"note (\S+)", b)
            if m:
                marks.append({"page": p, "note": m.group(1), "mark": a})
    res[f] = marks
print(json.dumps(res, indent=1))
PY
