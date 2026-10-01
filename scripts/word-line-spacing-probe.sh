#!/usr/bin/env bash
# Ask Word how tall a line is under each w:spacing/@w:lineRule (ADR 029).
# Builds docx_conformance::line_spacing_docx(), has Word save it as PDF
# (through word-pagination-probe.sh, which owns the Office automation
# traps), and measures the FIRST page of every section: how many lines Word
# put on it and the baseline pitch between them (pdftotext -bbox).
#
#   bash scripts/word-line-spacing-probe.sh <out-dir>
#
# Prints one JSON object per case; fixtures/line-spacing.word.json is that
# output plus provenance. Check the "fonts" line: Word silently substitutes
# a face it cannot load (Inter came back as Calibri), and a substituted face
# measures the wrong font.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/line-spacing.docx" line-spacing
bash scripts/word-pagination-probe.sh "$OUT/line-spacing.docx" "$OUT/line-spacing.pdf"
echo "fonts: $(pdffonts "$OUT/line-spacing.pdf" | awk 'NR>2 {sub(/^[A-Z]+\+/, "", $1); print $1}' | sort -u | tr '\n' ' ')"
pdftotext -bbox "$OUT/line-spacing.pdf" "$OUT/line-spacing.bbox.html"
python3 - "$OUT/line-spacing.bbox.html" <<'PY'
import json, re, sys
pages = re.split(r"<page ", open(sys.argv[1]).read())[1:]
seen = set()
for page in pages:
    words = re.findall(r'yMax="([\d.]+)">([^<]*)</word>', page)
    lines = []  # (baseline-ish y, first word, second word)
    for i, (y, w) in enumerate(words):
        y = float(y)
        if not lines or abs(lines[-1][0] - y) > 0.5:
            lines.append((y, w, words[i + 1][1] if i + 1 < len(words) else ""))
    if not lines:
        continue
    label = lines[0][1]
    if label in seen or lines[0][2] != "P001":
        continue  # only each section's first (full) page measures the pitch
    seen.add(label)
    ys = [l[0] for l in lines]
    steps = [round(b - a, 3) for a, b in zip(ys, ys[1:])]
    print(json.dumps({
        "case": label,
        "lines": len(lines),
        "last": f"{label} {lines[-1][2]}",
        "pitch_pt": round((ys[-1] - ys[0]) / (len(ys) - 1), 3),
        "step_min": min(steps),
        "step_max": max(steps),
    }))
PY
