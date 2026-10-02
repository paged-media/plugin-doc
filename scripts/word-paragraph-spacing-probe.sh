#!/usr/bin/env bash
# Ask Word how much room it leaves between a paragraph with space AFTER and
# one with space BEFORE, and whether space before survives at the top of a
# page (ADR 029). Builds docx_conformance::paragraph_spacing_docx() in each
# compatibility mode, has Word save it as PDF (through
# word-pagination-probe.sh, which owns the Office automation traps), and
# measures the baseline distance of every pair (pdftotext -bbox).
#
#   bash scripts/word-paragraph-spacing-probe.sh <out-dir>
#
# Prints one JSON object per mode; fixtures/paragraph-spacing.word.json is
# that output plus provenance.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"
cd "$ROOT"
for which in paragraph-spacing paragraph-spacing-11 paragraph-spacing-none; do
  cargo run -q -p docx-conformance --bin dump-fixture -- "$OUT/$which.docx" "$which"
  bash scripts/word-pagination-probe.sh "$OUT/$which.docx" "$OUT/$which.pdf" >/dev/null
  pdftotext -bbox "$OUT/$which.pdf" "$OUT/$which.bbox.html"
  python3 - "$OUT/$which.bbox.html" "$which" <<'PY'
import json, re, sys
pages = re.split(r"<page ", open(sys.argv[1]).read())[1:]
out = {"fixture": sys.argv[2], "pairs": {}, "top": {}}
for pi, page in enumerate(pages):
    first = {}
    for y, w in re.findall(r'yMax="([\d.]+)">([^<]*)</word>', page):
        first.setdefault(w, float(y))
    line = None
    for w, y in first.items():
        m = re.fullmatch(r"(S\d\d)a", w)
        if m and m.group(1) + "b" in first:
            # one line is the pitch of the "first line" -> S01a step
            out["pairs"][m.group(1)] = round(first[m.group(1) + "b"] - y, 2)
    ys = sorted(set(first.values()))
    for label in ("TOP", "TOP2"):
        if label in first:
            out["top"][label] = {"page": pi + 1, "baseline_y": round(first[label], 2),
                                 "next_step": round(first.get(label + "next", 0) - first[label], 2)}
print(json.dumps(out))
PY
done
