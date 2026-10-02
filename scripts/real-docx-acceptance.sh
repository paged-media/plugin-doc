#!/usr/bin/env bash
# ADR 029 decision 6 against REAL Word documents (docs/acceptance-real-docx.md):
# Word's own page map (its PDF) versus the editor's (File▸Open through the
# real pipeline), per page the paragraph its first body line belongs to.
#
#   bash scripts/real-docx-acceptance.sh <editor-checkout> <out-dir> <a.docx> [b.docx …]
#
# <editor-checkout> is an editor worktree whose `@paged-media/doc` resolves to
# THIS plugin-doc (its `link:` override) with a built bin/ (build-wasm.sh).
# Needs Microsoft Word (macOS) and poppler. Up to three measurements per
# document (MODES, default all three): `none` (as a user opens it: the
# editor's default face), `word` (Word's own faces registered before the
# open) and `word-reload` (the same, then a save + reload: the workaround an
# engine before core 133f19b needed, see measure.spec.ts).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
EDITOR="$(cd "$1" && pwd)"; shift
OUT="$(mkdir -p "$1" && cd "$1" && pwd)"; shift
mkdir -p "$OUT/in" "$OUT/word" "$OUT/ours"
DOCS=()
for f in "$@"; do
  # Name by the WHOLE source name: corpus files repeat stems.
  name="$(basename "$f")"; name="${name//./_}.docx"
  cp "$f" "$OUT/in/$name"
  DOCS+=("$OUT/in/$name")
  bash "$ROOT/scripts/word-pagination-probe.sh" "$OUT/in/$name" "$OUT/word/${name%.docx}.pdf"
  python3 "$ROOT/scripts/real-docx/word_map.py" "$OUT/in/$name" "$OUT/word/${name%.docx}.pdf" \
    "$OUT/word/${name%.docx}.map.json" > /dev/null
done
SPEC="$EDITOR/apps/canvas/tests/e2e/real-docx-measure.spec.ts"
cp "$ROOT/scripts/real-docx/measure.spec.ts" "$SPEC"
trap 'rm -f "$SPEC"' EXIT
LIST="$(IFS=,; echo "${DOCS[*]}")"
MODES="${MODES:-none word word-reload}"
for mode in $MODES; do
  fonts=none; reload=0
  [ "$mode" = word ] && fonts=word
  [ "$mode" = word-reload ] && fonts=word && reload=1
  (cd "$EDITOR/apps/canvas" && REAL_DOCX="$LIST" REAL_OUT="$OUT/ours" REAL_FONTS=$fonts \
    REAL_RELOAD=$reload IDML_CANVAS_TEST_PORT="${PORT:-5291}" \
    npx playwright test tests/e2e/real-docx-measure.spec.ts --reporter=list --retries=0)
done
for d in "${DOCS[@]}"; do
  n="$(basename "$d" .docx)"
  for mode in $MODES; do
    echo "== $n ($mode)"
    python3 "$ROOT/scripts/real-docx/compare.py" "$OUT/word/$n.map.json" \
      "$OUT/ours/$n.ours.$mode.json" "$OUT/$n.cmp.$mode.json"
  done
done
