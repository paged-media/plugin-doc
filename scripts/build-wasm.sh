#!/usr/bin/env bash
# Build the paged.doc engine wasm (docx-js) and land the wasm-bindgen
# `--target web` output in packages/doc-bundle/bin/ — the path the manifest
# declares under capabilities.wasm[] (governance + the 100 MB plugin-cli size
# gate). The bundle loads it via the wasm-bindgen glue (the core/canvas-wasm
# pattern), NOT via loadBundleWasm.
#
# wasm-opt: CI pins binaryen (old apt binaryen breaks wasm-bindgen externref
# table grow — the "Table.grow failed" gotcha); locally it is applied when
# present, skipped with a warning when absent.
set -euo pipefail
cd "$(dirname "$0")/.."

OUT=packages/doc-bundle/bin
# The budget is now 100 MB for the WHOLE APP including every plugin
# (maintainer decision, 2026-08-19), enforced as a SUM by the editor's
# scripts/wasm-budget.mjs. This per-artifact stop keeps a runaway build
# from sailing through unnoticed here; the number that governs is the app
# total. Mirrors plugin-sdk WASM_BUDGETS — change them together.
BUDGET=$((100 * 1000 * 1000))
mkdir -p "$OUT"

# Budget guard: the save-back patcher (docx-export) MUST stay a byte-level splice
# and never call the ooxmlsdk serializer — the first `write_to`/`to_xml`/
# `serialize_root` monomorphization links the whole WML write codegen and blows
# the app wasm budget. Fail fast on an actual call (comments/descriptions excluded).
if grep -rnE 'serialize_root\(|\.write_to\(|\.to_xml\(' docx-export/src/ >/dev/null; then
  echo "error: docx-export calls the ooxmlsdk serializer — save-back must byte-splice" >&2
  grep -rnE 'serialize_root\(|\.write_to\(|\.to_xml\(' docx-export/src/ >&2
  exit 1
fi

# The `wasm-release` profile (Cargo.toml) keeps the artifact under the
# manifest's `maxBytes`; its output lands under wasm-release/, not release/.
cargo build --profile wasm-release --target wasm32-unknown-unknown -p docx-js

# Pin check: wasm-bindgen-cli must match the Cargo.lock wasm-bindgen.
LOCKED=$(grep -A1 '^name = "wasm-bindgen"$' Cargo.lock | grep version | head -1 | cut -d'"' -f2)
CLI=$(wasm-bindgen --version | awk '{print $2}')
if [ "$LOCKED" != "$CLI" ]; then
  echo "error: wasm-bindgen-cli $CLI != Cargo.lock wasm-bindgen $LOCKED" >&2
  echo "       cargo install wasm-bindgen-cli --version $LOCKED" >&2
  exit 1
fi

wasm-bindgen "${CARGO_TARGET_DIR:-target}/wasm32-unknown-unknown/wasm-release/docx_js.wasm" \
  --target web --out-dir "$OUT"

if command -v wasm-opt >/dev/null 2>&1; then
  wasm-opt -Oz "$OUT/docx_js_bg.wasm" -o "$OUT/docx_js_bg.wasm"
else
  echo "warning: wasm-opt not found — shipping unoptimized wasm (CI optimizes)" >&2
fi

SIZE=$(wc -c < "$OUT/docx_js_bg.wasm" | tr -d ' ')
echo "docx_js_bg.wasm: $SIZE bytes (budget $BUDGET)"
if [ "$SIZE" -gt "$BUDGET" ]; then
  echo "error: wasm artifact exceeds the 100 MB app wasm budget" >&2
  exit 1
fi

# The manifest declares a tighter limit for this module (`maxBytes`), and a
# host that honours it refuses a larger file. Hold the build to it, so the
# artifact cannot outgrow its own declaration again unnoticed.
MANIFEST=packages/doc-bundle/manifest.json
MAX=$(grep -o '"maxBytes": *[0-9]*' "$MANIFEST" | head -1 | grep -o '[0-9]*$')
if [ -z "$MAX" ]; then
  echo "error: no maxBytes in $MANIFEST" >&2
  exit 1
fi
echo "docx_js_bg.wasm: $SIZE bytes (manifest maxBytes $MAX)"
if [ "$SIZE" -gt "$MAX" ]; then
  echo "error: wasm artifact exceeds the manifest's maxBytes ($MAX)" >&2
  exit 1
fi
