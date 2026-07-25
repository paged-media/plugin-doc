#!/usr/bin/env bash
# Build the paged.doc engine wasm (docx-js) and land the wasm-bindgen
# `--target web` output in packages/doc-bundle/bin/ — the path the manifest
# declares under capabilities.wasm[] (governance + the 8 MiB plugin-cli size
# gate). The bundle loads it via the wasm-bindgen glue (the core/canvas-wasm
# pattern), NOT via loadBundleWasm.
#
# wasm-opt: CI pins binaryen (old apt binaryen breaks wasm-bindgen externref
# table grow — the "Table.grow failed" gotcha); locally it is applied when
# present, skipped with a warning when absent.
set -euo pipefail
cd "$(dirname "$0")/.."

OUT=packages/doc-bundle/bin
BUDGET=$((8 * 1024 * 1024))
mkdir -p "$OUT"

# Budget guard: the save-back patcher (docx-export) MUST stay a byte-level splice
# and never call the ooxmlsdk serializer — the first `write_to`/`to_xml`/
# `serialize_root` monomorphization links the whole WML write codegen and blows
# the 8 MiB budget. Fail fast on an actual call (comments/descriptions excluded).
if grep -rnE 'serialize_root\(|\.write_to\(|\.to_xml\(' docx-export/src/ >/dev/null; then
  echo "error: docx-export calls the ooxmlsdk serializer — save-back must byte-splice" >&2
  grep -rnE 'serialize_root\(|\.write_to\(|\.to_xml\(' docx-export/src/ >&2
  exit 1
fi

cargo build --release --target wasm32-unknown-unknown -p docx-js

# Pin check: wasm-bindgen-cli must match the Cargo.lock wasm-bindgen.
LOCKED=$(grep -A1 '^name = "wasm-bindgen"$' Cargo.lock | grep version | head -1 | cut -d'"' -f2)
CLI=$(wasm-bindgen --version | awk '{print $2}')
if [ "$LOCKED" != "$CLI" ]; then
  echo "error: wasm-bindgen-cli $CLI != Cargo.lock wasm-bindgen $LOCKED" >&2
  echo "       cargo install wasm-bindgen-cli --version $LOCKED" >&2
  exit 1
fi

wasm-bindgen target/wasm32-unknown-unknown/release/docx_js.wasm \
  --target web --out-dir "$OUT"

if command -v wasm-opt >/dev/null 2>&1; then
  wasm-opt -Oz "$OUT/docx_js_bg.wasm" -o "$OUT/docx_js_bg.wasm"
else
  echo "warning: wasm-opt not found — shipping unoptimized wasm (CI optimizes)" >&2
fi

SIZE=$(wc -c < "$OUT/docx_js_bg.wasm" | tr -d ' ')
echo "docx_js_bg.wasm: $SIZE bytes (budget $BUDGET)"
if [ "$SIZE" -gt "$BUDGET" ]; then
  echo "error: wasm artifact exceeds the 8 MiB plugin budget" >&2
  exit 1
fi
