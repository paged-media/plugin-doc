# paged.doc — status

## This pass: M0 (foundation) + M1 Tier-0 read path

**DONE + verified (all green):**

- **DOC-02 `ooxmlsdk` → wasm32 spike — GREEN.** `ooxmlsdk` 0.12 (+ `zip`,
  `quick-xml`, `bytes`, `flate2`) compiles to `wasm32-unknown-unknown`; the
  foundation is safe to build on the typed DOM.
- **`paged-ooxml`** — OPC container over `zip` + `quick-xml`, `[Content_Types].xml`
  + `_rels` graph, verbatim carry-through preservation, `ooxmlsdk` typed-DOM bridge.
- **Tier-0 read path** — `docx-import` (WordprocessingML body + styles + section →
  `docx-core`) → `docx-lower` (→ `LoweredDoc` IR with synthesized direct-format
  styles + minted swatches) → `docx-js` (one wasm module).
- **`@paged-media/doc-host-model`** — pure IR → `Mutation[]` (styles, swatches,
  offset-tracked pour), unit-tested.
- **`@paged-media/doc` bundle** — manifest (`media.paged.doc`, CLI-valid),
  `activate` registering importer/exporter/objectType/editContext/command/panel,
  wasm boot, embedded placement wiring.
- **Conformance** — a real in-memory `.docx` imports + lowers correctly, and the
  OPC package round-trips **byte-identical** including an unknown part; malformed
  input errors without panicking.

**Verification gates, all green:** `cargo build/nextest` (11 tests),
`cargo clippy -D warnings`, `cargo fmt --check`, `cargo deny check`, the wasm
build (7.5 MiB, under the 8 MiB budget — unoptimized; CI runs `wasm-opt -Oz`),
`pnpm typecheck`, `pnpm test` (host-model), the contract-import lint, and
`paged-plugin validate`.

## Tier-1a (added after the initial pass)

- **`docDefaults`** — the document-wide base run/paragraph properties become a
  `docx-Default` base paragraph style that every un-based style + un-styled
  paragraph inherits (closes the "font/size falls back to engine default" gap).
- **Tab stops** — `w:tabs` → `paragraphTabStops` (position pt + alignment; `clear`
  stops dropped; leader glyphs are a later refinement).
- **keep-with-next / keep-lines-together** — `w:keepNext`/`w:keepLines` lowered.
- **Underline correctness** — `w:u w:val="none"` now lowers to `underline:false`
  rather than being read as "on".

All Tier-1a lowering is verified through the real wasm artifact (the `tier1_docx`
conformance fixture).

## Deferred (labelled, never faked)

- **Edited save-back** (native → WordprocessingML projection, targeted patch) —
  **M2**. The exporter re-emits the retained package verbatim (zero-edit only).
- **Standalone true-open** — degrades to embedded placement + a diagnostic when
  `document.openNative@1` is unwired (the common case today); the `docx →
  native-bytes` producer is a future `plugin-publish` sibling.
- **Tier 1+ constructs** — numbering/lists, tables, inline + floating images,
  headers/footers, fields, notes, tracked changes → surfaced as honest ADR-007
  diagnostics on open, not silently dropped.
- **In-editor live render / real `host.nativeDocument`** — the bundle is written to
  the shipped contract and exercised to the seam; full end-to-end is
  host-integration-verified in a later milestone (needs the editor checkout + a
  wired `NativeDocumentBackend`).

## Platform doors (RFI)

- **DOC-01** `host.nativeDocument.open(bytes)` and **DOC-03**
  `readModel/readComposition/listParts` now **exist** in the SDK contract
  (conditional — live only when the editor injects a `NativeDocumentBackend`). The
  base-idea listed them as missing; they landed after it was drafted. The importer
  declares `openNative`/`readNative` and probes `host.supports(...)`.
- **DOC-04** (Word-reference fidelity harness), **DOC-05** (ratify embed-or-open as
  a platform pattern), **DOC-06** (C-1 per-glyph faithful text) remain open.
