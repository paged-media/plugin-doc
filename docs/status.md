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

## Tier-1b — lists / numbering

- **Bullet + numbered lists** — `w:numPr` (numId/ilvl) is resolved through
  `numbering.xml` (numId → abstractNumId → level → numFmt/lvlText) at import, then
  lowered to native list paragraphs: `paragraphListType` (`BulletList`/
  `NumberedList` — the field the renderer gates marker emission + auto-numbering
  on), `paragraphBulletCharacter` (Wingdings/Symbol glyphs normalized to Unicode,
  e.g. F0B7 → •), `paragraphNumberingFormat` (Word `numFmt` → the IDML sample
  string the engine's `format_number` reads: decimal/upper-lower Roman/alpha), and
  a per-level left indent. Identical list paragraphs share one synthesized style.
- **No core change was required** — the `paragraphListType` PropertyPath already
  existed; the engine already renders bullets + auto-counts numbered lists. So
  lists lower fully plugin-side (isolation-clean) and render through the existing
  renderer. Verified end-to-end through the real wasm artifact (`list_docx`).
- Not yet: multi-level list indent metrics from the list definition (a fixed
  18 pt/level default is used), list continuation/restart nuances, and bullet
  leader/suffix text.

## Tier-2 — tables

- **`w:tbl` → native tables.** The grid (`w:tblGrid/w:gridCol` widths), rows, and
  cells are parsed; `gridSpan` (horizontal) and `vMerge` (vertical) are resolved
  into positioned cells with row/col spans (a `vMerge`-continue cell is absorbed
  into its restart cell above and not emitted). The story IR is now
  **block-structured** (`LoweredBlock` = paragraph | table) so tables interleave
  with body paragraphs in document order.
- **Lowering → mutations.** `doc-host-model` emits `insertTable` (grid + column
  widths), then per cell an `insertText` addressed by `TextCellAddr{tableId,row,
  col}` and a `setCellSpan` for merged cells — mirroring the proven `sheet`
  table lane. Because `insertTable`'s outcome mints the table id, the host-model
  exposes a **phased story plan** (`buildStory`): text steps + table steps
  (`{insert, cells(tableId)}`), which the bundle executes in order.
- Verified end-to-end through the real wasm artifact (`table_docx`): block order,
  grid, widths, gridSpan, and vMerge all lower correctly.
- **Honest limitations:** cell-internal paragraph/character *styling* is not
  applied (`applyStyle` carries no cell qualifier — cell text pours at the
  default); and the exact story offset past a table (`TABLE_FOOTPRINT`) is a
  conservative constant refined during editor integration.

## Tier-2 — inline images (with a core door)

- **`w:drawing` inline images → native anchored frames.** A drawing run is parsed
  (typed: `wp:inline`/`wp:anchor` → `a:graphic` → `graphicData` → `pic:pic` →
  `blipFill` → `blip@r:embed`), the `r:embed` rel resolved to its `word/media/…`
  bytes, and the intrinsic `wp:extent` (EMU) converted to points. `docx-lower`
  emits a per-paragraph image placement with a self-contained `data:<mime>;base64`
  URI; `doc-host-model` turns each into an **`insertAnchoredFrame`** at the
  paragraph's story offset. Verified end-to-end through the real wasm artifact
  (`image_docx`).
- **This required a core door (opened): `InsertAnchoredFrame`.** The anchored-image
  model, inline render path (`paged-renderer` `anchored.rs`), and property edits
  already existed, but there was no *mutation to create* an anchored frame —
  `InsertFrame` is page-positioned only. Added `InsertAnchoredFrame` /
  `RemoveAnchoredFrame` (paged-mutate + wire + protocol **51 → 52**): it pushes an
  image-bearing anchored Rectangle onto the paragraph at a story offset
  (`setting: None` ⇒ the renderer's default `InlinePosition`), so the frame draws
  inline via the existing path — no renderer change. (This is a cross-repo change;
  `plugin-doc` stays isolation-clean, `core` carries the door.)
- **Honest limitations:** large images embed as inline base64 (a part reference is
  the later refinement); anchored frames position at the paragraph level (not a
  precise intra-paragraph offset — the renderer's own current behavior).

## Tier-3 — hyperlinks

- **`w:hyperlink` → styled + preserved.** A hyperlink's `r:id` is resolved through
  the document rels to its external URL (or `#anchor` for an internal bookmark),
  carried on the run, and lowered with the conventional hyperlink look — blue
  (`0000FF`) + underline — via a synthesized character style (unless the run sets
  those directly). No core door + no TS change (standard character props).
  Verified end-to-end through the real wasm artifact (`hyperlink_docx`).
- **Honest limitation (diagnosed):** the run is styled and the target is preserved
  in the source `.docx`, but it is not yet a *native clickable* link — the native
  hyperlink model is multi-part (a source span + a designmap `Hyperlink` + a
  `HyperlinkDestination`) with no create-mutation, a larger core door than the
  inline-image one. An ADR-007 diagnostic surfaces this on every hyperlinked doc.

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
