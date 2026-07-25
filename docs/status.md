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
- **Cell-internal styling — CLOSED (core door, protocol 55).** `applyStyle` used
  to carry no cell qualifier, so cell text could only pour at the default
  formatting. `Mutation::ApplyStyle` (and `Operation::ApplyStyle`) now take an
  optional `cell` address, and the appliers walk that cell's own paragraph stream
  instead of the story body; `buildTableCells` emits cell-qualified `applyStyle`
  ops for each cell run/paragraph, offsets cell-local. A host below v55 rejects
  the op and cell text keeps its default formatting (honest degrade — never wrong
  styling).
- **Honest limitation:** the exact story offset past a table (`TABLE_FOOTPRINT`)
  is still a conservative constant refined during editor integration.

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

## Tier-3 — hyperlinks (now native clickable, with a core door)

- **`w:hyperlink` → native clickable link.** A hyperlink's `r:id` is resolved
  through the document rels to its external URL (or `#anchor` for an internal
  bookmark) and carried on the run. Two things happen at lowering: (1) the
  conventional *look* — blue (`0000FF`) + underline — rides on a synthesized
  character style (unless the run sets those directly); (2) the external URL is
  carried on `LoweredRun.hyperlink_url`, and `doc-host-model` emits an
  **`insertHyperlink`** over the run's story range so the span is genuinely
  clickable. Verified end-to-end through the real wasm artifact (`hyperlink_docx`:
  the linked run carries the URL + the diagnostic reports it became a native link).
- **This required a core door (opened): `InsertHyperlink`.** The IDML hyperlink
  model — a `HyperlinkTextSource` run tag + a designmap `Hyperlink{source,
  destination}` + a `HyperlinkURLDestination` — parsed and rendered
  (`links.rs`), but had no *create* mutation. Added `InsertHyperlink` /
  `RemoveHyperlink` (paged-mutate + wire + protocol **52 → 53**): it splits runs
  over `[start, end)`, tags the middle `hyperlink_source`, and registers the two
  designmap resources; the engine mints the three cross-referencing ids so the
  wire stays `{storyId, start, end, url}`. (Cross-repo, like the anchored-frame
  door: `plugin-doc` stays isolation-clean, `core` carries the door.)
- **Both of Word's hyperlink forms.** Word writes links two ways: the
  `w:hyperlink` element (above) AND a `HYPERLINK` *field* — either the complex
  `fldChar begin / instrText / separate / result / end` run sequence (instruction
  may be split across several `w:instrText` runs) or the one-shot `w:fldSimple`.
  docx-import runs a field state machine over the paragraph's runs: control runs
  (`fldChar`/`instrText`) carry no display text and are dropped; result runs of a
  `HYPERLINK` field inherit the URL. Both forms then flow through the SAME
  `hyperlink_url` → `insertHyperlink` path — **no docx-lower / host-model / core
  change, purely a parse addition.** Verified through the real wasm (`fieldlink`
  fixture: two field links, one with a split instruction, both native-clickable).
- **Honest limitations:** internal `#anchor` / `HYPERLINK \l` bookmark targets
  keep the blue+underline look but are not yet native links (the core door
  registers URL destinations, not text anchors, and the renderer resolves a
  text-destination best-effort to the story's first page — useless for a
  multi-page story's mid-doc jumps; both are labelled, not faked). The diagnostic
  reports how many became clickable vs styled-only.

## Offset convention (correctness note)

Story character offsets in `doc-host-model` are **contiguous** across paragraphs:
the engine consumes the paragraph-break `\n` on `insertText` split — it is not a
stored character — so `applyStyle` / `insertAnchoredFrame` / `insertHyperlink`
ranges must NOT count it. (Fixed a latent off-by-one where the pour advanced the
offset past the separator; it was invisible in IR/wasm tests and would only have
bitten a live multi-paragraph apply.) NB: this matches core's range-styling apply
ops (`apply::character`/`paragraph`/`hyperlink`, all char-contiguous); it is
deliberately NOT the `mutate::locate`/`InsertText` byte+`\n` address space — those
two conventions diverge inside core today (a pre-existing split worth an RFI note).
paged.doc issues a single `insertText` at the base offset then contiguous-char
range ops, so it sits entirely on the contiguous side.

## M2 — edited save-back, Increment 1 (the targeted-patch spine)

- **The engine exists and is proven.** `docx-js` `save_edited(EditSet)` writes run
  edits back into the retained `.docx` as a **byte-level targeted patch**: only the
  changed `<w:t>`/`<w:rPr>` subtrees are rewritten; every other part AND every
  untouched subtree of `word/document.xml` stays **byte-identical**. New crate
  `docx-export` (`splice.rs`) uses quick-xml purely as a *locator* — it computes
  the source byte ranges to replace, raw-copies everything outside them, and splices
  hand-built fragments into the holes (byte-identity by construction). It never
  calls the ooxmlsdk serializer (`build-wasm.sh` guards this — linking `write_to`
  would blow the budget; wasm is 8.12 MiB / 8 MiB budget).
- **Provenance.** `docx-import` stamps each run's source `<w:r>` ordinal
  (`Run.source`) and each paragraph's body ordinal (`Paragraph.source_para_ord`);
  `docx-export::build_bindings` projects them onto lowered story `(block, run)`
  coordinates (replaying lowering's empty-run filter — a conformance test guards
  against drift). `DocSession` now retains the `OpcPackage` + bindings.
- **Style round-trip.** A synthesized `docx-auto-cN` style projects back to a
  **direct `<w:rPr>`** on save (so Word gets no synthetic-style clutter);
  `rpr.rs` renders `RunProps → <w:rPr>` as the exact inverse of the import parse.
- **The differ (`diff.rs`).** `diff(base, edited) → EditSet` — the seam the live
  readback door will target (it will re-lower the edited native model to a
  `LoweredDoc` and diff it against the import baseline). Runs are compared on their
  RESOLVED effective formatting, not lowered style-id strings (synth `docx-auto-cN`
  ids are positional and renumber between two lowerings); it inverts a synth style's
  props back to `RunProps` and recovers a real `<w:rStyle>` via an import-built
  `token → styleId` map. Structure-preserving edits only (count mismatch → skipped).
- **Verified end-to-end** (`docx-conformance/tests/save_back.rs`): import the memo
  fixture, change one run's text + toggle bold off another, save, and assert the
  targets changed, every other part + untouched subtree byte-identical, and the
  edits survive a re-import round-trip. Plus `splice`/`rpr` unit tests (exact
  byte-identity) and the drift guard.
- **Covered (Increment 1):** run text change + run property change on direct
  `<w:r>` children.
  **Deferred (labelled):** table structure;
  editing hyperlink/field runs and table-cell content (non-patchable bindings);
  and — the one platform seam —
  the LIVE editor wiring (below).

## DOC-03 — structured read + the LIVE save-back path (built; awaits the wasm publish)

The live path now exists end-to-end in code. **Core** (branch) ships a structured
read — `RequestStoryContent` → `StoryContent` (protocol 54): a story's paragraphs →
runs → text + applied styles + direct overrides, `CanvasModel::story_content`.
**paged.doc** consumes it: `docx-export::overlay_story_content` overlays the
read-back onto the import baseline lowering (a run's `characterStyle` IS the
plugin's own `char_style_id` token, since paged.doc created every style via
`applyStyle`), then `diff → EditSet → save_edited`.
`DocSession::save_edited_from_content(json)` + the wasm shim expose the one-call
live path. **Verified end-to-end with a MOCK read** (`doc03_read_overlay_diff_save_
round_trips`): read → overlay → diff → save → re-import reflects the edits; an
identity read is a verbatim no-op.

The one remaining seam is a RELEASE step, not code: core publishes
`@paged-media/canvas-wasm` v54, plugin-sdk `sync-wire.mjs` pulls the new wire types
and exposes `host.document.storyContent()` + `supports("document.readStory@1")`, and
the bundle calls it → `save_edited_from_content`. Until then the SDK method is a
reserved seam and the bundle degrades honestly (as it does for `openNative`).
Structure-preserving edits only; table-cell content not yet round-tripped.

**That release step is pre-verified (2026-07-25).** canvas-wasm was built locally
from the core DOC-03 branch (`sync-wasm.sh`'s path, the documented tool for
unpublished engine changes) and the generated `.d.ts` was checked to carry
`requestStoryContent`/`storyContentResult` with `selfId`/`characterStyle`/
`fontStyle`/`pointSize` — **matching both the hand-written plugin-api contract and
paged.doc's serde twin** (the real cross-layer drift risk, since three layers are
hand-maintained). The vendored `wire.d.ts` was then temporarily synced from that
build and the SDK adapter's reserved throw swapped for the real
`requestStoryContent` forward: **both packages typecheck**. Both changes were
reverted (a wire sync from an unpublished build must never be committed — the gate
compares against the published package), so the shipped state stays honest. Net:
when v54 publishes, the adapter is a proven ~6-line swap.

**The bundle is wired for it.** The exporter now attempts the live path: probe
`supports("document.readStory@1")` → `host.document.storyContent(storyId)` →
`engine.saveEditedFromContent(content)` (the storyId is captured at placement and
carried on `LastDoc`). It degrades honestly at every step — no read door, no story
id, or a failing save all fall back to the verbatim source rather than exporting
something wrong. (The `storyContent` call goes through a narrow cast until the
plugin-api canary carrying it publishes — the same pattern the v52/v53 mutation
ops use.)

## Deferred (labelled, never faked)

- **Edited save-back — LIVE editor run**: the full chain is built and wired; the
  engine side is verified with a mock read and the SDK adapter swap is
  pre-verified against a local v54 build. The in-browser round-trip lands with the
  canvas-wasm v54 publish + SDK sync (a release step, not code).
- **Standalone true-open** — degrades to embedded placement + a diagnostic when
  `document.openNative@1` is unwired (the common case today); the `docx →
  native-bytes` producer is a future `plugin-publish` sibling.
- **Constructs still past the tier line** — headers/footers, tracked changes,
  non-HYPERLINK fields (PAGE/DATE/REF), floating (anchored-with-wrap) images,
  page/section breaks and internal `#anchor` links (both need a renderer door) →
  surfaced as honest ADR-007 diagnostics on open, not silently dropped.
  (Numbering/lists, tables, inline images, hyperlinks and footnotes have SHIPPED
  — see the tier sections above.)
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
- **DOC-03 — structured whole-document read (NEW ask, M2 blocker).** `readModel`
  returns the core-owned `.pgm` model bytes, which an isolation-clean plugin cannot
  parse or diff. Edited save-back needs the edited stories' **text + styles** back
  in a structured form (JSON, or re-lowered to the plugin's `LoweredDoc`), so the
  plugin can diff against its import baseline and produce an `EditSet`. Until then
  the save-back engine (proven) runs test-driven only.
- **DOC-04** (Word-reference fidelity harness), **DOC-05** (ratify embed-or-open as
  a platform pattern), **DOC-06** (C-1 per-glyph faithful text) remain open.

## M2 Increment 2 — structural edits (insert / delete)

Save-back is no longer limited to in-place edits. `EditSet` gains a `structural`
lane — `DeleteRun`, `InsertRun` (after a run, or at the paragraph start),
`DeleteParagraph`, `InsertParagraph` — and the splicer grew the matching
byte-level moves: dropping a `<w:r>`/`<w:p>` subtree (skip its byte range) and
emitting a rendered fragment at an anchor. All coordinates address the BASELINE,
so ops never shift each other's addresses; one pass over the unmodified source
applies them all, and untouched bytes stay verbatim as before.

`diff` derives them automatically: when a paragraph's run count changes it runs an
**LCS alignment** over the runs' `(text, style)` identity — matched runs are left
alone, unmatched baseline runs become `DeleteRun`, unmatched edited runs become
`InsertRun` anchored after the preceding match. Paragraph count changes map to
`DeleteParagraph` / `InsertParagraph`. The DOC-03 overlay now replaces a
paragraph's run list wholesale when it changed, so the live editor path drives
structural edits too.

Verified: splice unit tests (delete run, insert-after, prepend, paragraph
delete+append) and two end-to-end conformance tests — one through the DOC-03
overlay (delete a run + append a run, other parts byte-identical), one via a
hand-authored `EditSet` (delete the heading, append a paragraph).

**Still deferred:** table STRUCTURE (rows/cells), and a paragraph↔table swap.

## Table-cell round-trip

Cell content is no longer non-patchable. `docx-import` stamps each cell paragraph
with a **`CellPath`** (`w:tbl` / `w:tr` / `w:tc` / `w:p` ordinals);
`docx-export::build_bindings` projects those onto lowered cell coordinates,
replaying `lower_table`'s emission order exactly (a vMerge-continue cell is
absorbed, not emitted) so cell indices line up with `LoweredTable.cells`. The
splicer gained a second locator path — `tbl → tr → tc → p → r` counters that run
alongside the body-paragraph counters (each ignores the other's runs) — and
`EditSet.cells` carries `CellRunEdit { block, cell, para, run, … }`. `diff` emits
them automatically when a table's cell text or style changes.

Verified: a direct cell text edit and a diff-driven one both round-trip, with the
other cells, the row/column spans, the grid widths, and every other part
unchanged.

**Honest limitations:** cell STRUCTURE (adding/removing rows or cells) is not
patched, and the locator assumes tables are not NESTED — a `<w:tbl>` inside a
`<w:tc>` would mis-count, so nested-table cells are never addressed by the
bindings.

## Read-path fidelity — footnotes / endnotes

`w:footnoteReference` / `w:endnoteReference` runs were silently dropped (they fell
into the importer's catch-all with no diagnostic). Now they are modelled:
`docx-import` resolves the `footnotes`/`endnotes` parts through the document rels,
parses each note's body paragraphs into `docx-core::Note` (skipping Word's
`separator` / `continuationSeparator` pseudo-notes), and stamps the referencing
run with `note_ref` (the `w:id`).

Lowering emits an **honest diagnostic** naming the footnote/endnote counts and the
number of in-text references, and states plainly that the note text is preserved
in the source `.docx` (so it round-trips on save) but is **not placed on the page**
— the native model has no footnote construct and there is no "insert footnote"
mutation. Note text is deliberately NOT inlined into the flow; a conformance test
asserts it never appears there, so the gap stays visible rather than faked.

Verified through the real wasm artifact (`footnote` fixture).

## M2 Increment 3 — paragraph properties (`<w:pPr>`)

Increment 1 noted `<w:pPr>` edits as a fast-follow and Increment 2 did not land
them, so a paragraph's own formatting never round-tripped — and the differ did not
even LOOK at paragraph styles (it compared runs only, so repointing a paragraph at
a different style produced an empty EditSet).

- `EditSet.paragraphs` (`ParaEdit { block, new_props, pstyle }`), the paragraph
  twin of the run lane: a synthesized paragraph style is projected into direct
  `<w:pPr>` formatting, and a REAL Word style id is recovered through a new
  `para_token_to_style_id` map (the lowered token is lossy, exactly as for runs).
- `rpr.rs::render_ppr` emits `<w:pPr>` in WML `CT_PPr` child order, packing the
  indents into one `<w:ind>` and before/after into one `<w:spacing>`, and emitting
  a hanging indent as `w:hanging` rather than a negative `firstLine`.
- The splicer replaces a targeted paragraph's `<w:pPr>`, or INSERTS one before the
  first run when the paragraph has none.
- `diff` now compares each paragraph's effective style + direct formatting and
  emits a `ParaEdit`, so the live DOC-03 path carries paragraph changes too.

Verified: a paragraph gains a centred `<w:pPr>` where it had none, a heading's
`pStyle` is dropped, run text is untouched, every other part stays byte-identical,
and the differ detects a paragraph-style repoint end-to-end.

## M2 Increment 3b — hyperlink / field runs are editable

Runs inside a link or field were marked non-patchable outright, on the assumption
that editing them would desync the target. That was too broad — in every form the
target lives somewhere OTHER than the run being edited:

- `<w:hyperlink>` — the `r:id` is on the WRAPPER, so the run's `<w:t>`/`<w:rPr>`
  are free to change. `RunSource::Hyperlink` now carries `{link_ord, run_ord}` and
  the splicer counts `<w:hyperlink>` children + the runs inside the open one,
  exactly as the table-cell path does.
- `w:fldSimple` — the instruction is a wrapper ATTRIBUTE; same treatment
  (`RunSource::Field { field_ord, run_ord }`).
- the complex `fldChar` **result** run — a plain direct `<w:r>` child whose URL
  lives in a separate `instrText` run. It was already addressable; only an
  over-broad `hyperlink.is_some()` check was rejecting it.

Verified: editing the display text of all three forms keeps the link target
(`https://paged.media/`, `…/complex`, `…/simple`) and leaves siblings untouched.

**Multi-`<w:t>` runs were never actually a limitation** — a run whose text is split
across several `<w:t>` children (Word does this after edits) already collapses
correctly into one on a text replacement, because `LoweredRun.text` is the
concatenation. A unit test now pins that, and the stale "deferred" claim is gone.

Still deferred: nested tables, a wrapped run INSIDE a table cell (it needs both
locator paths at once), and paragraph↔table swap.

## M2 Increment 3c — table row structure

Rows can now be added and removed, not just re-texted. `StructuralEdit::DeleteRow`
/ `InsertRow` address `(block, row)`; bindings carry the table's `<w:tbl>` ordinal
so a row resolves to `(table_ord, row)`, and the splicer drops a `<w:tr>` subtree
or emits a rendered one after a given row's `</w:tr>` — the same byte-splice
machinery as paragraphs, so untouched bytes stay verbatim. `render_table_row`
emits one `<w:tc>` per cell, each with the `<w:p>` a cell is required to contain.

`diff` derives them: a table whose row count shrank yields `DeleteRow` for the
trailing rows, one that grew yields `InsertRow` carrying the new rows' cell text
(superseded — rows are now identity-aligned too, see Increment 3e).

Verified on the table fixture — which has both a `gridSpan` header row and a
`vMerge` pair — that deleting the merge-continue row and inserting a fresh 2-cell
row leaves the spanning row, the surrounding body paragraphs and every other part
untouched; plus a diff-derived deletion end-to-end.

**Cell-level** add/remove within a row is still deferred: it changes the grid and
interacts with `gridSpan`/`vMerge`, so it needs its own design pass.

## M2 Increment 3d — block alignment (a preservation fix)

The structural diff paired story blocks BY INDEX, which quietly violated the
preservation invariant on a very ordinary edit. Deleting a paragraph in the
MIDDLE of a document produced the right text but kept the wrong `<w:p>`: with
`[Plain, Heading, Mix]` → `[Plain, Mix]`, it deleted the trailing paragraph and
rewrote the *heading's* node with Mix's runs — so the surviving paragraph still
carried `<w:pStyle w:val="Heading1"/>`, and the third paragraph's own node (its
`<w:pPr>`, rsids and any unmodelled children) was destroyed.

Blocks are now aligned by an **LCS over a block identity key** (a paragraph's
style + full text; a table's shape) rather than by position, so a deletion deletes
that block and the survivors keep their own nodes. One post-pass matters: an
adjacent Del+Ins is coalesced into a MATCH — without it, *editing* a paragraph's
text changes its key and the pair would be reported as delete-then-insert,
destroying the very `<w:p>` we want to patch in place. Coalescing keeps an edit an
edit and leaves a true deletion a deletion.

Regression-tested by the case that exposed it (the survivor must NOT inherit the
deleted heading's style), and every pre-existing save-back test still passes.

## M2 Increment 3e — row alignment (the same preservation fix, one level down)

Rows had the defect the block alignment had just fixed: the row diff deleted
TRAILING rows and let the cell comparison rewrite the rest by index. Deleting a
MIDDLE row therefore produced the right text while destroying the wrong `<w:tr>`.

Rows are now aligned by an LCS over a row identity key (its cells' text), reusing
the same `lcs_align` (including the Del+Ins→Match coalescing that keeps an edit an
edit). Cell content is compared only within MATCHED row pairs, so an edit can
never land on a row that is about to be deleted.

The regression test is the interesting part: the fixture gives each row a distinct
**unmodelled** `w:trHeight`. paged never reads it, so it is the only thing that can
prove WHICH node survived — asserting on cell text alone passes even when the
wrong row is kept. The test now asserts row 2's marker (`w:val="102"`) is intact
and the middle row's (`101`) is the one that went.
