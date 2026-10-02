# Architecture

How the `paged.doc` plugin is built: Microsoft Word documents (`.docx`, `.dotx`) as content
in a page-layout editor. It describes what the code does at commit `76e1d06`. The reason
behind each choice is in an ADR under [`adr/`](adr/README.md), linked where it applies.

## Crates and packages

The repo is a Cargo workspace and a pnpm workspace side by side. The eight crates sit at the
repo root, not under `crates/`; the two TypeScript packages are under `packages/`. The list
after the table gives the dependency direction.

| Crate or package | What it owns |
|---|---|
| `paged-ooxml` | The format-mechanical OOXML layer; nothing in it is Word-specific. `OpcPackage` reads a package into an ordered list of parts, stored decompressed, and writes it back with every untouched part's bytes unchanged. Beside it: the content-type table, the relationship graph, and `parse_root`, the bridge to the typed part tree of the third-party crate `ooxmlsdk`. |
| `docx-core` | Plain serde structs for what the plugin reads of a Word document: body blocks (paragraphs, runs, tables), the style catalogue, sections, notes, headers and footers, pictures. Lengths stay in Word's units. Depends on `serde` only. |
| `docx-import` | `.docx` bytes to `docx-core`. Finds the main document, styles, numbering, settings, footnote, endnote, header and footer parts through the relationship graph, parses them with `ooxmlsdk`, and records where each paragraph and run came from in the source XML. |
| `docx-lower` | A pure function from `docx-core` to `LoweredDoc` (`docx-lower/src/ir.rs`): swatches, a style catalogue ordered parents first, one story of blocks, the geometry of every section, diagnostics. `sections.rs` decides which Word sections share a native story; `line_height.rs` holds Word's line height per typeface. |
| `docx-export` | Save-back: the provenance map (`bindings.rs`), the overlay of a read-back story onto the lowering (`overlay.rs`), the differ (`diff.rs`) and the byte patcher (`splice.rs`, `rpr.rs`). |
| `docx-skeleton` | Writes a minimal IDML package for standalone open, with `zip` and string templates. |
| `docx-js` | The one `cdylib`. `DocSession` (`docx-js/src/core.rs`) is plain Rust and does the work; `DocEngine` (`docx-js/src/lib.rs`) is compiled only for `wasm32` and forwards to it. |
| `docx-conformance` | Test-only: builders that assemble `.docx` packages in memory, the recorded answers of Microsoft Word (`fixtures/*.word.json`), and the integration tests. |
| `packages/doc-host-model` | `@paged-media/doc-host-model`, private. Pure TypeScript: `LoweredDoc` to host mutations (`src/mutations.ts`), and a hand-written TypeScript copy of the IR types (`src/lowered.ts`). |
| `packages/doc-bundle` | `@paged-media/doc`. The manifest, `activate(host)`, the engine facade, the two entry points, the pour, the outline panel and the menu entry. The only package that touches the host. |

- `docx-import` depends on `paged-ooxml` and `docx-core`; `docx-lower` on `docx-core`;
  `docx-export` on `docx-core`, `docx-lower` and `paged-ooxml`; `docx-skeleton` on
  `docx-core` and `docx-lower`; `docx-js` on all six. `paged-ooxml` and `docx-import` are
  the only crates that name an `ooxmlsdk` type.
- No crate depends on an engine crate or on another plugin, and `Cargo.toml` has no git
  dependency. `deny.toml` denies unknown registries and git sources;
  `scripts/check-contract-imports.mjs` rejects any static import in the packages' non-test
  source files that is outside the contract packages `@paged-media/plugin-api` and
  `@paged-media/plugin-sdk`, this repo's packages and `react`
  ([ADR 315](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/315-isolation-contract.md)).
- `doc-bundle` takes the two contract packages and `react` as peer dependencies; tsup
  inlines `doc-host-model` into the built bundle. The split into Rust semantics, a
  forwarding wasm class and a translating TypeScript package is the shape of
  [ADR 314](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/314-plugin-shape.md).
- Third-party Rust: `ooxmlsdk` 0.12 (feature `mce`), `zip` 2, `quick-xml` 0.41, `wasm-bindgen`
  0.2. `ooxmlsdk` comes from crates.io: the README and several comments call it "vendored",
  but there is no `vendor/` directory and no patch ([ADR 601](adr/601-own-package-container.md)).

## From a `.docx` to native content

```
.docx bytes
   |  OpcPackage::read, parse_root                         paged-ooxml
   v
DocxDocument, retained package, provenance                 docx-import, docx-core
   |  lower_with(options)                                  docx-lower
   v
LoweredDoc as JSON  ......... wasm boundary .........      docx-js
   |  buildStyleMutations, buildStory / buildStoryBlocks   packages/doc-host-model
   v
host.document.mutate(...)                                  packages/doc-bundle/src/pour.ts
   v
swatches, styles, stories, tables, anchored frames, hyperlinks, laid out by the host engine
```

- **Reading.** `OpcPackage::read` returns a named error for a legacy binary `.doc` or an
  RTF file before the zip reader runs. `parse_root` refuses XML nested deeper than 256
  elements before `ooxmlsdk`'s recursive parser sees it. An unreadable container or an
  unparseable main document is an error; any other part that fails to parse is skipped.
- **Lowering** ([ADR 600](adr/600-docx-lowered-onto-native-model.md)). The plugin has no
  layout engine and paints nothing. Every id in the IR is a complete native token
  (`ParagraphStyle/docx-…`, `CharacterStyle/docx-…`, `Color/docx-RRGGBB`) and every property
  value serialises as the host's value union. Direct formatting becomes a synthesized named
  style (`…/docx-auto-pN`, `…/docx-auto-cN`); document defaults become
  `ParagraphStyle/docx-Default`. Word's layout rules are written as native property values:
  breaks as `paragraphStartParagraph`, keep rules as the `paragraphKeep…` properties,
  first-fit line breaking as `paragraphComposer`, line spacing as `characterLeading`.
- **Not placed.** Footnotes, endnotes, headers and footers are parsed but not placed; a
  floating picture is placed inline, without its position and wrap; a field other than a
  hyperlink is placed as its last computed text. `LoweredDoc.diagnostics` reports each.
- **The wasm boundary.** `DocEngine` has eleven methods besides its constructor. Byte arrays
  cross it (the `.docx` in; the skeleton and the saved `.docx` out) and JSON strings (the
  lowering, story ids, the read-back story, the skipped edits), with no version number.
  `packages/doc-bundle/src/engine.ts` imports the wasm-bindgen glue itself, not through the
  host's wasm loader, and boots the engine for one import or one export, then frees it.

## Two entry points

Both run `ingest` in `packages/doc-bundle/src/activate.ts` and share the lowering and the
pour ([ADR 603](adr/603-two-entry-points.md)).

**Place.** The command `media.paged.doc.command.placeDoc` ("Place Word document…"), the
panel button and the menu entry `Object/Insert Word document…` pick a file and call
`placeEmbedded` (`src/place.ts`). It inserts one text frame on the active page (else the
first), sized to the margin box of the Word document's first section, finds its story by a
hit test, applies the style catalogue and pours the whole story. Nothing here adds a frame
or a page.

**Open.** The importer `media.paged.doc.importer.docx` (`.docx`, `.dotx`) calls
`openStandalone` (`src/open.ts`) when the host reports `document.openNative@1`; otherwise it
places the document as above and logs that it did.

1. `docx-skeleton` writes an IDML package with one spread, page, text frame and empty story
   per native story, with the fixed ids `docx_sp{k}`, `docx_p{k}`, `docx_f{k}`, `docx_s{k}`.
   The page has the size and margins of the story's first section; the frame is the margin
   box with the story's columns, a `LeadingOffset` first baseline and zero insets. It writes
   no master spread. `host.nativeDocument.open` makes this package the open document.
2. The style catalogue is applied. Each story then gets `setFlowGrowRule` with `grow: true`
   and `copyFrameOptions: true`, because IDML cannot carry the rule. From then on the host
   engine adds pages while a story oversets; the plugin inserts no page.
3. `sectionBlocks` cuts the block list at each story's first section, and each group is
   poured into its story.

A Word section that starts a page is a new story. A `continuous` or `nextColumn` section
joins the story before it when the native model can say what Word does there: other left
and right margins become paragraph indents, a column change becomes span or split columns
(`docx-lower/src/sections.rs`). The skeleton and the lowering call the same
`place_sections_with`, so their story indices agree. The page-growth design this rests on
is the subject of ADR 029, a proposal that is not published yet.

## The pour

- `buildStoryBlocks` turns a block list into steps; `pourSteps` (`src/pour.ts`) runs them.
- Consecutive paragraphs are one text step: one `insertText` with the paragraphs joined by
  `\n`, then `applyStyle` ranges, an `insertAnchoredFrame` per inline picture (the image
  travels in the mutation as a `data:` URI) and an `insertHyperlink` per linked run, sent
  as one `batch`. Blank lines are styled in a last step, by zero-length `applyStyle` ranges.
- A table is `insertTable`, whose outcome carries the new table's id, then one batch for
  its cells: `insertText`, cell-addressed `applyStyle`, `setElementProperty` for insets and
  vertical alignment, `setCellSpan` for merged cells. This is why a story is not one batch.
- Two offsets run side by side: style, anchor and link ranges count characters, contiguous
  across paragraphs; `insertText` counts UTF-8 bytes including the separators.

## One build, several engine versions

Optional doors are probed with `host.supports(...)`
([ADR 305](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/305-doors-always-present.md)).
No flag reports whether an engine knows a single operation or property path, so the bundle
sends the batch and reads the outcome: a refused batch applies nothing, and the operations
are replayed one by one. A refused style operation becomes a warning under
`media.paged.doc/styles`, a refused grow rule one under `media.paged.doc/open`; a refused
pour operation is logged. On open, if the engine refuses a span or split column property
that the lowering uses, the bundle lowers again with a page break at each column change,
rebuilds the skeleton, reopens it and warns under `media.paged.doc/columns`.

## Save-back

The exporter `media.paged.doc.exporter.docx` runs `exportDocx` in `activate.ts`
([ADR 602](adr/602-save-back-is-a-byte-splice.md)).

- With no document imported in this session it returns nothing. Without
  `document.readStory@1`, or without a story id, it returns the original bytes.
- Otherwise it reads each story with `host.document.storyContent`, joins them into one body
  (`mergeSectionContents`), boots the engine, loads the original bytes again and calls
  `save_edited_from_content`: lower again as the baseline; overlay the read-back onto it;
  diff the two lowerings into an `EditSet` in lowered `(block, run)` coordinates; resolve
  those through `DocxBindings` to source ordinals; patch the main document part.
- `splice.rs` uses `quick-xml` only to find byte ranges in one streaming pass and copies
  every byte outside them. The patched part goes into a clone of the retained package, and
  `OpcPackage::write` re-emits all other parts unchanged. The ZIP is written anew, so the
  guarantee is per part, not per file. `scripts/build-wasm.sh` fails if `docx-export/src`
  calls the `ooxmlsdk` serializer.
- An edit the patcher cannot place is skipped and named; the bundle publishes the list
  under `media.paged.doc/save-back`. If the save throws, the original bytes are exported.
- The overlay starts from a copy of the baseline lowering and walks its blocks in order:
  each paragraph block takes the next read-back paragraphs by position (one, plus one for
  each break that split it) and receives only their run text and character style; table
  blocks are stepped over; the read-back paragraph style is not applied. It never adds or
  removes a block and does not read read-back paragraphs beyond those the baseline's blocks
  take (`docx-export/src/overlay.rs:96-140`). The differ and the patcher also handle
  paragraph insert and delete, paragraph properties, table-cell runs, rows and columns; the
  tests in `docx-conformance/tests/save_back.rs` reach those through lowerings or edit sets
  built in the test, and no test of the read-back path adds or removes a paragraph.

## Where data is stored

- **In the document:** ordinary native content, plus a metadata envelope on one frame,
  `{ v: 1, data: { part, blocks } }` (and `sections` after a standalone open). The object
  type `wordDocument` matches a frame whose `data.part` is a string.
- **As a container part:** the original file, written through `host.parts.write` as
  `<story id>/source.docx` inside the plugin's part namespace. The bundle never reads it
  back: no `host.parts.read` call exists.
- **In memory, for the session:** the original bytes, the story ids, the paragraph count
  poured per story and the column flag (`last` in `activate.ts`); the lowering, for the
  panel. The exporter works from these bytes, so save-back to `.docx` is available only in
  the session that imported the file. The provenance map is rebuilt by `DocSession::load`.

## Host doors

| Door | What the plugin uses it for |
|---|---|
| `host.contribute.panel`, `.command`, `.menu` | the "Document outline" panel, one command, one menu entry (skipped when the host has no menu door) |
| `host.contribute.importer`, `.exporter` | open a `.docx` or `.dotx`; save a `.docx` |
| `host.contribute.objectType`, `.editContext` | the `wordDocument` type; its double-click edit context names the host's own tools `paged.tool.type` and `paged.tool.select` |
| `host.document.mutate` | every write: styles, the pour, the text frame, grow rules |
| `host.document.meta`, `.collection("pages")`, `.hitTest` | the page to place on; the story of the new frame |
| `host.document.storyContent` | read the edited stories back for save-back |
| `host.nativeDocument.open` | open the skeleton as the document |
| `host.parts.write`, `host.document.setMetadata` | store the original file and bind it to a frame |
| `host.diagnostics.set` | lowering diagnostics on a place (`media.paged.doc`), and the `/styles`, `/open`, `/columns`, `/save-back` keys |
| `host.shell.pickFile`, `.openPanel`, `host.selection.set`, `host.supports`, `host.log` | pick a file; open the panel; the panel's "Select" button; probing optional doors; logging |

After a standalone open the lowering's diagnostics are shown in the panel but not sent to
`host.diagnostics`. The manifest declares `document` (read `broad`, write `scoped`,
`readNative`, `openNative`), `rendering` (`hitTest`), `editContext` (`wordDocument`),
`clipboard: "none"` and one wasm module, `bin/docx_js_bg.wasm`, with `maxBytes` 8388608. No
code uses `readNative` or the declared part type `docLowered`.

## Build and test

- `bash scripts/build-wasm.sh` builds `docx-js` for `wasm32-unknown-unknown` in release
  mode, checks that the `wasm-bindgen` CLI equals the version in `Cargo.lock`, runs
  `wasm-bindgen --target web` into `packages/doc-bundle/bin/` (gitignored) and, if present,
  `wasm-opt -Oz`. It fails above 100 000 000 bytes, a higher limit than the manifest's
  `maxBytes`. `.github/workflows/publish.yml` runs it, then `pnpm -r build`, then publishes
  `@paged-media/doc` under the `canary` tag unless that version exists.
- Rust lane: `.github/workflows/rust.yml` runs `cargo fmt --all --check`, clippy with
  warnings as errors, `cargo deny`, `cargo nextest run --workspace --profile ci` and the
  wasm build. Most tests are in `docx-conformance/tests/` and call the crates natively. The
  tests of layout rules read Word's recorded answers and need neither Word nor the host
  engine ([ADR 604](adr/604-word-is-the-oracle.md)). Two lanes over real Word files are
  ignored by default and need a corpus outside this repo: `real_word_corpus.rs`
  (`PAGED_DOC_CORPUS`) and `poi_corpus.rs` (`PAGED_DOCX_CORPUS`).
- TypeScript lane: `pnpm test` runs the import lint, then vitest in both packages;
  `.github/workflows/vitest.yml` adds a typecheck and a manifest check with the published
  `@paged-media/plugin-cli`. The tests use recording hosts and stub engines, not the wasm.
- The `scripts/word-*-probe.sh` scripts put the questions to desktop Microsoft Word on
  macOS; the fixtures record their output. `scripts/real-docx-acceptance.sh` needs an editor
  checkout; its results are [`reference/acceptance-real-docx.md`](reference/acceptance-real-docx.md).
