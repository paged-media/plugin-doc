# ADR 602 — Save-back is a byte-level splice into the original file, never a regeneration

- **Status:** Accepted. Recorded retroactively on 2026-10-02 from the code at `76e1d06`.
- **Scope:** `docx-export`, `docx-js/src/core.rs`, the exporter in `packages/doc-bundle/src/activate.ts`

## Context

A Word file becomes native content ([ADR 600](600-docx-lowered-onto-native-model.md)), so
edits happen in the host's document and must be projected back on save. Two written rules
constrain how. Preservation: "Paged never destroys a document." (`CLAUDE.md:73-76`), which
the concept paper turns into "Edits patch, they don't rewrite." (`docs/concept.md:450`,
section 10). Size: the first call into the `ooxmlsdk` serializer
"links the whole WML write codegen" and exceeds the wasm budget (`scripts/build-wasm.sh:23-26`).

The plugin also has to learn what was edited. The host's whole-document read returns
"opaque core-native bytes this isolation-clean plugin cannot diff"
(`docx-js/src/core.rs:139`). A structured read of one story, `host.document.storyContent`,
was added to the host for this (`docs/status.md:319-324`, `:211-217`).

## Decision

The original `.docx` is kept and patched. Only the main document part (`word/document.xml`)
is changed, and in it only the byte ranges of the elements that changed. What changed is
found at save time by reading the stories back from the host and comparing them with a fresh
lowering of the kept source.

- At import, each paragraph, run and table-cell paragraph records its ordinal in the source
  XML. `build_bindings` maps lowered `(block, run)` coordinates to those ordinals.
- The exporter reads every story with `storyContent`, merges the per-section stories into
  one body and hands it to the engine with the kept source bytes. The engine lowers the
  source again as the baseline, overlays the read-back text and character style tokens, and
  diffs the two lowerings into an `EditSet`.
- Runs are compared by resolved formatting, because synthesized style ids renumber between
  lowerings. Blocks and table rows are aligned by a longest common subsequence over an
  identity key. A synthesized style goes back as direct `<w:rPr>` or `<w:pPr>`,
  "so Word gets no synthetic-style clutter" (`docs/status.md:188`).
- In the patcher, "quick-xml is used ONLY as a locator" (`docx-export/src/splice.rs:19`):
  one streaming pass finds the ranges, hand-rendered fragments fill them, every other byte
  is copied. `OpcPackage` writes all other parts from their stored bytes. The wasm build
  fails if `docx-export/src` calls `serialize_root(`, `.write_to(` or `.to_xml(`.
- An edit the patcher cannot place is skipped and named; the list reaches the host's
  diagnostics under `media.paged.doc/save-back`. Without the read door, without a story id,
  or on an error, the exporter returns the source bytes unchanged.

## Evidence

- `docx-export/src/lib.rs:19-27`, `:49-61`, `:298-316` — the contract of `apply_edits`, the
  skip list, and that only the main part is replaced, and only if its bytes changed
- `docx-export/src/splice.rs:19-24`, `scripts/build-wasm.sh:23-31` — the locator; the ban
- `docx-export/src/bindings.rs:19-42`, `docx-core/src/lib.rs:100`, `:129`, `:192`, `:305` —
  the provenance map and the ordinals recorded at import
- `docx-export/src/overlay.rs:19-27`, `:96-140`, `docx-export/src/diff.rs:19-28`, `:43-47` —
  overlay, comparison by resolved formatting, alignment by identity
- `docx-js/src/core.rs:125-161` — `save_verbatim`, `save_edited`, `save_edited_from_content`
- `packages/doc-bundle/src/activate.ts:142-189` — the exporter and its fallbacks
- `docx-conformance/tests/save_back.rs:19-24`, `:120-131` — what the tests of the patcher
  assert: the targets changed, every other part and untouched subtree byte-identical

## Alternatives considered

Serialising through the typed DOM: `serialize_root` exists
([ADR 601](601-own-package-container.md)) and the build guard forbids it here. Tracking
edits instead of reading back: the concept paper planned a back-reference on each native
node and a stored `bindings.json` (`docs/concept.md:166-169` in section 4.3, `:326` in
section 6); neither was built, the bindings are rebuilt from the kept source on every load
(`docx-js/src/core.rs:59-71`).
Pairing blocks and rows by index shipped first and kept the wrong `<w:p>` or `<w:tr>` node
after a deletion in the middle; identity alignment replaced it (`docs/status.md:462-498`).

## Consequences

Parts other than the main document cannot be changed on save: headers, footers, notes and
styles go back as they came. With the read door a save without edits is rewritten through
`OpcPackage`: every part is byte-identical, the ZIP itself is not.

The live path is narrower than the patcher. The overlay starts from a copy of the baseline
lowering and walks the baseline's blocks in order: each paragraph block takes the next
read-back paragraphs by position and receives only their run text and character style
tokens; table blocks are stepped over; the read-back paragraph style is parsed and not
applied (`docx-export/src/overlay.rs:96-140`). The overlay has no code that adds or removes a
block, does not read read-back paragraphs beyond those the baseline's blocks take, and
reports nothing. The edited lowering therefore always has the baseline's blocks, paragraph
styles and tables. No test of this path adds or removes a paragraph in the read-back.
Paragraph insert and delete, paragraph formatting, and table cell, row and column edits are
tested only through an `EditSet` or an edited lowering built by hand
(`docx-conformance/tests/save_back.rs`).

Save-back works only in the session that imported the file: the exporter uses the bytes held
in memory for the last import and returns `null` without one
(`packages/doc-bundle/src/activate.ts:45-59`, `:146`); the stored source part is written and
never read ([ADR 603](603-two-entry-points.md)). The baseline must be lowered at save as it
was at import (`packages/doc-bundle/src/engine.ts:93-99`).

## Related

- [ADR 601](601-own-package-container.md), [ADR 600](600-docx-lowered-onto-native-model.md), [ADR 603](603-two-entry-points.md) — the container that carries untouched parts; the lowering that is diffed; where the source bytes come from
- [ADR 007](https://github.com/paged-media/core/blob/main/docs/adr/007-carry-through-rendering-honesty.md), [ADR 017](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/017-importer-exporter-door-shape.md), [ADR 315](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/315-isolation-contract.md) — carry-through; the exporter door; why the engine's model bytes are opaque to a plugin
- [ADR 503](https://github.com/paged-media/plugin-sheets/blob/main/docs/adr/503-xlsx-patched-not-regenerated.md), [ADR 652](https://github.com/paged-media/plugin-publish/blob/main/docs/adr/652-idml-save-back-patches.md) — the same rule for XLSX and for IDML
