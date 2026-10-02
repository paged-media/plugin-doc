# ADR 603 — Two entry points: Open opens the file as the document, Insert places it in a frame

- **Status:** Accepted. Recorded retroactively on 2026-10-02 from the code at `76e1d06`.
- **Scope:** `packages/doc-bundle` (`activate.ts`, `open.ts`, `place.ts`, `menu.ts`,
  `manifest.json`), `docx-skeleton`, `docx-lower/src/sections.rs`

## Context

The concept paper asks for two modes of one engine: a Word document embedded in a frame of
another document, and a Word document opened as the whole document. The two differ only in
placement (`docs/concept.md:43-71`, section 2). The second needs a host door that opens
bytes as a new document; the paper lists it as missing (`:55`), and `docs/status.md:314-318`
records that `host.nativeDocument.open` exists since.

When the open path was first built (commit `c73381b`), the insert command went through it
as well, and inserting a Word file replaced the document the user had open. Commit `e2a3c23`
separated the two; the comment at `packages/doc-bundle/src/activate.ts:65-70` keeps the reason.

## Decision

The bundle registers two entry points. They share the lowering, the pour and save-back, and
neither is routed through the other.

- **Open.** The importer for `.docx` and `.dotx` calls `ingest(…, "open")`. On a host that
  reports `document.openNative@1` the Word file becomes the whole document. On any other
  host it is placed embedded and a log line says so.
- **Insert.** The command `media.paged.doc.command.placeDoc`, the panel's button and the
  menu entry `Object/Insert Word document…` call `ingest(…, "place")`, which never opens.
- Embedded placement makes one text frame on the active page (the first page if none is
  active), finds its story by a hit test at the frame's centre, applies the styles and
  pours. The frame's bounds come from the Word document's first section (page size less
  margins), not from the host page.
- Standalone open is a skeleton plus a pour. The plugin's wasm writes a minimal IDML package
  itself (`docx-skeleton`: `zip` and string templates, no engine crate): per native story a
  page with the size of the story's first section, a text frame on that section's margin
  box with a `LeadingOffset` first baseline, zero insets and the story's frame columns, and
  a story with one empty paragraph.
  The bundle opens it with `host.nativeDocument.open`, applies the style catalogue, sets
  `setFlowGrowRule` with `grow: true` and `copyFrameOptions: true` on every story, then
  pours each group of blocks into its story. The host engine adds the pages. Why the
  content is poured and not written into the package: The repository does not record why.
- A Word section that starts a page is its own story. A `continuous` or `nextColumn` section
  joins the story before it where the rules in `docx-lower/src/sections.rs` find a native
  expression for what Word does. The skeleton and the lowering use the same placement.
- Both paths write the source file as the part `<story id>/source.docx` and stamp one frame
  with metadata naming it; the object type `wordDocument` matches on that metadata.

## Evidence

- `packages/doc-bundle/src/activate.ts:65-129`, `:191-234`, `:249-261` — `ingest` and its
  two modes; command, panel and importer; the object type
- `packages/doc-bundle/src/menu.ts:26-28` — the menu entry
- `packages/doc-bundle/src/place.ts:61-142` — embedded placement
- `packages/doc-bundle/src/open.ts:19-28`, `:71-76`, `:116-191` — the open path, the grow
  rule, the four steps
- `docx-skeleton/src/lib.rs:19-45`, `:69-88`, `:302-347` — the skeleton, its ids, the spread
- `docx-lower/src/sections.rs:19-48`, `:76-79` — which sections share a story; the fallback
- `packages/doc-bundle/manifest.json:8-15`, `:26-37` — capabilities and contributions

## Alternatives considered

One path for both, removed in `e2a3c23` as described. A skeleton built as a `.paged` file
with the engine's own crates as git dependencies (commit `c73381b`), replaced the same day
by `dc65ac9`, whose message says those dependencies broke the isolation rule and failed the
dependency check; the crate comment adds "IDML is a public interchange format, like DOCX"
(`docx-skeleton/src/lib.rs:22`). A frame chain managed by the plugin for an embedded
document that overflows (`docs/concept.md:380-382`, section 8.1): not built. A page break
at every column change: kept as the lowering for an engine that refuses span and split
columns.

## Consequences

The source part is written and never read back: no code calls `host.parts.read`, the
metadata's part path is not followed, and the exporter works from bytes held in memory for
the last import (`packages/doc-bundle/src/activate.ts:45-59`, `:146`). Saving as `.docx`
therefore works only in the session that imported the file
([ADR 602](602-save-back-is-a-byte-splice.md)). The manifest also declares a part type
`docLowered` that is never written and a capability `readNative` that no call uses.

Embedded placement is one frame with no grow rule and no linked frames. IDML cannot carry
the grow rule, so it is set after the open; an engine without the operation leaves each
section on one page and the bundle reports it (`packages/doc-bundle/src/open.ts:78-113`).
The skeleton writes no master spreads (`docx-skeleton/src/lib.rs:326`).

Page growth is the host engine's mechanism (ADR 026 of the engine; the proposal ADR 029 is
not published). The repository's measurement on three real documents records the pagination
target as not met (`docs/reference/acceptance-real-docx.md:3`, `:520`). `README.md:20-23`
still lists standalone open as future work.

## Related

- [ADR 600](600-docx-lowered-onto-native-model.md), [ADR 602](602-save-back-is-a-byte-splice.md) — what is poured; how either mode saves
- [ADR 017](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/017-importer-exporter-door-shape.md), [ADR 305](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/305-doors-always-present.md), [ADR 311](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/311-plugin-state-under-own-id.md), [ADR 315](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/315-isolation-contract.md) — the importer door; probing a door and falling back; parts and metadata under the plugin's id; isolation
- [ADR 026](https://github.com/paged-media/core/blob/main/docs/adr/026-auto-growing-region-chains.md), [ADR 028](https://github.com/paged-media/core/blob/main/docs/adr/028-pagination-rules-are-engine-owned.md), [ADR 118](https://github.com/paged-media/core/blob/main/docs/adr/118-paged-file-is-a-valid-idml-package.md) — page growth and break rules in the engine; the container the source part travels in
- ADR 029, ADR 033 — proposals, not published, that concern standalone open and headers and footers
