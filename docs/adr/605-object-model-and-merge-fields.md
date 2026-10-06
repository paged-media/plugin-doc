# ADR 605 — A placed Word document is an object model; merge fields bind through it

- **Status:** Accepted 2026-10-06 (on the `om/object-model` branch, with the SDK contract
  0.2.43). Implements ADR 323 (the plugin object-model contract) for paged.doc.
- **Scope:** `docx-core` (`MergeField`), `docx-import` (`parse_merge_field`), `docx-lower`
  (`LoweredRun::merge_field`), `docx-js` (`object_edits` / `save_object_edits`),
  `packages/doc-host-model` (`bodyRunLocation`), `packages/doc-bundle/src/object-model.ts`,
  `doc-state.ts`, `manifest.json` (`contributes.objectModel`)

## Context

ADR 323 makes plugin state addressable through `host.objects` for scripts, the Node CLI, data
bindings and schema-driven fields. A placed Word document was reachable only through the
outline panel and one selection-driven command, and Word's mail-merge fields were read as a
frozen field name (`MERGEFIELD`) with the column name and switches dropped, so nothing could
bind data to them.

## Decision

1. **Merge fields are parsed.** `docx-import` parses `MERGEFIELD` in both field forms
   (complex `fldChar`/`instrText`, with the instruction split across runs, and
   `w:fldSimple`): the column name (quoted names with spaces included), `\b`, `\f`, `\m`,
   `\v`, every `\*` switch, the verbatim instruction and the field's ordinal in its
   paragraph. Each result run carries it, and the lowering passes it on
   (`LoweredRun.mergeField`).
2. **Four kinds**, declared once in the manifest: `docxDocument` (read-only facts; its
   content is the `.docx`), `block` (paragraph, table, inline picture), `run` (`text`
   writable for body paragraphs) and `mergeField` (`value` writable). Ids are lowered story
   coordinates (`<frame>/b<i>/r<j>`, `<frame>/b<i>/f<o>`), the coordinates save-back patches
   by. A binding stores a selector over the field's name, never an id.
3. **The state is what the frame's label names**: the `.docx` part and a docLowered cache
   part, both content-addressed after a write (`sources/<sha256>.docx`,
   `lowered/<sha256>.json`). A missing cache is rebuilt from the `.docx`.
4. **A write is one undo step and touches three things together**: the engine patches the
   `.docx` (a run's `w:t`; a merge field's result runs, never its field code — the engine
   decides which runs the result spans), the poured native text changes in the same batch
   (`deleteRange` + `insertText` in the engine's byte-and-break offset space), and the
   label names the new parts. The native paragraph is checked against what the document
   last held first; a paragraph edited on the canvas since is refused, not overwritten.
   All writes of a batch are planned by one plugin-level planner, so run and field writes
   on one document compose.
5. **Typed commands**: `placeDoc` (the typed twin of the untyped command), `exportDocx`
   (the `.docx` bytes) and `fillMergeFields` (values by column name, one undo step).

## Consequences

- **The export side.** A written value lives in the `.docx` as the field's result with the
  field code intact, so Word shows the value and can still run its own mail merge on the
  file. In the IDML side the value is plain text: InDesign sees no Data Merge placeholder
  for it (ADR 559's native Data Merge form is the data plugin's to write).
- Run ids after an emptied result run move (an empty run does not lower); field and block
  ids do not.
- Table-cell runs and fields are read but not writable yet; a paragraph split by a page or
  column break is not writable either.
- Found while building the lane, not caused by it: exporting a document whose same-styled
  runs the engine coalesced duplicates the paragraph text (the story read-back's runs
  differ from Word's, and wrapped field runs cannot be deleted). The lane pins it with an
  expected failure.
