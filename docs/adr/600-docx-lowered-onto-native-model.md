# ADR 600 — DOCX is lowered onto the engine's native text and style model; there is no Word layout engine

- **Status:** Accepted. Recorded retroactively on 2026-10-02 from the code at `76e1d06`.
- **Scope:** `docx-lower`, `packages/doc-host-model`, the manifest and edit context of `packages/doc-bundle`

## Context

The concept paper names two ways a plugin can show foreign content: compile it to native
content that the host engine lays out, or lay it out inside the plugin and paint the result
into a frame (`docs/concept.md:119-122`, section 4.1). It chooses by distance from the
engine's own model. WordprocessingML is styled paragraphs and runs, styles, sections, tables,
images, lists and tabs: "the same concept set the engine implements for IDML" (`:130`), so
the step is a mapping, not a layout engine (`:131-132`).

The paper lists what this buys: editing with the host's own text tools and caret, pagination
by the host engine, one shaping stack (`:140-150`, section 4.2). The price: edits land on
the native model, so saving needs a projection back to WordprocessingML (`:157-159`,
section 4.3), the subject of [ADR 602](602-save-back-is-a-byte-splice.md).

## Decision

A `.docx` becomes ordinary native content (swatches, paragraph and character styles, stories
of paragraphs and tables, anchored image frames, hyperlinks), written through the host's
mutation door and laid out by the host engine. The plugin paints nothing.

- `docx-lower` is a pure function from the parsed document to a `LoweredDoc`: swatches, a
  style catalogue ordered parents first, a story of blocks, section geometry, diagnostics.
  Every id is a complete native token; every value serialises as the host wire's value union.
- `doc-host-model` maps that IR to mutations and invents no ids: `createSwatch`,
  `createParagraphStyle`, `createCharacterStyle`, `setStyleProperty`, `insertText`,
  `applyStyle`, `insertAnchoredFrame`, `insertHyperlink`, `insertTable`,
  `setElementProperty`, `setCellSpan`, `batch`.
- Direct formatting becomes a synthesized named style (`…/docx-auto-pN`, `…/docx-auto-cN`),
  based on the referenced style and reused when collection, parent and properties are equal,
  because "The host's only range-styling op is `applyStyle(named style)`"
  (`docx-lower/src/lib.rs:27`). `docDefaults` become the base style `…/docx-Default`.
- Word's layout rules are written as values of native properties: breaks and odd/even
  sections as `paragraphStartParagraph`, first-fit line breaking as
  `paragraphComposer = "HL Single"`, a column change inside a page as the span and split
  column properties, line spacing as `characterLeading`. Comments explain single mappings
  (`docx-lower/src/lib.rs:1058-1062`). For the rule as a whole:
  The repository does not record why.
- What has no native construct is not imitated (`CLAUDE.md:77-78`). Footnotes, headers and
  footers are parsed and carried in the IR but not placed; a floating picture is placed
  inline; a field other than `HYPERLINK` keeps its last computed text. Each adds a diagnostic.

## Evidence

- `docx-lower/src/lib.rs:19-32`, `docx-lower/src/ir.rs:19-30`, `:35-58` — the pure
  lowering, why direct formatting becomes styles, and the IR as the contract
- `packages/doc-host-model/src/mutations.ts:32-55`, `:224-261`, `:305-318`, `:383-420` — operations
- `docx-lower/src/lib.rs:1919-1944`, `:2596-2602`, `:1040-1092`, `:59-77`,
  `docx-lower/src/line_height.rs:19-26` — `synthesize`, its signature, the base style, and
  the property paths for breaks, the composer, columns and line pitch
- `docx-lower/src/lib.rs:199-227`, `:681-745`, `:747-779`, `:781-830` — floats, notes,
  fields, headers and footers: carried, diagnosed, not imitated
- `packages/doc-bundle/manifest.json:14`, `packages/doc-bundle/src/activate.ts:263-294` —
  the only rendering capability is `hitTest`; the edit context names the host's own tools

## Alternatives considered

The concept paper compares this route with an in-plugin Word layout engine painted into a
frame, and keeps that as an escalation if measurement shows the mapping cannot express Word's
layout (`docs/concept.md:174-190`, section 4.4). It also sketches a hybrid: a fixed,
non-editable render of pages the mapping cannot reach, with a `docx-render` crate and a
`sceneLayer` capability (`:190-193`, and `:290`, `:308` in section 6). Neither the crate nor
the capability exists (`Cargo.toml:11-20`, `packages/doc-bundle/manifest.json:7-25`). For
line breaking the paper thought a compatibility mode in the host's text engine likely
(`docs/concept.md:361-364`, section 7); the code selects the engine's existing single-line
composer.

## Consequences

A Word feature reaches the page only when a native construct and a mutation that creates it
exist. Footnotes, headers and footers, and floating drawings wait on that (the proposals
ADR 034, ADR 033 and ADR 035). Some operations and property paths exist only from later
engine protocol versions (`packages/doc-host-model/src/mutations.ts:238-240`, `:254-257`,
`:283-286`); an older engine refuses them and the bundle reports it
(`packages/doc-bundle/src/pour.ts:83-95`, `:128-135`). Line breaking and pagination are the host
engine's; the repository's measurement against three real documents records its page-match
target as not met (`docs/reference/acceptance-real-docx.md:3`, `:520`).

Lowering diagnostics reach the host's diagnostics only on embedded placement
(`packages/doc-bundle/src/place.ts:134-139`); after a standalone open they appear only in
the plugin's own panel (`packages/doc-bundle/src/panels/outline-panel.tsx:213-219`). The IR
is kept by hand in two places (`packages/doc-host-model/src/lowered.ts:1-6`). `README.md:20-23`
still lists tables, images, lists and standalone open as future work.

## Related

- [ADR 601](601-own-package-container.md), [ADR 602](602-save-back-is-a-byte-splice.md), [ADR 603](603-two-entry-points.md), [ADR 604](604-word-is-the-oracle.md) — reading, saving, the two entry points, and how the mappings are checked
- [ADR 314](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/314-plugin-shape.md), [ADR 316](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/316-native-content-and-baking.md), [ADR 310](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/310-one-write-door.md), [ADR 305](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/305-doors-always-present.md) — the plugin shape, native content, the write door, and how an older engine is detected
- [ADR 007](https://github.com/paged-media/core/blob/main/docs/adr/007-carry-through-rendering-honesty.md), [ADR 028](https://github.com/paged-media/core/blob/main/docs/adr/028-pagination-rules-are-engine-owned.md), [ADR 024](https://github.com/paged-media/editor/blob/main/docs/adr/024-context-sensitivity-is-a-core-concept.md) — report instead of imitating; the engine-owned break rules; the edit context declaration
- [ADR 020](https://github.com/paged-media/plugin-web/blob/main/docs/adr/020-paged-web-native-engine-defer-frame-threading.md) — the opposite choice for HTML: an in-plugin engine painted into a frame
