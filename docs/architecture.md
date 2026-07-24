# paged.doc — architecture

This is the buildable architecture of `paged.doc` as shipped in the M0 + Tier-0
pass. The originating concept is [`base-idea.md`](./base-idea.md); the milestone
ledger is [`status.md`](./status.md).

## The one-paragraph shape

`paged.doc` makes Microsoft Word `.docx` a first-class Paged content type by
**lowering DOCX onto the engine's native text model** rather than forking a
layout engine — because WordprocessingML is the same "styled paragraphs of runs"
model the engine already implements for IDML. A `.docx` is parsed by a shared
OOXML foundation, mapped to a clean semantic model, lowered to a plugin-local
**Lowered IR**, and the IR is turned into host mutations (`createParagraphStyle`,
`applyStyle`, `insertText`, `createSwatch`, …) that build native stories, styles,
and swatches inside the editor. Nothing about the layout is re-implemented; the
engine's own Knuth–Plass pagination flows the result across pages.

## Crates (Rust workspace, top level)

```
                ┌──────────────────────────────────────────────┐
                │  paged-ooxml   (MPL-2.0 OR PMEL)              │
                │  · OPC container over zip (ordered parts,     │
                │    verbatim carry-through preservation)       │
                │  · [Content_Types].xml + _rels graph          │
                │  · typed part DOM via vendored ooxmlsdk 0.12  │
                └───────────────┬──────────────────────────────┘
                                │ (semantic parse of known parts)
   docx-core  ◄── docx-import ──┘        docx-lower ──► ir::LoweredDoc
   (frozen        (WordprocessingML      (docx-core -> native IR +
    WML view)      -> docx-core,          synthesized style catalog +
                   scanner/no-panic)      swatches + section geometry)
                                │
                          docx-js  (single wasm-bindgen surface;
                                    core.rs plain-Rust DocSession +
                                    lib.rs #[cfg(wasm32)] shim)
                                │
                          docx-conformance  (TEST-ONLY: fixture builders
                                             + byte-identical round-trip)
```

- **`paged-ooxml`** owns the *format-mechanical* layer, shared (later) with
  `paged.slide`/`paged.sheet`. It does **not** delegate the container to
  `ooxmlsdk`'s `parts` feature — `ooxmlsdk` does not document byte-exact fidelity,
  and the preservation invariant needs verbatim carry-through, so the container +
  carry-through are our own over `zip` + `quick-xml`. `ooxmlsdk` is used only for
  the typed DOM of the parts we read. **DOC-02** (the `ooxmlsdk` → `wasm32`
  feasibility spike) is **GREEN**.
- **`docx-import`** is a scanner, not a validator: a missing styles part yields an
  empty catalog; malformed input yields a partial model + an error, never a panic.
- **`docx-lower`** is pure `docx-core -> LoweredDoc`, no SDK and **no core
  dependency** (the isolation-superset rule). Because the host's only range-styling
  op is `applyStyle(named style)`, direct Word formatting (a bold word in a Normal
  paragraph) is lowered by **synthesizing** a named style carrying the override
  (deduped by property signature, `basedOn` the referenced style).

## TS packages (`packages/`)

- **`@paged-media/doc-host-model`** — pure `LoweredDoc -> Mutation[]`. Its only
  `@paged-media` dependency is `plugin-api` (types). Every id in the IR is already
  a fully-formed Paged token, so this is a *dumb* translator: swatches →
  `createSwatch`, styles → `create{Paragraph,Character}Style` + `setStyleProperty`,
  the body → one `insertText` + `applyStyle` over code-point ranges.
- **`@paged-media/doc`** (the bundle) — `manifest.json` (`media.paged.doc`) +
  `activate(host)`: engine boot, importer/exporter, `objectType` + `editContext`,
  a place command, and an outline panel. All DOCX semantics live in the wasm; this
  is thin glue.

## The two faces (embed / open)

- **Embedded** — a `.docx` placed in a frame: `placeEmbedded` creates a text
  frame, resolves its story via the `hitTest` read door, and pours the IR as one
  atomic batch, then persists the `source.docx` via `host.parts` and stamps the
  binding envelope. This is the demonstrable path this pass.
- **Standalone** — open a `.docx` as the whole canvas: routed via the importer,
  it probes `host.nativeDocument.open` (**DOC-01**, now a conditional SDK door).
  The `docx -> native-bytes` producer that a true fresh-document open needs is
  deferred (it would be a `plugin-publish` sibling, like `idml-import`); until then
  the importer degrades to embedded placement + an honest diagnostic.

## Preservation invariant

The original `.docx` OPC package is retained. A zero-edit round-trip is per-part
**byte-identical** (decompressed): unknown parts and untouched parts re-emit their
stored bytes verbatim. Only parts the user actually changed are re-serialized — and
that edited save-back (native → WordprocessingML projection) is **M2**, not this
pass; the exporter currently re-emits the retained package verbatim.

## Licensing

`paged-ooxml` is a shared embeddable library → **MPL-2.0 OR PMEL** (one crate;
overrides the workspace license). The `docx-*` crates + the bundle are
**AGPL-3.0-only OR PMEL**, like the other first-party plugins. `ooxmlsdk`
(MIT OR Apache-2.0) is a compatible vendored dependency of the MPL crate.
