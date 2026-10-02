# Documentation

What this folder holds.

- [`concept.md`](concept.md): the concept paper. Why a Word document is lowered onto the
  native text model instead of getting its own layout engine, the two modes (placed in a
  frame, opened as the document), the shared OOXML foundation and the preservation policy.
- [`architecture.md`](architecture.md): how it is built. The eight crates and two packages,
  the path from a `.docx` to native content, the two entry points, the pour, save-back,
  where data is stored, the host doors used, and how it is built and tested.
- [`status.md`](status.md): the running log of the work, increment by increment: what was
  built, how it was checked, what is deferred. It is a log, not a summary: an older entry
  can describe a state that later work has overtaken. For how the code is built today, read
  `architecture.md`.
- [`adr/`](adr/README.md): the decision records of this repository, 600–604, and a list of
  the proposals that concern it.
- [`reference/acceptance-real-docx.md`](reference/acceptance-real-docx.md): the editor's
  pagination of three real Word documents measured against Word's own PDF export: the
  numbers, every deviation, what was fixed and what is left.

## Decisions in other repositories that bind this one

These records live in other public paged-media repositories. The code here rests on each of
them. The last column says what the decision means for this plugin.

| ADR | Repository | Decision | What it means here |
|---|---|---|---|
| [010](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/010-raw-mutate-gate-capability-enforcement.md) | plugin-sdk | The raw-mutate gate and the capability enforcement line | The bundle writes with raw `document.mutate` operations under `document.write: "scoped"`. The manifest also declares what its other calls need: `rendering: ["hitTest"]` for the story lookup on a place, `document.openNative` for standalone open (`packages/doc-bundle/manifest.json`). |
| [017](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/017-importer-exporter-door-shape.md) | plugin-sdk | Importer and exporter door shape | Both halves are used. `.docx` and `.dotx` files are routed to this plugin's importer by extension or MIME type; it receives the file's name and bytes. The exporter returns the bytes and a file name for a `.docx` (`packages/doc-bundle/src/activate.ts`). |
| [305](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/305-doors-always-present.md) | plugin-sdk | Every door is always present; `supports()` reports a missing backend | Seven flags are probed with `host.supports(...)`. Without `document.openNative@1` an opened file is placed in a frame; without `document.readStory@1` the export is the unedited file; without `shell.pickFile@1` the command only logs; a contribution door the host lacks is skipped. No flag exists for a single operation or property path, so there a refused batch is the probe (`packages/doc-bundle/src/pour.ts`, `src/open.ts`). |
| [307](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/307-contract-as-peer-dependency.md) | plugin-sdk | Bundles take the contract packages as peer dependencies | `@paged-media/plugin-api` and `@paged-media/plugin-sdk` are peer dependencies of `@paged-media/doc` with the range `>=0.2.29-canary.0`, and pinned to `0.2.37-canary.0` for development (`packages/doc-bundle/package.json`). |
| [308](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/308-plugin-wasm.md) | plugin-sdk | Plugin wasm is a declared capability, loaded by the bundle, under one size budget | The manifest declares one wasm module, `docx-engine` at `bin/docx_js_bg.wasm`, purpose `compute`. `packages/doc-bundle/src/engine.ts` loads it with the wasm-bindgen glue, and `scripts/build-wasm.sh` fails above 100 MB. The manifest's own `maxBytes` is lower, 8388608. |
| [310](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/310-one-write-door.md) | plugin-sdk | One write door: `document.mutate`, engine-owned history, failures as outcomes | Styles, text, tables, the text frame and the grow rules all go through `host.document.mutate`. The code reads `outcome.applied` and `createdId` (the new frame, the new table) and keeps no history of its own; the style catalogue and each text step are sent as one `batch`. |
| [311](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/311-plugin-state-under-own-id.md) | plugin-sdk | Plugin state lives only under the plugin's own id | The original `.docx` is written as a part with a path relative to the plugin's own part namespace, and the frame is bound to it through `host.document.setMetadata` (`packages/doc-bundle/src/place.ts`, `src/open.ts`). |
| [314](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/314-plugin-shape.md) | plugin-sdk | The plugin shape: semantics in Rust behind one wasm module, a logic-free shim, one published package | All four parts are here: the Rust crates, `DocEngine` forwarding to `DocSession` in `docx-js`, the private `doc-host-model`, and the one package `@paged-media/doc`. |
| [315](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/315-isolation-contract.md) | plugin-sdk | The isolation contract: a plugin depends only on the published contract | `deny.toml` denies git sources and unknown registries, and `scripts/check-contract-imports.mjs` runs before the tests. Where an engine shape is needed the plugin keeps its own copy (`StoryContentIn` in `docx-export/src/overlay.rs`), and `docx-skeleton` writes IDML without an engine crate. |
| [316](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/316-native-content-and-baking.md) | plugin-sdk | Plugin content is stored as valid native document content; baking is the fallback | A Word document becomes native styles, stories, tables and frames from the start, so the plugin has no baking step. The `wordDocument` object type declares `bakedFallback: "group"`. |
| [007](https://github.com/paged-media/core/blob/main/docs/adr/007-carry-through-rendering-honesty.md) | core | Rendering and save-back honesty: parse, do not fake; carry through | What cannot be placed is reported as a diagnostic and not imitated: footnote text is not inlined into the flow. Save-back keeps the original package and patches only what changed; an edit it cannot place is skipped and listed. See [ADR 600](adr/600-docx-lowered-onto-native-model.md) and [ADR 602](adr/602-save-back-is-a-byte-splice.md). |
| [026](https://github.com/paged-media/core/blob/main/docs/adr/026-auto-growing-region-chains.md) | core | Auto-growing region chains: pages grow in core, at composition time | Standalone open sets `setFlowGrowRule` on each section story and lets the engine add the pages. The plugin writes one page per story into the skeleton and inserts no other (`packages/doc-bundle/src/open.ts`). |
| [028](https://github.com/paged-media/core/blob/main/docs/adr/028-pagination-rules-are-engine-owned.md) | core | Pagination rules are engine-owned and shared by every format | Word's breaks, keep rules, widow control and column changes are lowered onto native paragraph properties (`paragraphStartParagraph`, `paragraphKeepWithNext`, the span and split column properties); where a page ends is left to the engine (`docx-lower/src/lib.rs`, `docx-lower/src/sections.rs`). |
| [118](https://github.com/paged-media/core/blob/main/docs/adr/118-paged-file-is-a-valid-idml-package.md) | core | A `.paged` file is a ZIP that stays a valid IDML package | The original `.docx` is stored as a part of the document's container through `host.parts.write`, so it is saved with the document (`packages/doc-bundle/src/place.ts`). |
| [024](https://github.com/paged-media/editor/blob/main/docs/adr/024-context-sensitivity-is-a-core-concept.md) | editor | Context-sensitivity is a core concept | The `wordDocument` edit context declares the host's own tools `paged.tool.type` and `paged.tool.select` and the outline panel, because the content is native text (`packages/doc-bundle/src/activate.ts`); `packages/doc-bundle/test/activate.spec.ts` checks the declaration. |
