# paged.doc

**Microsoft Word (`.docx`) as a first-class [Paged](https://paged.media) content
type.**

`paged.doc` lets you place a Word document inside a Paged layout (**embedded**) or
open a whole `.docx` as the canvas (**standalone**) — and edit it with the native
editor. Unlike a viewer, it does not fork a Word layout engine: it **lowers**
WordprocessingML onto Paged's native text + flow model (styled paragraphs of runs,
paragraph/character styles, sections), so you get native-quality editing and the
engine's own pagination, print-correct by construction.

Word content is parsed and re-serialized through a shared OOXML foundation
(`paged-ooxml`, over the vendored [`ooxmlsdk`](https://github.com/KaiserY/ooxmlsdk)
crate) with a **preservation invariant**: the original package is retained, and a
zero-edit round-trip is byte-identical — Paged never destroys a document.

## Status

**M0 foundation + Tier-0 read path** — a `.docx` opens, its styled paragraphs,
runs, styles and page geometry lower to native content via host mutations, and the
package round-trips byte-identically. Edited save-back, tables/images/lists, and
true standalone open are on the roadmap. See [`docs/status.md`](docs/status.md).

## Layout

- `paged-ooxml/` — shared OPC container + carry-through + `ooxmlsdk` typed DOM
  (MPL-2.0 OR PMEL).
- `docx-core/`, `docx-import/`, `docx-lower/`, `docx-js/`, `docx-conformance/` —
  the Word engine (AGPL-3.0-only OR PMEL).
- `packages/doc-host-model/`, `packages/doc-bundle/` — the TS glue + the
  `@paged-media/doc` plugin bundle.
- `docs/` — the concept, architecture, and status.

## Build

```bash
cargo test --workspace          # the engine
bash scripts/build-wasm.sh      # the wasm artifact (-> packages/doc-bundle/bin)
pnpm install && pnpm test       # the bundle glue
```

## License

Dual-licensed **AGPL-3.0-only OR PMEL** (the Paged Media Enterprise License, a
commercial license from And The Next GmbH). The shared `paged-ooxml` library is
**MPL-2.0 OR PMEL**. See [`LICENSE.md`](LICENSE.md).
