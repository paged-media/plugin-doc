# CLAUDE.md — paged-media/plugin-doc

Orientation for Claude sessions in **paged-media/plugin-doc** — the paged.doc
Word/DOCX subsystem, delivered as a Paged plugin (dual-licensed AGPL-3.0-only OR
PMEL, And The Next GmbH; the shared `paged-ooxml` crate is MPL-2.0 OR PMEL).

## What this is

A Rust/WASM engine that makes Microsoft Word **`.docx`** a first-class Paged
content type, in two symmetric modes — **embedded** (a `.docx` in a frame of a
host layout) and **standalone** (open a whole `.docx` as the canvas). The central
bet: **lower DOCX onto the engine's NATIVE text + flow stack** (native
stories/paragraphs/runs + styles), not onto a forked layout engine — because
WordprocessingML is the same styled-paragraphs-of-runs model the engine already
implements for IDML. DOCX is parsed/re-serialized through a shared **`paged-ooxml`**
foundation over the vendored **`ooxmlsdk`** crate.

Spec (the concept authority): [`docs/base-idea.md`](docs/base-idea.md) (kept IN
this repo — it is self-documenting). Architecture: [`docs/architecture.md`](docs/architecture.md).
Status/milestones: [`docs/status.md`](docs/status.md).

**STATUS: M0 foundation + M1 Tier-0 read path shipped** (green). Rust crates
(Cargo workspace, top level): `paged-ooxml` (OPC + carry-through + ooxmlsdk typed
DOM), `docx-core` (frozen semantic WML view), `docx-import` (WML → docx-core),
`docx-lower` (docx-core → Lowered IR), `docx-js` (wasm-bindgen surface),
`docx-conformance` (TEST-ONLY). TS packages (`packages/*`): `doc-host-model` (pure
IR → Mutation[]) + `doc-bundle` (`@paged-media/doc`: manifest + `activate`).

## Project state & feature matrix (paged-media/state)

The canonical feature inventory + live status for ALL Paged repos live in
`paged-media/state` (dashboard: https://state.paged.media). There is NO feature
matrix in this repo; do not create one. NEW CAPABILITY → registry row; STATUS
CHANGE → registry, not prose.

## Hard rules (this repo's constitution)

- **ALL DOCX SEMANTICS LIVE IN RUST.** OPC parsing, WordprocessingML mapping,
  native lowering, preservation, and save-back are `paged-ooxml` + `docx-*` crates
  compiled to ONE wasm module (`docx-js`). The TS packages are thin glue: bundle
  lifecycle, panel, file input, and translating the engine's already-computed
  Lowered IR into host mutations. **Never implement a DOCX operation in
  TypeScript** — if the bundle seems to need one, the missing piece is a `docx-js`
  API.
- **ISOLATION CONTRACT, superset.** Zero core contact AND zero inter-plugin
  contact: the only `@paged-media/*` dependencies are `plugin-api`, `plugin-sdk`,
  and published contracts — never core, never another plugin, not at build time,
  runtime, or via side channels. `docx-lower` is pure model→IR with **no core
  dependency**; the native model is reached only through `host.document.mutate`.
  Rust guard: `deny.toml` `[sources]`. TS guard: `scripts/check-contract-imports.mjs`.
  SDK gaps become RFI entries — NEVER core modifications from this repo.
- **PURE LOWERING.** `docx-lower` is pure `docx-core → LoweredDoc`;
  `doc-host-model` is pure `LoweredDoc → Mutation[]`. Every behavior change lands
  with a test.
- **PRESERVATION INVARIANT.** "Paged never destroys a document." Unknown parts +
  unknown subtrees survive re-emission byte-identical; understood-but-untouched
  parts re-emit original bytes (lazy-verbatim). Zero-edit round-trip is per-part
  byte-identical (decompressed).
- **HONEST DEGRADATION (ADR-007).** Constructs past the current tier are surfaced
  as diagnostics on open, never silently dropped or faked.
- **LICENSE ASYMMETRY.** `paged-ooxml/**/*.rs` carry the MPL/PMEL header (crate
  `license = "MPL-2.0 OR LicenseRef-PMEL"`); `docx-*/**/*.rs` carry the AGPL/PMEL
  header (the workspace default). TS files carry the AGPL header block, no MPL.
- **ooxmlsdk is vendored + pinned** (the DuckDB-WASM / Blitz posture): treat
  upstream as a starting point; be ready to carry patches. OPC/carry-through is
  ALWAYS our own `zip`+`quick-xml` — never trust `ooxmlsdk` for byte-exact fidelity.

## Commands

```bash
# Rust (the engine)
cargo build --workspace && cargo test --workspace
cargo clippy --workspace --all-targets -- -D warnings
cargo fmt --all --check
cargo deny check

# wasm artifact (8 MiB budget; lands in packages/doc-bundle/bin/)
bash scripts/build-wasm.sh

# TS (the bundle) — consumes the PUBLISHED @paged-media canaries (no link chain)
pnpm install && pnpm typecheck && pnpm test
pnpm validate:manifest
```
