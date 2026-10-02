# ADR 601 — The package container is the plugin's own; the OOXML library is used for typed reads only

- **Status:** Accepted. Recorded retroactively on 2026-10-02 from the code at `76e1d06`.
- **Scope:** the crate `paged-ooxml`, and the part parsing in `docx-import`

## Context

A `.docx` is an OPC package: a ZIP of XML parts, a content-type table and a relationship
graph. The repository's preservation rule, "Paged never destroys a document."
(`CLAUDE.md:73-76`), requires every part the plugin does not change to come back byte for
byte, including parts it does not understand.

The crate `ooxmlsdk` offers code-generated typed trees for the parts and its own package
API. The concept paper assessed it (`docs/concept.md:209-226`, section 5.2): no documented
byte-exact round trip, element children flattened into enum vectors, no statement about
wasm. It concluded: use it as the typed DOM behind a wrapper, keep the container, and first
check that it builds for wasm. `Cargo.toml:9` records that check as passed on 2026-07-23.

## Decision

`paged-ooxml` owns the container and uses `ooxmlsdk` only to deserialise single parts. The
crate comment gives the reason: the container is not delegated
"because `ooxmlsdk` does not document byte-exact fidelity" (`paged-ooxml/src/lib.rs:28`) and
the preservation rule needs verbatim carry-through.

- `OpcPackage::read` opens the ZIP with the `zip` crate and keeps every entry, in on-disk
  order, as decompressed bytes. `write` emits each entry from those bytes; `set_part`
  replaces one and marks it changed.
- The guarantee is identity of each part's decompressed bytes. The file as a whole is not
  identical, because the deflate stream is encoded again.
- The relationship parts (`_rels`) are read by a tolerant `quick-xml` scan.
- `parse_root` turns one part's bytes into an `ooxmlsdk` root type. Its only callers are in
  `docx-import`, for eight roots: document, styles, numbering, settings, footnotes,
  endnotes, header, footer.
- `ooxmlsdk` is re-exported from `paged-ooxml` so that the other crates do not name its
  version. Only `docx-import` uses its types, and maps them into the plugin's own model
  `docx-core`; `docx-lower` reads that model and never the typed trees.
- Two guards run before the recursive parser. `parse_root` refuses XML nested deeper than
  256 elements, because nesting depth becomes stack depth. `OpcPackage::read` recognises a
  legacy binary `.doc` and an RTF file by their first bytes and returns a named error.

## Evidence

- `paged-ooxml/src/lib.rs:24-40`, `:56-59` — the two responsibilities, the reason for
  keeping the container, the re-export
- `paged-ooxml/src/opc.rs:17-33`, `:78-119`, `:148-170`, `:179-202` — the preservation
  model, `read` with the format checks, `set_part`, `write`
- `paged-ooxml/src/dom.rs:42-58`, `:85-94`, `paged-ooxml/src/rels.rs:15-21` — the depth
  limit, `parse_root`, the relationship scan
- `docx-import/src/lib.rs:35-44`, `:192-217`, `:316`, `:333`, `:488`, `:501` — the only
  imports of `ooxmlsdk` types and every call of `parse_root`
- `docx-core/src/lib.rs:19-27` — the model the rest of the pipeline reads
- `Cargo.toml:31-34`, `Cargo.lock:321-323` — `ooxmlsdk` 0.12 with the feature `mce`, `zip` 2,
  `quick-xml` 0.41; the lock file resolves `ooxmlsdk` 0.12.0 from crates.io

## Alternatives considered

The package API of `ooxmlsdk` (its `parts` feature) was not used, for the reason quoted
above (`paged-ooxml/src/lib.rs:26-32`, `paged-ooxml/src/opc.rs:17-22`). Had the wasm check
failed, the concept paper's fallback was a thinner own layer of OPC and `quick-xml` for the
parts the plugin touches (`docs/concept.md:223-226`, section 5.2). Writing parts through
the typed DOM is possible, `serialize_root` exists (`paged-ooxml/src/dom.rs:96-104`), and is
not done: see [ADR 602](602-save-back-is-a-byte-splice.md).

## Consequences

What the importer can read is bounded by the root types and the parser of `ooxmlsdk`. A file
written by `OpcPackage::write` is not bit-identical to its input even without an edit.
Nothing calls `serialize_root`; its comment says it is
"exercised now only for the round-trip identity harness" (`paged-ooxml/src/dom.rs:97`), and
no such caller exists. The reader for `[Content_Types].xml`
(`paged-ooxml/src/content_types.rs`) has no caller outside its own file either.

Three written claims do not match the code:

- `ooxmlsdk` is called "vendored" in `CLAUDE.md:82-84`, `README.md:14`, `deny.toml:7` and
  the crate comments, as the concept paper intended (`docs/concept.md:220`, section 5.2).
  It is an ordinary crates.io dependency with the requirement `"0.12"`, fixed only by
  `Cargo.lock`; the repository holds no copy of it. The repository does not record why.
- `paged-ooxml` is described as shared with the spreadsheet and presentation plugins
  (`paged-ooxml/src/lib.rs:17-22`). In this repository its only dependents are the `docx-*`
  crates, and its manifest sets `publish = false` (`paged-ooxml/Cargo.toml:9`). Its
  `license` field differs from the workspace default (`:6-8`; see `LICENSE.md`).
- On `mc:AlternateContent` the comments disagree: `paged-ooxml/src/lib.rs:38-40` leaves
  fallback selection to later, `paged-ooxml/src/dom.rs:24-26` says the typed DOM handles it
  at parse time. `docx-import` has no code that mentions it.

## Related

- [ADR 602](602-save-back-is-a-byte-splice.md), [ADR 600](600-docx-lowered-onto-native-model.md) — how a changed part is written; what the parsed model is lowered to
- [ADR 007](https://github.com/paged-media/core/blob/main/docs/adr/007-carry-through-rendering-honesty.md) — carry-through as a platform rule
- [ADR 503](https://github.com/paged-media/plugin-sheets/blob/main/docs/adr/503-xlsx-patched-not-regenerated.md) — the spreadsheet plugin's rule that XLSX is patched, never regenerated; `paged-ooxml/src/opc.rs:24` says this container mirrors that plugin's
- [ADR 314](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/314-plugin-shape.md), [ADR 315](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/315-isolation-contract.md) — the plugin shape and the isolation contract
