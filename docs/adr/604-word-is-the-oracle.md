# ADR 604 — Microsoft Word is the test oracle

- **Status:** Accepted. Recorded retroactively on 2026-10-02 from the code at `76e1d06`.
- **Scope:** `docx-conformance` (builders, `fixtures/*.word.json`, tests), the probe scripts
  in `scripts/`, the measured table in `docx-lower/src/line_height.rs`

## Context

The lowering writes Word's layout rules as native properties
([ADR 600](600-docx-lowered-onto-native-model.md)). Whether a mapping is right can only be
judged against what Word itself does with the same file. The concept paper saw the problem
before any code: Word cannot run headless in CI, so references would come from PDFs that
Word exports, authored once, and from a headless LibreOffice as a continuous proxy
(`docs/concept.md:434-437`, section 9). The first probe script states the stance: Word is
the oracle for opening a `.docx` the way InDesign is for IDML
(`scripts/word-pagination-probe.sh:4-5`).

## Decision

A Word behaviour the lowering reproduces is asked of desktop Word. For the fixtures this
repository builds, Word's answer is committed as a JSON file, and the tests read that file.
Four later mappings (line height by the sizes on a line, a line holding a picture, a table
row taller than its page, kerning and ligatures off) rest instead on measurements of real
documents recorded in `docs/reference/acceptance-real-docx.md`
(`docx-conformance/tests/mixed_sizes.rs:19-22`, `docx-conformance/tests/pictures.rs:19-24`,
`docx-conformance/tests/tall_rows.rs:19-24`, `docx-lower/src/lib.rs:1066-1073`).

- Builders in `docx-conformance/src/lib.rs` (30 of them) assemble each small `.docx` in
  memory, "so the conformance suite carries no binary blobs"
  (`docx-conformance/src/lib.rs:20`).
- A probe script per topic (12 files `scripts/word-<topic>-probe.sh`) writes the fixture to
  disk, calls `scripts/word-pagination-probe.sh`, which has Microsoft Word open it and save
  it as PDF through `osascript`, and reads positions out of the PDF with `pdftotext`.
- The output, with a note of how it was produced, is committed as
  `docx-conformance/fixtures/<topic>.word.json`, 13 files: page sizes and the lines on each
  page, line and word boxes, footnote numbers, header and footer text per page.
- Tests in `docx-conformance/tests/` load the file with `include_str!` (12 of the 13 files)
  and hold the lowering against it. For the column and continuous-section fixtures a small
  model of the engine's layout rule, written in the test support code, places the lines.
- Measured numbers are also production data: `docx-lower/src/line_height.rs` carries a table
  of single-line heights per typeface, taken from the PDFs Word wrote.
- Word also judges save-back: two answer files record that Word opened an edited save
  without a repair prompt.

Real documents are checked separately and outside CI. Two ignored test lanes
(`PAGED_DOC_CORPUS`, `PAGED_DOCX_CORPUS`) run the importer over real Word files that are not
in this repository; "What it asserts is deliberately structural, not fidelity"
(`docx-conformance/tests/real_word_corpus.rs:39`). `scripts/real-docx-acceptance.sh`
compares the editor's pagination of real documents with Word's PDF.

## Evidence

- `scripts/word-pagination-probe.sh:2-12`, `:27-39` — the stance, and the Word automation
- `scripts/word-breaks-probe.sh:2-20` — one probe: build, ask Word, read the PDF
- `docx-conformance/fixtures/pagination.word.json:1-12`,
  `docx-conformance/fixtures/breaks.word.json:1-5` — an answer file and its provenance
- `docx-conformance/tests/breaks.rs:19-45` — a test that reads its answer file
- `docx-conformance/tests/support/layout.rs:19-48`, `docx-conformance/tests/columns.rs:49`,
  `docx-conformance/tests/continuous.rs:40` — the layout model and the tests that include it
- `docx-lower/src/line_height.rs:19-56` — the measured line heights and why they are a table
- `docx-conformance/fixtures/run-specials.word.json:2`,
  `docx-conformance/fixtures/symbols.word.json:2` — Word reopening an edited save
- `docx-conformance/tests/real_word_corpus.rs:31-43`, `:133`,
  `docx-conformance/tests/poi_corpus.rs:184`, `scripts/real-docx-acceptance.sh:1-14` — the
  lanes outside CI

## Alternatives considered

The LibreOffice proxy of the concept paper was not built: no file in the repository other
than that paper mentions LibreOffice. The same passage models the gate on the rendered-image
comparison used for IDML; the oracle here is structural instead (which line lands on which
page, and where). Reading line heights from font files was rejected in
`docx-lower/src/line_height.rs:38-46`: the lowering receives no font bytes, and the font
file on disk is not the face Word uses.

## Consequences

CI needs neither Word nor the host engine: it runs `cargo nextest run --workspace`
(`.github/workflows/rust.yml:47`) against the committed answers. Regenerating an answer
needs a Mac with desktop Word and the poppler tools. No test here runs the host engine, so a
change in the engine's layout does not turn these tests red.

A typeface that is not in the measured table falls back to the default theme faces' line
height and is named in a diagnostic (`docx-lower/src/line_height.rs:48-53`).
`pagination.word.json` is read by no test here; `docx-conformance/tests/sections.rs:5-6` and
`docx-skeleton/tests/skeleton.rs:19-21` cite it in comments. The claim of no binary blobs
has one exception, `docx-conformance/fixtures/annual-report.docx`
(`docx-conformance/tests/annual_report.rs:19-30`).

The acceptance run is a script because it needs documents and fonts that are not in the
repository (`docs/reference/acceptance-real-docx.md:55-56`). Its record says the target is
not met (`:3`, `:520`).

## Related

- [ADR 600](600-docx-lowered-onto-native-model.md), [ADR 603](603-two-entry-points.md) — the mappings under test; the open path whose pagination is measured
- [ADR 120](https://github.com/paged-media/core/blob/main/docs/adr/120-indesign-is-the-oracle.md), [ADR 653](https://github.com/paged-media/plugin-publish/blob/main/docs/adr/653-indesign-is-the-oracle.md) — the same stance with InDesign as the oracle
- [ADR 309](https://github.com/paged-media/plugin-sdk/blob/main/docs/adr/309-conformance-against-real-engine.md) — the plugin family's conformance rule; no test in this repository runs the engine
- ADR 029 — the proposal, not published, whose acceptance target the real-document run measures
