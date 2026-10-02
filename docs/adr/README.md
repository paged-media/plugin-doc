# Architecture decision records

An ADR records one load-bearing decision that has already been made: what was decided, what
in the code shows it, and what it obliges other code to do. It is a record, not a proposal.
When the code stops matching a record, the body is left as it is and a dated amendment is
added at the end.

ADR numbers are unique across the paged-media repositories, so a number names the same
record wherever it is cited. New records in this repository use 600–649. Lower numbers
cited here belong to other repositories or predate that scheme. Records 600–604 were written
on 2026-10-02 from the code as it stood, for decisions made earlier; their status says so.

| ADR | Title | Status |
|---|---|---|
| [600](600-docx-lowered-onto-native-model.md) | DOCX is lowered onto the engine's native text and style model; there is no Word layout engine | Accepted, recorded retroactively 2026-10-02 |
| [601](601-own-package-container.md) | The package container is the plugin's own; the OOXML library is used for typed reads only | Accepted, recorded retroactively 2026-10-02 |
| [602](602-save-back-is-a-byte-splice.md) | Save-back is a byte-level splice into the original file, never a regeneration | Accepted, recorded retroactively 2026-10-02 |
| [603](603-two-entry-points.md) | Two entry points: Open opens the file as the document, Insert places it in a frame | Accepted, recorded retroactively 2026-10-02 |
| [604](604-word-is-the-oracle.md) | Microsoft Word is the test oracle | Accepted, recorded retroactively 2026-10-02 |

Decisions made in other repositories that this plugin's code rests on are listed in
[`../README.md`](../README.md).

## Proposals

Four proposals concern this plugin: ADR 029, 033, 034 and 035. They are not published yet
and will be added here when they are accepted. Code comments in this repository cite them
by number; this is what each is about, as far as the code and
[`../status.md`](../status.md) show.

- **ADR 029**: standalone open. A Word document becomes the whole document, not content in
  a frame. The entry point as built is recorded in [ADR 603](603-two-entry-points.md); the
  measurements against the proposal's acceptance criterion are in
  [`../reference/acceptance-real-docx.md`](../reference/acceptance-real-docx.md).
- **ADR 033**: headers and footers. Every section's header and footer references are read
  and carried in the lowering; none is placed on a page yet.
- **ADR 034**: footnotes and endnotes. The notes and their numbering are parsed; no note is
  placed on a page yet.
- **ADR 035**: floating drawings. Position and wrap are carried in the lowering; the
  picture is placed inline, with a warning.
