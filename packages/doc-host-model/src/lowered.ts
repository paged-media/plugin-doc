// paged.doc — the Lowered IR, mirrored on the TS side.
//
// This is the exact JSON shape `docx-lower` (Rust) serializes and `docx-js`
// hands across the wasm boundary (serde `rename_all = "camelCase"`). Keeping a
// local structural twin — rather than importing anything from the engine — keeps
// this package dependency-free apart from `@paged-media/plugin-api`.

/** A tab stop, shaped as the host `TabStopSpec` (position in points). */
export interface TabStopSpec {
  position: number;
  alignment?: string;
  alignmentCharacter?: string;
  leader?: string;
}

/** A wire `Value` (the union `docx-lower`'s `PropValue` serializes to). */
export type PropValue =
  | { type: "text"; value: string }
  | { type: "length"; value: number }
  | { type: "bool"; value: boolean }
  | { type: "colorRef"; value: string }
  | { type: "tabStops"; value: TabStopSpec[] };

/** A single style-property assignment. */
export interface StyleProp {
  /** A `PropertyPath` wire string, e.g. `"characterFontStyle"`. */
  path: string;
  value: PropValue;
}

export type StyleCollection = "paragraph" | "character";

/** A native style to create + populate. */
export interface LoweredStyle {
  /** Full token, e.g. `ParagraphStyle/docx-Heading1`. */
  id: string;
  name: string;
  collection: StyleCollection;
  basedOn?: string | null;
  props: StyleProp[];
}

/** A color to mint via `createSwatch`. */
export interface LoweredSwatch {
  /** `Color/docx-RRGGBB`. */
  id: string;
  name: string;
  /** `"RGB"` this pass. */
  space: string;
  /** Channel values in `space` — `[r, g, b]` on 0–255. */
  value: number[];
}

export interface LoweredRun {
  text: string;
  charStyleId?: string | null;
  /** When the run is a hyperlink, its external target URL — the host-model emits
   *  an `insertHyperlink` over the run range so it becomes natively clickable.
   *  (The blue+underline look still rides on `charStyleId`.) Absent otherwise. */
  hyperlinkUrl?: string | null;
}

export interface LoweredParagraph {
  paraStyleId?: string | null;
  runs: LoweredRun[];
  /** Inline images anchored to this paragraph (placed via insertAnchoredFrame). */
  images?: LoweredImage[];
  sourceIndex: number;
  /** ADR 028/029 — where a page/column break INSIDE the Word paragraph
   *  splits it: the pour starts a new native paragraph at each `at`, styled
   *  with the segment's style (it carries the break-before rule). Absent for
   *  an unsplit paragraph. */
  segments?: LoweredSegment[];
}

/** The second or later part of a Word paragraph split by a break. */
export interface LoweredSegment {
  /** Contiguous char offset into the paragraph's run text where it starts. */
  at: number;
  paraStyleId?: string | null;
}

/** An image lowered to an anchored-frame placement. */
export interface LoweredImage {
  /** Contiguous char offset into the paragraph's run text where the picture
   *  is addressed (where it sits in Word's paragraph; `docx-lower` decides
   *  it). Absent from an older IR: the paragraph's start. */
  at?: number;
  widthPt: number;
  heightPt: number;
  /** A self-contained `data:<mime>;base64,…` URI. */
  uri: string;
  /** ADR 035 — a FLOATING Word drawing's position and wrap, carried for the
   *  later lowering. It is still placed inline (with a diagnostic) until the
   *  engine can create a positioned, wrapped anchored object. Absent for an
   *  inline picture. */
  float?: LoweredFloat;
}

/** A floating drawing's position and wrap, in points, Word's vocabulary. */
export interface LoweredFloat {
  horizontal?: LoweredFloatPosition;
  vertical?: LoweredFloatPosition;
  /** `wrapNone` / `wrapSquare` / `wrapTight` / `wrapThrough` / `wrapTopAndBottom`. */
  wrap: string;
  /** `bothSides` / `left` / `right` / `largest`. */
  wrapText?: string;
  distTopPt: number;
  distBottomPt: number;
  distLeftPt: number;
  distRightPt: number;
  behindDoc: boolean;
  allowOverlap: boolean;
  layoutInCell: boolean;
  locked: boolean;
  relativeHeight: number;
  /** `wp:simplePos` (x, y) from the page's top-left, when used. */
  simplePosPt?: [number, number];
}

/** One axis of a float's position. */
export interface LoweredFloatPosition {
  /** `page`, `margin`, `column`, `character`, `paragraph`, `line`, … */
  relativeFrom: string;
  offsetPt?: number;
  align?: string;
  /** `wp14:pctPos*Offset`, in percent. */
  percent?: number;
}

/** The body as a sequence of blocks (paragraphs + tables) in document order. */
export interface LoweredStory {
  blocks: LoweredBlock[];
}

export type LoweredBlock =
  | ({ kind: "paragraph" } & LoweredParagraph)
  | {
      kind: "table";
      rows: number;
      cols: number;
      columnWidthsPt: number[];
      cells: LoweredCell[];
      /** Non-empty: the table has a row taller than its page, which Word
       *  splits across pages and a native row never does. It pours as these
       *  paragraphs (its cells' text in reading order) instead of a table. */
      flow?: LoweredParagraph[];
    };

/** A native table to build via insertTable + per-cell insertText + setCellSpan. */
export interface LoweredTable {
  rows: number;
  cols: number;
  columnWidthsPt: number[];
  cells: LoweredCell[];
  /** Leading rows Word repeats on every page (`w:tblHeader`). */
  headerRows?: number;
  /** Each row's least height in points (`w:trHeight`), 0 where a row
   *  declares none; absent when no row does. */
  rowHeightsPt?: number[];
  /** See `LoweredBlock`'s table: the paragraphs it pours as when it cannot
   *  be a native table. */
  flow?: LoweredParagraph[];
}

/** One table cell, addressed by its resolved grid position. */
export interface LoweredCell {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  paragraphs: LoweredParagraph[];
  /** Native insets (top, left, bottom, right, pt) that give the cell Word's
   *  height and text width. */
  insetsPt?: [number, number, number, number];
  /** `TopAlign` / `CenterAlign` / `BottomAlign` (Word's `w:vAlign`). */
  vAlign?: string;
}

export interface LoweredSection {
  pageWidthPt: number;
  pageHeightPt: number;
  marginTopPt: number;
  marginBottomPt: number;
  marginLeftPt: number;
  marginRightPt: number;
  columns: number;
  /** ADR 029 — index into `story.blocks` of this section's first block. */
  firstBlock?: number;
  /** ADR 029 — the native story (skeleton page) this section pours into.
   *  Consecutive sections share one when Word continues the later section on
   *  the same page (continuous / nextColumn). Absent from older lowerings:
   *  every section is its own story. */
  story?: number;
  /** ADR 033 — the headers and footers this section shows (after Word's
   *  inheritance). Carried, not yet placed. */
  headerFooter?: LoweredHeaderFooter;
  /** ADR 034 — the section's own footnote numbering (Word reads nothing
   *  else). Absent: Word's defaults. */
  footnoteNumbering?: LoweredNoteNumbering;
  endnoteNumbering?: LoweredNoteNumbering;
}

/** One section's headers and footers, by part name (`word/header1.xml`). */
export interface LoweredHeaderFooter {
  header: LoweredHeaderFooterParts;
  footer: LoweredHeaderFooterParts;
  titlePage: boolean;
  headerDistancePt?: number;
  footerDistancePt?: number;
  pageNumberStart?: number;
  pageNumberFormat?: string;
}

/** The part of each kind; absent is blank. */
export interface LoweredHeaderFooterParts {
  default?: string;
  first?: string;
  even?: string;
}

/** Footnote/endnote numbering, Word's vocabulary; absent fields are Word's defaults. */
export interface LoweredNoteNumbering {
  numFmt?: string;
  numStart?: number;
  /** `continuous` / `eachSect` / `eachPage`. */
  numRestart?: string;
  /** `pageBottom` / `beneathText` / `sectEnd` / `docEnd`. */
  pos?: string;
}

export interface Diagnostic {
  severity: "info" | "warning" | "error";
  message: string;
  tier: number;
}

/** The whole Tier-0 lowering of one Word document body. */
export interface LoweredDoc {
  swatches: LoweredSwatch[];
  styles: LoweredStyle[];
  story: LoweredStory;
  section: LoweredSection;
  /** ADR 029 — every section in order (standalone open pours each into its
   *  own story). Absent from older lowerings: treat as `[section]`. */
  sections?: LoweredSection[];
  /** ADR 033 — `w:evenAndOddHeaders`: even pages show the `even` pair. */
  evenAndOddHeaders?: boolean;
  diagnostics: Diagnostic[];
}

/** Parse the JSON string `docx-js` produces into a typed [`LoweredDoc`]. */
export function parseLoweredDoc(json: string): LoweredDoc {
  return JSON.parse(json) as LoweredDoc;
}
