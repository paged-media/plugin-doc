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
}

export interface LoweredParagraph {
  paraStyleId?: string | null;
  runs: LoweredRun[];
  sourceIndex: number;
}

export interface LoweredStory {
  paragraphs: LoweredParagraph[];
}

export interface LoweredSection {
  pageWidthPt: number;
  pageHeightPt: number;
  marginTopPt: number;
  marginBottomPt: number;
  marginLeftPt: number;
  marginRightPt: number;
  columns: number;
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
  diagnostics: Diagnostic[];
}

/** Parse the JSON string `docx-js` produces into a typed [`LoweredDoc`]. */
export function parseLoweredDoc(json: string): LoweredDoc {
  return JSON.parse(json) as LoweredDoc;
}
