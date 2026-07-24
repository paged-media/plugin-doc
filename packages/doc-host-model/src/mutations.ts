// paged.doc — pure Lowered IR -> host `Mutation[]`.
//
// The same role as `sheet-host-model`: a *dumb* translator. Every id in the IR
// is already a fully-formed Paged token, so this file never invents ids — it only
// maps IR nodes to `host.document.mutate(...)` ops. No engine logic lives here;
// the semantics were decided in `docx-lower` (Rust).

import type { Mutation } from "@paged-media/plugin-api";

import type {
  LoweredCell,
  LoweredDoc,
  LoweredParagraph,
  LoweredStyle,
  LoweredTable,
} from "./lowered.js";

/**
 * Count a string's length in Unicode scalar values (code points), matching the
 * `char`-offset convention of the engine's `InsertText`/`ApplyStyle` ops.
 */
function codePointLen(s: string): number {
  return Array.from(s).length;
}

/**
 * Swatch + style-catalog mutations. Emitted before any pour so `applyStyle`
 * references resolve. Swatches precede styles (a style's `characterFillColor`
 * references a swatch); styles are already topologically ordered by `docx-lower`.
 */
export function buildStyleMutations(ir: LoweredDoc): Mutation[] {
  const ops: Mutation[] = [];
  for (const sw of ir.swatches) {
    ops.push({
      op: "createSwatch",
      args: { spec: { selfId: sw.id, name: sw.name, space: sw.space, value: sw.value } },
    } as Mutation);
  }
  for (const style of ir.styles) {
    ops.push(createStyleOp(style));
    for (const prop of style.props) {
      ops.push({
        op: "setStyleProperty",
        args: { collection: style.collection, styleId: style.id, path: prop.path, value: prop.value },
      } as Mutation);
    }
  }
  return ops;
}

function createStyleOp(style: LoweredStyle): Mutation {
  const op = style.collection === "paragraph" ? "createParagraphStyle" : "createCharacterStyle";
  return { op, args: { selfId: style.id, name: style.name, basedOn: style.basedOn ?? null } } as Mutation;
}

// ---------------------------------------------------------------------------
// Text pour (a contiguous run of paragraph blocks)

interface Range {
  start: number;
  end: number;
  style: string;
  scope: "paragraph" | "character";
}

/** An inline image + the story offset of its anchoring paragraph. */
interface ImageAt {
  offset: number;
  widthPt: number;
  heightPt: number;
  uri: string;
}

/** A hyperlink span: the `[start, end)` story offsets of a linked run + target. */
interface LinkAt {
  start: number;
  end: number;
  url: string;
}

/** The joined text of a paragraph run + the style ranges over it + image
 *  placements, offsets relative to `base`.
 *
 *  IMPORTANT — offset convention: the story's char-offset space is CONTIGUOUS
 *  across paragraphs (the engine consumes the `\n` on paragraph split, it is not
 *  a stored character). So the inserted `text` carries `\n` separators (to create
 *  the paragraph breaks), but the style/image OFFSETS advance only by run text —
 *  never by the separator. `length` is the resulting contiguous story growth. */
function poured(
  paragraphs: LoweredParagraph[],
  base: number,
): { text: string; ranges: Range[]; images: ImageAt[]; links: LinkAt[]; length: number } {
  const ranges: Range[] = [];
  const images: ImageAt[] = [];
  const links: LinkAt[] = [];
  let text = "";
  let offset = base;
  paragraphs.forEach((para, pIdx) => {
    const paraStart = offset;
    for (const run of para.runs) {
      const runStart = offset;
      text += run.text;
      offset += codePointLen(run.text);
      if (run.charStyleId) {
        ranges.push({ start: runStart, end: offset, style: run.charStyleId, scope: "character" });
      }
      if (run.hyperlinkUrl && offset > runStart) {
        links.push({ start: runStart, end: offset, url: run.hyperlinkUrl });
      }
    }
    if (para.paraStyleId) {
      ranges.push({ start: paraStart, end: offset, style: para.paraStyleId, scope: "paragraph" });
    }
    // Images anchor at the paragraph level (paged anchors a frame to a
    // paragraph), so any offset within the paragraph resolves to it.
    for (const img of para.images ?? []) {
      images.push({ offset: paraStart, widthPt: img.widthPt, heightPt: img.heightPt, uri: img.uri });
    }
    // Separator text for insertText, but NOT an offset advance (contiguous).
    if (pIdx < paragraphs.length - 1) {
      text += "\n";
    }
  });
  return { text, ranges, images, links, length: offset - base };
}

/** insertText + applyStyle + insertAnchoredFrame (inline images) for a
 *  contiguous paragraph run, offsets from `base`. */
export function buildTextPour(
  paragraphs: LoweredParagraph[],
  storyId: string,
  base: number,
): { mutations: Mutation[]; length: number } {
  const { text, ranges, images, links, length } = poured(paragraphs, base);
  const ops: Mutation[] = [];
  if (text.length > 0) {
    ops.push({ op: "insertText", args: { storyId, offset: base, text, cell: null } } as Mutation);
  }
  for (const r of ranges.filter((r) => r.scope === "paragraph")) {
    ops.push(applyStyleOp(storyId, r.start, r.end, r.style, "paragraph"));
  }
  for (const r of ranges.filter((r) => r.scope === "character")) {
    ops.push(applyStyleOp(storyId, r.start, r.end, r.style, "character"));
  }
  for (const img of images) {
    // `insertAnchoredFrame` is a v52 wire op (core protocol 52); it postdates the
    // published plugin-api Mutation union, so cast via `unknown`. The host applies
    // it once running the v52+ canvas-wasm; older hosts reject it (honest degrade).
    ops.push({
      op: "insertAnchoredFrame",
      args: {
        storyId,
        offset: img.offset,
        width: img.widthPt,
        height: img.heightPt,
        imageUri: img.uri,
      },
    } as unknown as Mutation);
  }
  for (const link of links) {
    // `insertHyperlink` is a v53 wire op (core protocol 53) — like
    // insertAnchoredFrame it postdates the published Mutation union, so cast via
    // `unknown`. The engine mints the source/destination/hyperlink ids and makes
    // the span clickable; older hosts reject it (the blue+underline still shows).
    ops.push({
      op: "insertHyperlink",
      args: { storyId, start: link.start, end: link.end, url: link.url },
    } as unknown as Mutation);
  }
  return { mutations: ops, length };
}

function applyStyleOp(
  storyId: string,
  start: number,
  end: number,
  style: string,
  scope: "paragraph" | "character",
): Mutation {
  return { op: "applyStyle", args: { storyId, start, end, style, scope } } as Mutation;
}

// ---------------------------------------------------------------------------
// Tables

/** The `insertTable` op (its outcome mints the tableId). */
export function buildTableInsert(table: LoweredTable, storyId: string): Mutation {
  return {
    op: "insertTable",
    args: {
      storyId,
      rows: table.rows,
      cols: table.cols,
      headerRows: 0,
      footerRows: 0,
      columnWidths: table.columnWidthsPt,
      rowHeights: [],
    },
  } as Mutation;
}

/** One flattened cell's text (paragraphs joined by newline). NOTE: cell-internal
 *  paragraph/character styling is not applied — `applyStyle` carries no cell
 *  qualifier, so ranged styling can't reach cell interiors (a Tier-2 limitation);
 *  cell text is poured at the cell's default formatting. */
function cellText(cell: LoweredCell): string {
  return cell.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}

/** The cell-pour + merge batch for a resolved `tableId`: `insertText` per cell
 *  (addressed by TextCellAddr) + `setCellSpan` per merged cell. */
export function buildTableCells(table: LoweredTable, storyId: string, tableId: string): Mutation {
  const ops: Mutation[] = [];
  for (const cell of table.cells) {
    const text = cellText(cell);
    if (text.length > 0) {
      ops.push({
        op: "insertText",
        args: { storyId, offset: 0, text, cell: { tableId, row: cell.row, col: cell.col } },
      } as Mutation);
    }
    if (cell.rowSpan > 1 || cell.colSpan > 1) {
      ops.push({
        op: "setCellSpan",
        args: {
          storyId,
          tableId,
          row: cell.row,
          col: cell.col,
          rowSpan: cell.rowSpan,
          columnSpan: cell.colSpan,
        },
      } as Mutation);
    }
  }
  return { op: "batch", args: { ops } } as Mutation;
}

// ---------------------------------------------------------------------------
// The story plan (block-aware; tables need mid-execution tableId resolution)

/** One step the bundle executes in order against a resolved `storyId`. A text
 *  step builds its ops given the running story offset (advancing it by `length`);
 *  a table step inserts the table (its outcome mints the id), then pours cells. */
export type StoryStep =
  | { kind: "text"; length: number; mutations: (baseOffset: number) => Mutation[] }
  | { kind: "table"; insert: Mutation; cells: (tableId: string) => Mutation };

/** Split the story's blocks into executable steps: consecutive paragraph blocks
 *  coalesce into one text step; each table becomes a table step. */
export function buildStory(ir: LoweredDoc, storyId: string): StoryStep[] {
  const steps: StoryStep[] = [];
  let pending: LoweredParagraph[] = [];
  const flush = () => {
    if (pending.length === 0) return;
    const paras = pending;
    pending = [];
    steps.push({
      kind: "text",
      length: buildTextPour(paras, storyId, 0).length,
      mutations: (base) => buildTextPour(paras, storyId, base).mutations,
    });
  };
  for (const block of ir.story.blocks) {
    if (block.kind === "table") {
      flush();
      const table: LoweredTable = block;
      steps.push({
        kind: "table",
        insert: buildTableInsert(table, storyId),
        cells: (tableId) => buildTableCells(table, storyId, tableId),
      });
    } else {
      pending.push(block);
    }
  }
  flush();
  return steps;
}

/**
 * Everything needed to realize a TEXT-ONLY lowering into `storyId`, as one atomic
 * `batch`: style catalog + swatches, then the paragraph pour. For documents with
 * tables use {@link buildStory} (tables need mid-execution id resolution).
 */
export function buildDocumentMutations(ir: LoweredDoc, opts: { storyId: string }): Mutation {
  const paragraphs = ir.story.blocks.filter((b) => b.kind === "paragraph") as LoweredParagraph[];
  const ops: Mutation[] = [
    ...buildStyleMutations(ir),
    ...buildTextPour(paragraphs, opts.storyId, 0).mutations,
  ];
  return { op: "batch", args: { ops } } as Mutation;
}
