// paged.doc — pure Lowered IR -> host `Mutation[]`.
//
// The same role as `sheet-host-model`: a *dumb* translator. Every id in the IR
// is already a fully-formed Paged token, so this file never invents ids — it only
// maps IR nodes to `host.document.mutate(...)` ops. No engine logic lives here;
// the semantics were decided in `docx-lower` (Rust).

import type { Mutation } from "@paged-media/plugin-api";

import type {
  LoweredBlock,
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
  /** The native paragraph a CARET range names (core wire v65, RFI C-53):
   *  its 0-based index in the story. Only set on carets, and only when the
   *  pour knows where in the story it starts. */
  paragraph?: number;
}

/** An inline image + its story offset (contiguous char space). */
interface ImageAt {
  offset: number;
  /** The native paragraph holding it (wire v65), when known. */
  paragraph?: number;
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
  /** Index of the native paragraph the text starts in; `undefined` when the
   *  pour does not know (no paragraph addresses are emitted then). */
  paraBase?: number,
): { text: string; ranges: Range[]; images: ImageAt[]; links: LinkAt[]; length: number; breaks: number } {
  const ranges: Range[] = [];
  const images: ImageAt[] = [];
  const links: LinkAt[] = [];
  let text = "";
  let offset = base;
  // Separators emitted so far: each starts a new native paragraph.
  let breaks = 0;
  const addressed = (r: Range, index: number): Range =>
    paraBase !== undefined && r.start === r.end ? { ...r, paragraph: paraBase + index } : r;
  paragraphs.forEach((para, pIdx) => {
    const paraStart = offset;
    const firstIndex = breaks;
    // ADR 028/029 — a break inside the Word paragraph: a separator at each
    // segment start (a new native paragraph), which, like every paragraph
    // separator, does not advance the contiguous offsets.
    const cuts = (para.segments ?? []).map((s) => s.at);
    let local = 0;
    for (const run of para.runs) {
      const runStart = offset;
      const chars = Array.from(run.text);
      chars.forEach((ch, i) => {
        if (cuts.includes(local + i) && local + i > 0) {
          text += "\n";
          breaks += 1;
        }
        text += ch;
      });
      local += chars.length;
      offset += chars.length;
      if (run.charStyleId) {
        ranges.push({ start: runStart, end: offset, style: run.charStyleId, scope: "character" });
      }
      if (run.hyperlinkUrl && offset > runStart) {
        links.push({ start: runStart, end: offset, url: run.hyperlinkUrl });
      }
    }
    // Breaks at the paragraph's very end (an empty trailing part).
    for (const at of cuts) {
      if (at >= local) {
        text += "\n";
        breaks += 1;
      }
    }
    const segs = para.segments ?? [];
    // A cut starts a new native paragraph unless it sits at the very start
    // of a paragraph that has text (see the two separator rules above).
    const splits = (at: number) => (at > 0 && at < local) || at >= local;
    // The native paragraph (relative to this pour) segment `k` is.
    const segIndex = (k: number) => firstIndex + segs.slice(0, k + 1).filter((s) => splits(s.at)).length;
    const firstEnd = segs.length > 0 ? paraStart + segs[0].at : offset;
    if (para.paraStyleId) {
      ranges.push(
        addressed({ start: paraStart, end: firstEnd, style: para.paraStyleId, scope: "paragraph" }, firstIndex),
      );
    }
    segs.forEach((seg, k) => {
      const start = paraStart + seg.at;
      const end = k + 1 < segs.length ? paraStart + segs[k + 1].at : offset;
      if (seg.paraStyleId) {
        ranges.push(addressed({ start, end, style: seg.paraStyleId, scope: "paragraph" }, segIndex(k)));
      }
    });
    // An inline picture is a character of its line (core 17d3d3d places the
    // frame AT its offset): where the lowering says it sits in the paragraph.
    for (const img of para.images ?? []) {
      const at = img.at ?? 0;
      // The segment the picture stands in: the last one starting at or
      // before it.
      const inSeg = segs.filter((s) => splits(s.at) && s.at <= at).length;
      images.push({
        offset: paraStart + at,
        ...(paraBase !== undefined ? { paragraph: paraBase + firstIndex + inSeg } : {}),
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        uri: img.uri,
      });
    }
    // Separator text for insertText, but NOT an offset advance (contiguous).
    if (pIdx < paragraphs.length - 1) {
      text += "\n";
      breaks += 1;
    }
  });
  return { text, ranges, images, links, length: offset - base, breaks };
}

/** A zero-length paragraph-style range: a CARET, which styles the empty
 *  paragraph(s) at its offset (core `65cf615`). A Word blank line pours as an
 *  empty native paragraph, which occupies no characters in the contiguous
 *  space, so a caret is the only range that can name it. */
function isCaret(r: Range): boolean {
  return r.scope === "paragraph" && r.start === r.end;
}

/** insertText + applyStyle + insertAnchoredFrame (inline images) for a
 *  contiguous paragraph run, offsets from `base`.
 *
 *  Blank lines (empty paragraphs) are styled with a caret `applyStyle`
 *  (`start === end`). With `deferCarets`, those are left out of `mutations`
 *  and returned in `carets` instead, for the caller to apply after the whole
 *  story is poured (see {@link buildStoryBlocks}). */
export function buildTextPour(
  paragraphs: LoweredParagraph[],
  storyId: string,
  /** Where the `insertText` lands — the engine's byte + synthetic-break space. */
  textBase: number,
  /** Where style/anchor/link RANGES start — the contiguous character space.
   *  Defaults to `textBase` (they coincide only before the first table). */
  styleBase: number = textBase,
  opts: {
    deferCarets?: boolean;
    /** Index of the native paragraph the text starts in. With it, blank
     *  lines and pictures name their paragraph (core wire v65, RFI C-53);
     *  an engine without that address ignores the field. */
    paraBase?: number;
  } = {},
): { mutations: Mutation[]; length: number; byteLength: number; carets: Mutation[]; breaks: number } {
  const { text, ranges, images, links, length, breaks } = poured(paragraphs, styleBase, opts.paraBase);
  const ops: Mutation[] = [];
  const carets: Mutation[] = [];
  if (text.length > 0) {
    ops.push({
      op: "insertText",
      args: { storyId, offset: textBase, text, cell: null },
    } as Mutation);
  }
  for (const r of ranges.filter((r) => r.scope === "paragraph")) {
    const op = applyStyleOp(storyId, r.start, r.end, r.style, "paragraph", r.paragraph);
    if (isCaret(r)) carets.push(op);
    if (!(isCaret(r) && opts.deferCarets)) ops.push(op);
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
        ...(img.paragraph !== undefined ? { paragraph: img.paragraph } : {}),
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
  // The engine's insertText space counts BYTES, so report UTF-8 length (not code
  // points) for the caller's running text offset.
  // (U+2028, Word's line break, is 3 bytes here and 1 char in `length`.)
  const byteLength = new TextEncoder().encode(text).length;
  return { mutations: ops, length, byteLength, carets, breaks };
}

function applyStyleOp(
  storyId: string,
  start: number,
  end: number,
  style: string,
  scope: "paragraph" | "character",
  /** v65 paragraph address (see `Range.paragraph`). */
  paragraph?: number,
): Mutation {
  const args = { storyId, start, end, style, scope, ...(paragraph !== undefined ? { paragraph } : {}) };
  return { op: "applyStyle", args } as Mutation;
}

/** A CELL-qualified `applyStyle` (core protocol v55). The `cell` arg postdates
 *  the published plugin-api Mutation union, so cast via `unknown` — the same
 *  pattern the v52/v53 ops use. A host below v55 rejects the op and the cell text
 *  simply keeps its default formatting (honest degrade, no wrong styling). */
function applyStyleInCell(
  storyId: string,
  start: number,
  end: number,
  style: string,
  scope: "paragraph" | "character",
  cell: { tableId: string; row: number; col: number },
): Mutation {
  return {
    op: "applyStyle",
    args: { storyId, start, end, style, scope, cell },
  } as unknown as Mutation;
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
      headerRows: table.headerRows ?? 0,
      footerRows: 0,
      columnWidths: table.columnWidthsPt,
      // A row is never shorter than Word's `w:trHeight`.
      rowHeights: table.rowHeightsPt ?? [],
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

/** A cell whose Word content is one EMPTY paragraph that carries a paragraph
 *  style. `insertTable` mints its cells with NO paragraph at all (core
 *  `new_table_cell`), so a caret there is refused ("caret offset 0 addresses
 *  no paragraph (story length 0)"). An empty `insertText` seeds the cell's
 *  one empty paragraph (core `apply_insert_text`), which the caret can then
 *  style. */
function isStyledEmptyCell(cell: LoweredCell): boolean {
  return cellText(cell).length === 0 && cell.paragraphs.some((p) => p.paraStyleId);
}

/** The cell-pour + merge batch for a resolved `tableId`: `insertText` per cell
 *  (addressed by TextCellAddr) + `setCellSpan` per merged cell. A styled cell
 *  with no text gets an EMPTY `insertText`, so it has the paragraph its caret
 *  styles ({@link buildTableCellCarets}). */
export function buildTableCells(table: LoweredTable, storyId: string, tableId: string): Mutation {
  const ops: Mutation[] = [];
  for (const cell of table.cells) {
    const text = cellText(cell);
    if (isStyledEmptyCell(cell)) {
      ops.push({
        op: "insertText",
        args: { storyId, offset: 0, text: "", cell: { tableId, row: cell.row, col: cell.col } },
      } as Mutation);
    }
    if (text.length > 0) {
      const addr = { tableId, row: cell.row, col: cell.col };
      ops.push({
        op: "insertText",
        args: { storyId, offset: 0, text, cell: addr },
      } as Mutation);
      // v55 — style the cell's runs. `applyStyle` gained a cell qualifier, so
      // cell text no longer has to pour at the default formatting. Offsets are
      // CELL-LOCAL and contiguous, computed over the same joined text
      // `cellText` produced (paragraphs joined by the consumed `\n`).
      let offset = 0;
      cell.paragraphs.forEach((para, pIdx) => {
        const paraStart = offset;
        for (const run of para.runs) {
          const runStart = offset;
          offset += codePointLen(run.text);
          if (run.charStyleId) {
            ops.push(applyStyleInCell(storyId, runStart, offset, run.charStyleId, "character", addr));
          }
        }
        if (para.paraStyleId && offset > paraStart) {
          ops.push(applyStyleInCell(storyId, paraStart, offset, para.paraStyleId, "paragraph", addr));
        }
        void pIdx;
      });
    }
    // Word's cell geometry: margins, the first paragraph's space before and
    // what Word's line box has beyond the native cell's (docx-lower
    // `cell_insets`). Without them a native cell is the engine's zero-inset
    // default and every row comes out shorter than Word's.
    const elementId = {
      kind: "tableCell",
      id: { story_id: storyId, table_id: tableId, row: cell.row, col: cell.col },
    };
    if (cell.insetsPt) {
      const paths = ["cellInsetTop", "cellInsetLeft", "cellInsetBottom", "cellInsetRight"];
      cell.insetsPt.forEach((pt, i) => {
        ops.push({
          op: "setElementProperty",
          args: { elementId, path: paths[i], value: { type: "length", value: pt } },
        } as Mutation);
      });
    }
    if (cell.vAlign) {
      ops.push({
        op: "setElementProperty",
        args: {
          elementId,
          path: "cellVerticalJustification",
          value: { type: "text", value: cell.vAlign },
        },
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

/** The carets that style a table's BLANK cell lines (empty Word paragraphs
 *  inside a cell), in the cell's OWN contiguous offset space (`insertText`
 *  with a `cell` address restarts at 0, and so do the cell's style ranges).
 *  A blank line occupies no characters, so it sits at the offset its next
 *  paragraph starts at, and consecutive ones share it: only the LAST caret
 *  per cell offset is emitted (the body's rule). A cell of one empty
 *  paragraph is addressed too: {@link buildTableCells} seeds its paragraph
 *  with an empty `insertText`, and its caret at 0 styles it.
 *  Applied after every cell is poured (see {@link buildStoryBlocks}), so an
 *  engine that refuses a caret costs no cell its text. */
export function buildTableCellCarets(table: LoweredTable, storyId: string, tableId: string): Mutation[] {
  const ops: Mutation[] = [];
  for (const cell of table.cells) {
    if (cellText(cell).length === 0 && !isStyledEmptyCell(cell)) continue;
    const addr = { tableId, row: cell.row, col: cell.col };
    const last = new Map<number, string>();
    let offset = 0;
    for (const para of cell.paragraphs) {
      const paraStart = offset;
      for (const run of para.runs) offset += codePointLen(run.text);
      if (para.paraStyleId && offset === paraStart) last.set(paraStart, para.paraStyleId);
    }
    for (const [at, style] of last) {
      ops.push(applyStyleInCell(storyId, at, at, style, "paragraph", addr));
    }
  }
  return ops;
}

// ---------------------------------------------------------------------------
// The story plan (block-aware; tables need mid-execution tableId resolution)

/** One step the bundle executes in order against a resolved `storyId`. A text
 *  step builds its ops given the running story offset (advancing it by `length`);
 *  a table step inserts the table (its outcome mints the id), then pours cells. */
export type StoryStep =
  | {
      kind: "text";
      /** Contiguous-character length (the style-range space). */
      length: number;
      /** UTF-8 byte length of the inserted text (the insertText space). */
      byteLength: number;
      mutations: (textBase: number, styleBase: number) => Mutation[];
    }
  | { kind: "table"; insert: Mutation; cells: (tableId: string) => Mutation };

/** Split the story's blocks into executable steps: consecutive paragraph blocks
 *  coalesce into one text step; each table becomes a table step. */
export function buildStory(ir: LoweredDoc, storyId: string): StoryStep[] {
  return buildStoryBlocks(ir.story.blocks, storyId);
}

/** ADR 029 — the story's blocks split per native STORY of the standalone
 *  skeleton, in order: one group per run of consecutive sections sharing a
 *  `story` (a continuous section Word keeps on the same page joins the story
 *  before it), cut at each story's first section's `firstBlock` (blocks map
 *  1:1 to Word body blocks). Index k pours into the skeleton's k-th story.
 *  A section without `story` is its own story; a lowering without `sections`
 *  is one story holding every block. */
export function sectionBlocks(ir: LoweredDoc): LoweredBlock[][] {
  const blocks = ir.story.blocks;
  const sections = ir.sections && ir.sections.length > 0 ? ir.sections : [ir.section];
  const starts: number[] = [];
  sections.forEach((s, k) => {
    const joins = k > 0 && s.story !== undefined && s.story === sections[k - 1].story;
    if (!joins) starts.push(s.firstBlock ?? 0);
  });
  return starts.map((start, k) => blocks.slice(start, k + 1 < starts.length ? starts[k + 1] : blocks.length));
}

/** How many native paragraphs a block list pours as: one per paragraph block,
 *  plus one per break that splits a paragraph (ADR 028/029). Save-back trims
 *  the section joins against this count. */
export function pouredParagraphCount(blocks: readonly LoweredBlock[]): number {
  const one = (p: LoweredParagraph) => 1 + (p.segments?.length ?? 0);
  return blocks.reduce(
    (n, b) => n + (b.kind === "paragraph" ? one(b) : (b.flow ?? []).reduce((k, p) => k + one(p), 0)),
    0,
  );
}

/** {@link buildStory} over an explicit block list (one section's blocks).
 *
 *  Blank lines are styled LAST, in a final step, once every paragraph of the
 *  story exists, one caret per blank line. A blank line has no characters,
 *  so consecutive ones share one offset (so do blank lines on either side of
 *  a table: its host paragraph has no characters either); each caret
 *  therefore also names its paragraph by index (core wire v65, RFI C-53).
 *
 *  An engine without that address styles EVERY empty paragraph at the
 *  caret's offset and refuses a caret over empty paragraphs whose styles
 *  differ. Deferred to the end, every caret there meets paragraphs that
 *  agree (all fresh, or all styled by the caret before it), so none is
 *  refused and the last caret of a group wins for all of them, which
 *  `docx-lower` reports as a warning. Blank lines inside table cells are
 *  styled in the same step, by cell-addressed carets
 *  ({@link buildTableCellCarets}). */
export function buildStoryBlocks(blocks: readonly LoweredBlock[], storyId: string): StoryStep[] {
  const steps: StoryStep[] = [];
  let pending: LoweredParagraph[] = [];
  // Each text step's carets, filled in when the pour runs it (the style base
  // is only known then).
  const caretsByStep: Mutation[][] = [];
  // Each table's cell carets, filled in once its id is minted.
  const cellCarets: Mutation[][] = [];
  let hasCarets = false;
  let nextParagraph = 0;
  const flush = () => {
    if (pending.length === 0) return;
    const paras = pending;
    pending = [];
    const probe = buildTextPour(paras, storyId, 0);
    const slot = caretsByStep.length;
    caretsByStep.push([]);
    // The story is fresh: its text starts in paragraph 0, each separator
    // starts another, and a table takes one (the text after it continues
    // in the table's own host paragraph).
    const paraBase = nextParagraph;
    nextParagraph += probe.breaks;
    steps.push({
      kind: "text",
      length: probe.length,
      byteLength: probe.byteLength,
      mutations: (textBase, styleBase) => {
        const out = buildTextPour(paras, storyId, textBase, styleBase, { deferCarets: true, paraBase });
        caretsByStep[slot] = out.carets;
        return out.mutations;
      },
    });
    if (probe.carets.length > 0) hasCarets = true;
  };
  for (const block of blocks) {
    if (block.kind === "table" && block.flow && block.flow.length > 0) {
      // A table with a row taller than its page (docx-lower): its text
      // flows as paragraphs, joining the text around it.
      pending.push(...block.flow);
      continue;
    }
    if (block.kind === "table") {
      flush();
      nextParagraph += 1;
      const table: LoweredTable = block;
      const slot = cellCarets.length;
      cellCarets.push([]);
      if (buildTableCellCarets(table, storyId, "").length > 0) hasCarets = true;
      steps.push({
        kind: "table",
        insert: buildTableInsert(table, storyId),
        cells: (tableId) => {
          cellCarets[slot] = buildTableCellCarets(table, storyId, tableId);
          return buildTableCells(table, storyId, tableId);
        },
      });
    } else {
      pending.push(block);
    }
  }
  flush();
  if (hasCarets) {
    steps.push({
      kind: "text",
      length: 0,
      byteLength: 0,
      mutations: () => {
        // The body's carets in order, then each poured table's (one per
        // cell offset; a table whose insert was refused has none). Each
        // names its own paragraph (wire v65). An engine without that
        // address styles every blank line at the caret's offset instead,
        // so there the last caret of a group wins for all of them.
        return [...caretsByStep.flat(), ...cellCarets.flat()];
      },
    });
  }
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

/** The read-back shape of one story (`host.document.storyContent`), kept
 *  structural so this package needs no newer plugin-api than it builds on. */
export interface StoryContentLike {
  selfId?: string;
  paragraphs: Array<{ runs: Array<{ text: string }> } & Record<string, unknown>>;
}

/**
 * ADR 029 — save-back reads a standalone document back SECTION BY SECTION
 * (one story each) and needs ONE body, as the import baseline is one body.
 * Concatenates the stories in section order. A story can end in an empty
 * paragraph the skeleton left (its single empty paragraph, which the pour
 * extends): trailing EMPTY paragraphs beyond the paragraph count the section
 * was poured with are dropped, so the joins add no paragraphs the Word
 * document never had. `pouredParagraphs[k]` is that count for section k.
 */
export function mergeSectionContents(
  contents: readonly StoryContentLike[],
  pouredParagraphs: readonly number[],
): StoryContentLike {
  const isEmpty = (p: { runs: Array<{ text: string }> }) => p.runs.every((r) => r.text === "");
  const paragraphs: StoryContentLike["paragraphs"] = [];
  contents.forEach((c, k) => {
    const ps = [...c.paragraphs];
    const expected = pouredParagraphs[k] ?? ps.length;
    while (ps.length > expected && isEmpty(ps[ps.length - 1])) ps.pop();
    paragraphs.push(...ps);
  });
  return { selfId: contents[0]?.selfId, paragraphs };
}
