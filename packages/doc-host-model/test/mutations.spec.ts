import { describe, expect, it } from "vitest";

import type { LoweredDoc } from "../src/lowered.js";
import {
  buildDocumentMutations,
  buildStory,
  buildStyleMutations,
  buildStoryBlocks,
  buildTableCellCarets,
  buildTableCells,
  buildTableInsert,
  buildTextPour,
} from "../src/mutations.js";

// Mirrors what docx-lower emits for a heading + a "plain / bold-red / plain"
// paragraph (the memo fixture), as block-structured story content.
function memoIr(): LoweredDoc {
  return {
    swatches: [
      { id: "Color/docx-FF0000", name: "docx FF0000", space: "RGB", value: [255, 0, 0] },
    ],
    styles: [
      { id: "ParagraphStyle/docx-Normal", name: "Normal", collection: "paragraph", basedOn: null, props: [] },
      {
        id: "ParagraphStyle/docx-Heading1",
        name: "heading 1",
        collection: "paragraph",
        basedOn: "ParagraphStyle/docx-Normal",
        props: [{ path: "paragraphJustification", value: { type: "text", value: "CenterAlign" } }],
      },
      {
        id: "CharacterStyle/docx-auto-c1",
        name: "docx direct format 1",
        collection: "character",
        basedOn: null,
        props: [
          { path: "characterFontStyle", value: { type: "text", value: "Bold" } },
          { path: "characterFillColor", value: { type: "colorRef", value: "Color/docx-FF0000" } },
        ],
      },
    ],
    story: {
      blocks: [
        { kind: "paragraph", paraStyleId: "ParagraphStyle/docx-Heading1", runs: [{ text: "Title", charStyleId: null }], sourceIndex: 0 },
        {
          kind: "paragraph",
          paraStyleId: null,
          runs: [
            { text: "Mix ", charStyleId: null },
            { text: "bold", charStyleId: "CharacterStyle/docx-auto-c1" },
          ],
          sourceIndex: 1,
        },
      ],
    },
    section: {
      pageWidthPt: 595, pageHeightPt: 842, marginTopPt: 72, marginBottomPt: 72,
      marginLeftPt: 72, marginRightPt: 72, columns: 1,
    },
    diagnostics: [],
  };
}

const P = (paraStyleId: string | null, runs: { text: string; charStyleId: string | null }[]) =>
  ({ paraStyleId, runs, sourceIndex: 0 });

describe("buildStyleMutations", () => {
  it("emits swatch, then create + setStyleProperty per style, parents first", () => {
    const ops = buildStyleMutations(memoIr());
    expect(ops[0]).toEqual({
      op: "createSwatch",
      args: { spec: { selfId: "Color/docx-FF0000", name: "docx FF0000", space: "RGB", value: [255, 0, 0] } },
    });
    const createOps = ops.filter((o) => o.op === "createParagraphStyle");
    expect((createOps[0].args as { selfId: string }).selfId).toBe("ParagraphStyle/docx-Normal");
    expect((createOps[1].args as { selfId: string }).selfId).toBe("ParagraphStyle/docx-Heading1");
    expect(ops).toContainEqual({
      op: "setStyleProperty",
      args: {
        collection: "character",
        styleId: "CharacterStyle/docx-auto-c1",
        path: "characterFillColor",
        value: { type: "colorRef", value: "Color/docx-FF0000" },
      },
    });
  });
});

describe("buildTextPour", () => {
  it("inserts joined text at the base offset and styles code-point ranges", () => {
    const { mutations, length } = buildTextPour(
      [P("ParagraphStyle/docx-Heading1", [{ text: "Title", charStyleId: null }]),
       P(null, [{ text: "Mix ", charStyleId: null }, { text: "bold", charStyleId: "CharacterStyle/docx-auto-c1" }])],
      "Story/u1",
      0,
    );
    // Offsets are CONTIGUOUS — the engine consumes the paragraph-break `\n`, so it
    // does not occupy a char position. Inserted text keeps the `\n` (to create the
    // break); the returned length + style ranges do not count it.
    expect(length).toBe("TitleMix bold".length); // 13, not 14
    const insert = mutations.find((o) => o.op === "insertText");
    expect((insert?.args as { text: string }).text).toBe("Title\nMix bold");
    expect(mutations).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 0, end: 5, style: "ParagraphStyle/docx-Heading1", scope: "paragraph" },
    });
    // Contiguous: "Title"=[0,5), "Mix "=[5,9), "bold"=[9,13) — no +1 for the break.
    expect(mutations).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 9, end: 13, style: "CharacterStyle/docx-auto-c1", scope: "character" },
    });
  });

  it("rebases offsets by the base and counts code points", () => {
    const { mutations } = buildTextPour([P(null, [{ text: "😀", charStyleId: null }, { text: "x", charStyleId: "CharacterStyle/docx-auto-c1" }])], "Story/u1", 100);
    const insert = mutations.find((o) => o.op === "insertText");
    expect((insert?.args as { offset: number }).offset).toBe(100);
    // "x" is 1 code point after 😀, rebased by 100 -> 101..102.
    expect(mutations).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 101, end: 102, style: "CharacterStyle/docx-auto-c1", scope: "character" },
    });
  });
});

describe("tables", () => {
  function tableIr(): LoweredDoc {
    const ir = memoIr();
    ir.story.blocks = [
      { kind: "paragraph", paraStyleId: null, runs: [{ text: "Before", charStyleId: null }], sourceIndex: 0 },
      {
        kind: "table",
        rows: 2,
        cols: 2,
        columnWidthsPt: [100, 150],
        cells: [
          { row: 0, col: 0, rowSpan: 2, colSpan: 1, paragraphs: [P(null, [{ text: "Merged", charStyleId: null }])] },
          { row: 0, col: 1, rowSpan: 1, colSpan: 1, paragraphs: [P(null, [{ text: "Top", charStyleId: null }])] },
          { row: 1, col: 1, rowSpan: 1, colSpan: 1, paragraphs: [P(null, [{ text: "Bottom", charStyleId: null }])] },
        ],
      },
      { kind: "paragraph", paraStyleId: null, runs: [{ text: "After", charStyleId: null }], sourceIndex: 2 },
    ];
    return ir;
  }

  it("insertTable carries the grid + column widths", () => {
    const table = (tableIr().story.blocks[1] as unknown) as import("../src/lowered.js").LoweredTable;
    expect(buildTableInsert(table, "Story/u1")).toEqual({
      op: "insertTable",
      args: { storyId: "Story/u1", rows: 2, cols: 2, headerRows: 0, footerRows: 0, columnWidths: [100, 150], rowHeights: [] },
    });
  });

  it("cells pour by TextCellAddr and merged cells get setCellSpan", () => {
    const table = (tableIr().story.blocks[1] as unknown) as import("../src/lowered.js").LoweredTable;
    const batch = buildTableCells(table, "Story/u1", "Table/u1");
    const ops = (batch.args as { ops: Array<{ op: string; args: Record<string, unknown> }> }).ops;
    expect(ops).toContainEqual({
      op: "insertText",
      args: { storyId: "Story/u1", offset: 0, text: "Merged", cell: { tableId: "Table/u1", row: 0, col: 0 } },
    });
    expect(ops).toContainEqual({
      op: "setCellSpan",
      args: { storyId: "Story/u1", tableId: "Table/u1", row: 0, col: 0, rowSpan: 2, columnSpan: 1 },
    });
  });

  it("cell runs get a CELL-qualified applyStyle (v55)", () => {
    const ir = tableIr();
    const table = (ir.story.blocks[1] as unknown) as import("../src/lowered.js").LoweredTable;
    // Style the first cell's run so the builder has something to apply.
    table.cells[0].paragraphs[0].runs[0].charStyleId = "CharacterStyle/docx-auto-c1";
    const batch = buildTableCells(table, "Story/u1", "Table/u1");
    const ops = (batch.args as { ops: Array<{ op: string; args: Record<string, unknown> }> }).ops;
    const styled = ops.find((o) => o.op === "applyStyle");
    expect(styled?.args).toEqual({
      storyId: "Story/u1",
      start: 0,
      end: "Merged".length,
      style: "CharacterStyle/docx-auto-c1",
      scope: "character",
      cell: { tableId: "Table/u1", row: 0, col: 0 },
    });
  });

  it("buildStory splits blocks into text/table/text steps", () => {
    const steps = buildStory(tableIr(), "Story/u1");
    expect(steps.map((s) => s.kind)).toEqual(["text", "table", "text"]);
    // The table step exposes insert + a cells(tableId) builder.
    const tableStep = steps[1] as { kind: "table"; insert: unknown; cells: (id: string) => unknown };
    expect((tableStep.insert as { op: string }).op).toBe("insertTable");
    expect((tableStep.cells("Table/u1") as { op: string }).op).toBe("batch");
  });
});

describe("blank lines in table cells", () => {
  type Op = { op: string; args: Record<string, unknown> };
  const run = (text: string) => [{ text, charStyleId: null }];
  const blank = (style: string) => P(style, []);
  type TableBlock = Extract<import("../src/lowered.js").LoweredBlock, { kind: "table" }>;
  function cellsTable(): TableBlock {
    return {
      kind: "table",
      rows: 1,
      cols: 4,
      columnWidthsPt: [100, 100, 100, 100],
      cells: [
        // A, two blank lines (one cell offset, 1), B.
        {
          row: 0,
          col: 0,
          rowSpan: 1,
          colSpan: 1,
          paragraphs: [P("PS/a", run("A")), blank("PS/tall"), blank("PS/short"), P("PS/a", run("B"))],
        },
        // Ünï (3 code points), a blank line at 3, Dé, a trailing blank at 5.
        {
          row: 0,
          col: 1,
          rowSpan: 1,
          colSpan: 1,
          paragraphs: [P("PS/a", run("Ünï")), blank("PS/tall"), P("PS/a", run("Dé")), blank("PS/end")],
        },
        // An empty cell: insertTable mints it with NO paragraph, so the pour
        // seeds one with an empty insertText for its caret to name.
        { row: 0, col: 2, rowSpan: 1, colSpan: 1, paragraphs: [blank("PS/tall")] },
        // An empty cell with no style to give: nothing poured, no caret.
        { row: 0, col: 3, rowSpan: 1, colSpan: 1, paragraphs: [P(null, [])] },
      ],
    };
  }
  const cell = (col: number) => ({ tableId: "Table/t1", row: 0, col });
  const caret = (at: number, style: string, col: number) => ({
    op: "applyStyle",
    args: { storyId: "s", start: at, end: at, style, scope: "paragraph", cell: cell(col) },
  });

  it("are styled by a cell-addressed caret in the cell's own offsets, the last per offset", () => {
    expect(buildTableCellCarets(cellsTable(), "s", "Table/t1")).toEqual([
      caret(1, "PS/short", 0),
      caret(3, "PS/tall", 1),
      caret(5, "PS/end", 1),
      caret(0, "PS/tall", 2),
    ]);
  });

  it("are not in the cell batch, which keeps only the ranges over text", () => {
    const batch = buildTableCells(cellsTable(), "s", "Table/t1");
    const ops = (batch.args as { ops: Op[] }).ops;
    expect(ops.filter((o) => o.op === "applyStyle" && o.args.start === o.args.end)).toEqual([]);
    expect(ops.find((o) => o.op === "insertText" && (o.args.cell as { col: number }).col === 0)?.args.text).toBe(
      "A\n\n\nB",
    );
    // The empty styled cell gets an EMPTY insertText (its paragraph); the
    // unstyled one gets nothing.
    expect(ops.filter((o) => o.op === "insertText" && (o.args.cell as { col: number }).col === 2)).toEqual([
      { op: "insertText", args: { storyId: "s", offset: 0, text: "", cell: cell(2) } },
    ]);
    expect(ops.some((o) => (o.args.cell as { col: number } | undefined)?.col === 3)).toBe(false);
  });

  it("are applied in the story's final caret step, after the body's, once the table exists", () => {
    const steps = buildStoryBlocks(
      [
        { kind: "paragraph", paraStyleId: "PS/a", runs: run("x"), sourceIndex: 0 },
        { kind: "paragraph", paraStyleId: "PS/blank", runs: [], sourceIndex: 1 },
        cellsTable(),
      ],
      "s",
    );
    expect(steps.map((s) => s.kind)).toEqual(["text", "table", "text"]);
    // Pour the way pourSteps does: the text, the table (its id minted), then
    // the caret step.
    const text = steps[0] as Extract<(typeof steps)[number], { kind: "text" }>;
    text.mutations(0, 0);
    const table = steps[1] as Extract<(typeof steps)[number], { kind: "table" }>;
    table.cells("Table/t1");
    const last = steps[2] as Extract<(typeof steps)[number], { kind: "text" }>;
    expect(last.length).toBe(0);
    expect(last.mutations(2, 1)).toEqual([
      { op: "applyStyle", args: { storyId: "s", start: 1, end: 1, style: "PS/blank", scope: "paragraph" } },
      caret(1, "PS/short", 0),
      caret(3, "PS/tall", 1),
      caret(5, "PS/end", 1),
      caret(0, "PS/tall", 2),
    ]);
  });

  it("add a caret step to a story whose only blank lines are in cells", () => {
    const steps = buildStoryBlocks([cellsTable()], "s");
    expect(steps.map((s) => s.kind)).toEqual(["table", "text"]);
    (steps[0] as Extract<(typeof steps)[number], { kind: "table" }>).cells("Table/t9");
    const ops = (steps[1] as Extract<(typeof steps)[number], { kind: "text" }>).mutations(1, 0) as unknown as Op[];
    expect(ops.map((o) => (o.args.cell as { tableId: string }).tableId)).toEqual(["Table/t9", "Table/t9", "Table/t9", "Table/t9"]);
  });
});

describe("inline images", () => {
  it("emits insertAnchoredFrame at the paragraph offset with a data URI", () => {
    const { mutations } = buildTextPour(
      [
        P(null, [{ text: "Above", charStyleId: null }]),
        {
          paraStyleId: null,
          runs: [],
          images: [{ widthPt: 72, heightPt: 54, uri: "data:image/png;base64,AAAA" }],
          sourceIndex: 1,
        },
      ],
      "Story/u1",
      0,
    );
    // "Above" = 5 code points and the break is not a char position, so the image
    // paragraph anchors at contiguous offset 5.
    expect(mutations).toContainEqual({
      op: "insertAnchoredFrame",
      args: { storyId: "Story/u1", offset: 5, width: 72, height: 54, imageUri: "data:image/png;base64,AAAA" },
    });
  });
});

describe("inline images in their line", () => {
  it("places each picture at its own offset inside the paragraph", () => {
    const { mutations } = buildTextPour(
      [
        P(null, [{ text: "Above", charStyleId: null }]),
        {
          paraStyleId: null,
          runs: [
            { text: "añb", charStyleId: null },
            { text: "cd", charStyleId: null },
          ],
          images: [
            { at: 3, widthPt: 10, heightPt: 10, uri: "data:image/png;base64,A" },
            { at: 4, widthPt: 20, heightPt: 20, uri: "data:image/png;base64,B" },
          ],
          sourceIndex: 1,
        },
      ],
      "Story/u1",
      0,
      /* styleBase */ 7,
    );
    const offsets = mutations
      .filter((m) => (m as { op: string }).op === "insertAnchoredFrame")
      .map((m) => (m as unknown as { args: { offset: number } }).args.offset);
    // styleBase 7 + "Above" (5) = paragraph start 12; then its own `at`.
    expect(offsets).toEqual([15, 16]);
  });
});

describe("offset spaces", () => {
  it("separates the insertText base from the style-range base", () => {
    // After a table the two diverge: a table contributes 1 to insertText's
    // byte+break space but ZERO to the contiguous style space (its host
    // paragraph carries no runs). buildTextPour must honour both bases.
    const { mutations } = buildTextPour(
      [P(null, [{ text: "abc", charStyleId: "CharacterStyle/x" }])],
      "Story/u1",
      /* textBase */ 11,
      /* styleBase */ 10,
    );
    const insert = mutations.find((o) => o.op === "insertText");
    expect((insert?.args as { offset: number }).offset).toBe(11);
    expect(mutations).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 10, end: 13, style: "CharacterStyle/x", scope: "character" },
    });
  });

  it("reports UTF-8 byteLength (not code points) for the text space", () => {
    // "é" is 1 code point but 2 UTF-8 bytes; the paragraph separator counts too.
    const r = buildTextPour(
      [P(null, [{ text: "é", charStyleId: null }]), P(null, [{ text: "b", charStyleId: null }])],
      "Story/u1",
      0,
    );
    expect(r.length).toBe(2); // contiguous code points: é + b
    expect(r.byteLength).toBe(4); // "é" (2) + "\n" (1) + "b" (1)
  });
});

describe("hyperlinks", () => {
  it("emits insertHyperlink over the linked run's contiguous range", () => {
    const { mutations } = buildTextPour(
      [
        {
          paraStyleId: null,
          runs: [
            { text: "Visit ", charStyleId: null },
            { text: "Paged", charStyleId: "CharacterStyle/docx-link", hyperlinkUrl: "https://paged.media/" },
            { text: " today.", charStyleId: null },
          ],
          sourceIndex: 0,
        },
      ],
      "Story/u1",
      0,
    );
    // "Visit "=[0,6), "Paged"=[6,11) — the clickable link spans [6,11).
    expect(mutations).toContainEqual({
      op: "insertHyperlink",
      args: { storyId: "Story/u1", start: 6, end: 11, url: "https://paged.media/" },
    });
    // The blue+underline look still rides on the run's character style.
    expect(mutations).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 6, end: 11, style: "CharacterStyle/docx-link", scope: "character" },
    });
  });

  it("emits no insertHyperlink for ordinary runs", () => {
    const { mutations } = buildTextPour([P(null, [{ text: "plain", charStyleId: null }])], "Story/u1", 0);
    expect(mutations.some((o) => (o.op as string) === "insertHyperlink")).toBe(false);
  });
});

describe("buildDocumentMutations (text-only)", () => {
  it("wraps styles + paragraph pour in one atomic batch", () => {
    const batch = buildDocumentMutations(memoIr(), { storyId: "Story/u1" });
    expect(batch.op).toBe("batch");
    const ops = (batch.args as { ops: unknown[] }).ops;
    expect(ops.length).toBeGreaterThan(5);
  });
});

describe("buildTableCells — Word's cell geometry", () => {
  it("sends each cell's insets and vertical alignment", () => {
    const table = {
      rows: 1,
      cols: 1,
      columnWidthsPt: [100],
      cells: [
        {
          row: 0,
          col: 0,
          rowSpan: 1,
          colSpan: 1,
          paragraphs: [{ runs: [{ text: "x" }] }],
          insetsPt: [6, 5.4, 13.7, 3],
          vAlign: "CenterAlign",
        },
      ],
    } as unknown as import("../src/lowered.js").LoweredTable;
    const batch = buildTableCells(table, "s", "Table/t");
    const ops = (batch.args as { ops: Array<{ op: string; args: Record<string, unknown> }> }).ops;
    const elementId = { kind: "tableCell", id: { story_id: "s", table_id: "Table/t", row: 0, col: 0 } };
    const set = ops.filter((o) => o.op === "setElementProperty").map((o) => o.args);
    expect(set).toEqual([
      { elementId, path: "cellInsetTop", value: { type: "length", value: 6 } },
      { elementId, path: "cellInsetLeft", value: { type: "length", value: 5.4 } },
      { elementId, path: "cellInsetBottom", value: { type: "length", value: 13.7 } },
      { elementId, path: "cellInsetRight", value: { type: "length", value: 3 } },
      { elementId, path: "cellVerticalJustification", value: { type: "text", value: "CenterAlign" } },
    ]);
  });
});
