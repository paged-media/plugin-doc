import { describe, expect, it } from "vitest";

import type { LoweredDoc } from "../src/lowered.js";
import {
  buildDocumentMutations,
  buildStory,
  buildStyleMutations,
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

  it("buildStory splits blocks into text/table/text steps", () => {
    const steps = buildStory(tableIr(), "Story/u1");
    expect(steps.map((s) => s.kind)).toEqual(["text", "table", "text"]);
    // The table step exposes insert + a cells(tableId) builder.
    const tableStep = steps[1] as { kind: "table"; insert: unknown; cells: (id: string) => unknown };
    expect((tableStep.insert as { op: string }).op).toBe("insertTable");
    expect((tableStep.cells("Table/u1") as { op: string }).op).toBe("batch");
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
