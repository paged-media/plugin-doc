import { describe, expect, it } from "vitest";

import type { LoweredDoc } from "../src/lowered.js";
import {
  buildDocumentMutations,
  buildPourMutations,
  buildStyleMutations,
} from "../src/mutations.js";

// Mirrors what docx-lower emits for a heading + a "plain / bold-red / plain"
// paragraph (the memo fixture).
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
      paragraphs: [
        { paraStyleId: "ParagraphStyle/docx-Heading1", runs: [{ text: "Title", charStyleId: null }], sourceIndex: 0 },
        {
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

describe("buildStyleMutations", () => {
  it("emits swatch, then create + setStyleProperty per style, parents first", () => {
    const ops = buildStyleMutations(memoIr());
    expect(ops[0]).toEqual({
      op: "createSwatch",
      args: { spec: { selfId: "Color/docx-FF0000", name: "docx FF0000", space: "RGB", value: [255, 0, 0] } },
    });
    // Normal is created before Heading1 (which is basedOn Normal).
    const createOps = ops.filter((o) => o.op === "createParagraphStyle");
    expect((createOps[0].args as { selfId: string }).selfId).toBe("ParagraphStyle/docx-Normal");
    expect((createOps[1].args as { selfId: string }).selfId).toBe("ParagraphStyle/docx-Heading1");
    // The synthesized character style is created with its two props.
    expect(ops).toContainEqual({
      op: "createCharacterStyle",
      args: { selfId: "CharacterStyle/docx-auto-c1", name: "docx direct format 1", basedOn: null },
    });
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

describe("buildPourMutations", () => {
  it("inserts the joined text once and styles the correct code-point ranges", () => {
    const ops = buildPourMutations(memoIr(), "Story/u1");
    const insert = ops.find((o) => o.op === "insertText");
    expect((insert?.args as { text: string }).text).toBe("Title\nMix bold");

    // Heading paragraph range [0,5]; second paragraph starts after "Title\n" = 6.
    expect(ops).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 0, end: 5, style: "ParagraphStyle/docx-Heading1", scope: "paragraph" },
    });
    // "bold" is at offset 6+4=10..14.
    expect(ops).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 10, end: 14, style: "CharacterStyle/docx-auto-c1", scope: "character" },
    });
  });

  it("counts code points, not UTF-16 units, for astral text", () => {
    const ir = memoIr();
    ir.story.paragraphs = [
      {
        paraStyleId: null,
        runs: [
          { text: "😀", charStyleId: null }, // 1 code point, 2 UTF-16 units
          { text: "x", charStyleId: "CharacterStyle/docx-auto-c1" },
        ],
        sourceIndex: 0,
      },
    ];
    const ops = buildPourMutations(ir, "Story/u1");
    // "x" must be at code-point offset 1..2, not 2..3.
    expect(ops).toContainEqual({
      op: "applyStyle",
      args: { storyId: "Story/u1", start: 1, end: 2, style: "CharacterStyle/docx-auto-c1", scope: "character" },
    });
  });
});

describe("buildDocumentMutations", () => {
  it("wraps everything in a single atomic batch", () => {
    const batch = buildDocumentMutations(memoIr(), { storyId: "Story/u1" });
    expect(batch.op).toBe("batch");
    const ops = (batch.args as { ops: unknown[] }).ops;
    expect(ops.length).toBeGreaterThan(5);
  });
});
