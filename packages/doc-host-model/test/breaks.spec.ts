// ADR 028/029 — a page/column break INSIDE a Word paragraph pours as a new
// native paragraph at the break, styled with the segment's style (which
// carries the break-before rule), while the contiguous style offsets ignore
// the separator like every paragraph separator.
import { describe, expect, it } from "vitest";

import { buildTextPour, pouredParagraphCount } from "../src/index.js";
import type { LoweredBlock, LoweredParagraph } from "../src/index.js";

type Op = { op: string; args: Record<string, unknown> };

const split: LoweredParagraph = {
  paraStyleId: "ParagraphStyle/docx-auto-p2",
  runs: [{ text: "A21a beforeA21b after", charStyleId: "CharacterStyle/c" }],
  sourceIndex: 20,
  segments: [{ at: 11, paraStyleId: "ParagraphStyle/docx-auto-p1" }],
};
const plain: LoweredParagraph = { paraStyleId: "ParagraphStyle/p", runs: [{ text: "A22" }], sourceIndex: 21 };

describe("a paragraph split by a break", () => {
  const { mutations, length } = buildTextPour([split, plain], "docx_s0", 0);
  const ops = mutations as unknown as Op[];

  it("inserts a paragraph separator at the break", () => {
    expect(ops[0].op).toBe("insertText");
    expect(ops[0].args.text).toBe("A21a before\nA21b after\nA22");
  });

  it("styles each part with its own paragraph style, offsets contiguous", () => {
    const para = ops.filter((o) => o.op === "applyStyle" && o.args.scope === "paragraph").map((o) => o.args);
    expect(para.map((a) => [a.start, a.end, a.style])).toEqual([
      [0, 11, "ParagraphStyle/docx-auto-p2"],
      [11, 21, "ParagraphStyle/docx-auto-p1"],
      [21, 24, "ParagraphStyle/p"],
    ]);
    // The run's character style spans both parts, contiguously.
    const chr = ops.filter((o) => o.op === "applyStyle" && o.args.scope === "character").map((o) => o.args);
    expect(chr.map((a) => [a.start, a.end])).toEqual([[0, 21]]);
    expect(length).toBe(24);
  });

  it("a break at the paragraph's end pours an empty trailing part", () => {
    const tail: LoweredParagraph = {
      paraStyleId: "ParagraphStyle/a",
      runs: [{ text: "end" }],
      sourceIndex: 0,
      segments: [{ at: 3, paraStyleId: "ParagraphStyle/b" }],
    };
    const out = buildTextPour([tail], "s", 0).mutations as unknown as Op[];
    expect(out[0].args.text).toBe("end\n");
  });

  it("counts the native paragraphs a section pours as", () => {
    const blocks = [
      { kind: "paragraph", ...split },
      { kind: "paragraph", ...plain },
      { kind: "table", rows: 1, cols: 1, columnWidthsPt: [], cells: [] },
    ] as LoweredBlock[];
    expect(pouredParagraphCount(blocks)).toBe(3);
  });
});
