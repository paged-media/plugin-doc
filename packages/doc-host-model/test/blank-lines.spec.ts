// Word blank lines and plain line breaks (core 65cf615 / ab383b1).
//
// - A blank line (an empty Word paragraph) is styled with a CARET — an
//   `applyStyle` paragraph range with start === end, which names the empty
//   paragraph(s) at that contiguous offset. The story plan applies every
//   caret in a final step, once all paragraphs exist, so blank lines that
//   share an offset (consecutive ones, or ones on either side of a table)
//   still agree when the caret meets them.
// - A plain `<w:br/>` is U+2028 in the run text: one char in the contiguous
//   style space, three bytes in insertText's.
import { describe, expect, it } from "vitest";

import { buildStoryBlocks, buildTextPour } from "../src/index.js";
import type { LoweredBlock, LoweredParagraph } from "../src/index.js";

type Op = { op: string; args: Record<string, unknown> };

const para = (style: string, text: string): LoweredBlock => ({
  kind: "paragraph",
  paraStyleId: style,
  runs: text ? [{ text, charStyleId: "CharacterStyle/c" }] : [],
  sourceIndex: 0,
});
const table: LoweredBlock = { kind: "table", rows: 1, cols: 1, columnWidthsPt: [], cells: [] };

/** Run a story plan the way `pourSteps` does (table footprint 0 / 1). */
function run(blocks: LoweredBlock[]): Op[][] {
  let style = 0;
  let text = 0;
  return buildStoryBlocks(blocks, "s").map((step) => {
    if (step.kind === "table") {
      text += 1;
      return [step.insert as unknown as Op];
    }
    const ops = step.mutations(text, style) as unknown as Op[];
    style += step.length;
    text += step.byteLength;
    return ops;
  });
}

const paraStyles = (ops: Op[]) =>
  ops
    .filter((o) => o.op === "applyStyle" && o.args.scope === "paragraph")
    .map((o) => [o.args.start, o.args.end, o.args.style]);

describe("blank lines", () => {
  it("are styled by a caret, after the text, once per offset", () => {
    const steps = run([
      para("ParagraphStyle/a", "ab"),
      para("ParagraphStyle/blank", ""),
      para("ParagraphStyle/blank", ""),
      para("ParagraphStyle/b", "cd"),
    ]);
    expect(steps).toHaveLength(2);
    expect(steps[0][0].args.text).toBe("ab\n\n\ncd");
    // The text step carries only the non-empty paragraphs' ranges ...
    expect(paraStyles(steps[0])).toEqual([
      [0, 2, "ParagraphStyle/a"],
      [2, 4, "ParagraphStyle/b"],
    ]);
    // ... and the final step one caret for both blank lines at offset 2.
    expect(paraStyles(steps[1])).toEqual([[2, 2, "ParagraphStyle/blank"]]);
    // No character-scope op ever has an empty range.
    for (const o of steps.flat()) {
      if (o.op === "applyStyle" && o.args.scope === "character") {
        expect(o.args.end).toBeGreaterThan(o.args.start as number);
      }
    }
  });

  it("on either side of a table share an offset, and the last caret is the one sent", () => {
    const steps = run([
      para("ParagraphStyle/a", "ab"),
      para("ParagraphStyle/x", ""),
      table,
      para("ParagraphStyle/y", ""),
      para("ParagraphStyle/b", "cd"),
    ]);
    expect(steps.map((s) => s[0].op)).toEqual(["insertText", "insertTable", "insertText", "applyStyle"]);
    expect(paraStyles(steps[3])).toEqual([[2, 2, "ParagraphStyle/y"]]);
  });

  it("a story with no blank line has no caret step", () => {
    expect(run([para("ParagraphStyle/a", "ab")])).toHaveLength(1);
  });

  it("buildTextPour still inlines carets unless asked to defer them", () => {
    const paras = [para("ParagraphStyle/a", "ab"), para("ParagraphStyle/blank", "")] as LoweredParagraph[];
    const inline = buildTextPour(paras, "s", 0);
    expect(paraStyles(inline.mutations as unknown as Op[])).toContainEqual([2, 2, "ParagraphStyle/blank"]);
    const deferred = buildTextPour(paras, "s", 0, 0, { deferCarets: true });
    expect(paraStyles(deferred.mutations as unknown as Op[])).toEqual([[0, 2, "ParagraphStyle/a"]]);
    expect(deferred.carets).toHaveLength(1);
  });
});

describe("plain line breaks (U+2028)", () => {
  const br: LoweredParagraph = {
    paraStyleId: "ParagraphStyle/p",
    runs: [
      { text: "L02a\u2028L02b", charStyleId: "CharacterStyle/c" },
      { text: "\u2028\u2028x", charStyleId: "CharacterStyle/d" },
    ],
    sourceIndex: 0,
  };
  const next: LoweredParagraph = { paraStyleId: "ParagraphStyle/q", runs: [{ text: "n" }], sourceIndex: 1 };
  const out = buildTextPour([br, next], "s", 10, 4);
  const ops = out.mutations as unknown as Op[];

  it("stay in the inserted text, inside ONE paragraph", () => {
    expect(ops[0].args.text).toBe("L02a\u2028L02b\u2028\u2028x\nn");
    expect((ops[0].args.text as string).split("\n")).toHaveLength(2);
  });

  it("count one char in the style space and three bytes in the text space", () => {
    expect(paraStyles(ops)).toEqual([
      [4, 16, "ParagraphStyle/p"],
      [16, 17, "ParagraphStyle/q"],
    ]);
    const chr = ops.filter((o) => o.op === "applyStyle" && o.args.scope === "character").map((o) => [o.args.start, o.args.end]);
    expect(chr).toEqual([
      [4, 13],
      [13, 16],
    ]);
    expect(out.length).toBe(13);
    // 4 + 3 + 4 + 3 + 3 + 1 bytes, the separator, "n".
    expect(out.byteLength).toBe(18 + 1 + 1);
  });
});
