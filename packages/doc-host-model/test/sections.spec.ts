// ADR 029 — standalone open pours each Word section into its own story and
// save-back merges them back into one body.
import { describe, expect, it } from "vitest";

import { mergeSectionContents, sectionBlocks } from "../src/index.js";
import type { LoweredBlock, LoweredDoc, LoweredSection } from "../src/index.js";

const section = (firstBlock: number): LoweredSection => ({
  pageWidthPt: 612,
  pageHeightPt: 792,
  marginTopPt: 72,
  marginBottomPt: 72,
  marginLeftPt: 72,
  marginRightPt: 72,
  columns: 1,
  firstBlock,
});

const para = (text: string): LoweredBlock =>
  ({ kind: "paragraph", style: "p", runs: [{ text }] }) as unknown as LoweredBlock;

function doc(blocks: LoweredBlock[], sections?: LoweredSection[]): LoweredDoc {
  return {
    swatches: [],
    styles: [],
    story: { blocks },
    section: section(0),
    sections,
    diagnostics: [],
  } as LoweredDoc;
}

describe("sectionBlocks", () => {
  it("splits the story at each section's first block", () => {
    const blocks = ["a", "b", "c", "d", "e"].map(para);
    const groups = sectionBlocks(doc(blocks, [section(0), section(3)]));
    expect(groups.map((g) => g.length)).toEqual([3, 2]);
  });

  it("treats a lowering without sections as one section", () => {
    const blocks = ["a", "b"].map(para);
    expect(sectionBlocks(doc(blocks)).map((g) => g.length)).toEqual([2]);
  });
});

describe("mergeSectionContents", () => {
  const p = (...texts: string[]) => ({ runs: texts.map((text) => ({ text })) });

  it("concatenates sections in order and drops join artifacts", () => {
    const merged = mergeSectionContents(
      [
        { selfId: "s0", paragraphs: [p("one"), p("two"), p("")] },
        { selfId: "s1", paragraphs: [p("three"), p("")] },
      ],
      [2, 1],
    );
    expect(merged.paragraphs.map((x) => x.runs.map((r) => r.text).join(""))).toEqual([
      "one",
      "two",
      "three",
    ]);
    expect(merged.selfId).toBe("s0");
  });

  it("keeps an empty paragraph the section was poured with", () => {
    const merged = mergeSectionContents([{ paragraphs: [p("x"), p("")] }], [2]);
    expect(merged.paragraphs.length).toBe(2);
  });

  it("keeps text a user added at a section's end", () => {
    const merged = mergeSectionContents([{ paragraphs: [p("x"), p("added")] }], [1]);
    expect(merged.paragraphs.length).toBe(2);
  });
});
