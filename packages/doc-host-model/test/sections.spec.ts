// ADR 029 — standalone open pours each Word section into its own story and
// save-back merges them back into one body.
import { describe, expect, it } from "vitest";

import { mergeSectionContents, pouredParagraphCount, sectionBlocks } from "../src/index.js";
import type { LoweredBlock, LoweredDoc, LoweredSection } from "../src/index.js";

const section = (firstBlock: number, story?: number): LoweredSection => ({
  pageWidthPt: 612,
  pageHeightPt: 792,
  marginTopPt: 72,
  marginBottomPt: 72,
  marginLeftPt: 72,
  marginRightPt: 72,
  columns: 1,
  firstBlock,
  ...(story === undefined ? {} : { story }),
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

  it("pours sections that share a story (continuous, ADR 029) as one group", () => {
    // Sections 0+1 share story 0 (section 1 is continuous on the same page),
    // section 2 opens story 1, sections 3+4 share story 2.
    const blocks = ["a", "b", "c", "d", "e", "f", "g", "h"].map(para);
    const groups = sectionBlocks(
      doc(blocks, [section(0, 0), section(2, 0), section(3, 1), section(5, 2), section(6, 2)]),
    );
    expect(groups.map((g) => g.length)).toEqual([3, 2, 3]);
    expect(groups.flat()).toEqual(blocks);
  });

  it("keeps every section its own story when the lowering has no story index", () => {
    const blocks = ["a", "b", "c"].map(para);
    expect(sectionBlocks(doc(blocks, [section(0), section(1), section(2)])).length).toBe(3);
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

  it("merges stories that each hold several Word sections back into one body", () => {
    const blocks = ["a", "b", "c", "d", "e"].map(para);
    const ir = doc(blocks, [section(0, 0), section(2, 0), section(3, 1)]);
    const counts = sectionBlocks(ir).map(pouredParagraphCount);
    expect(counts).toEqual([3, 2]);
    const merged = mergeSectionContents(
      [
        { selfId: "docx_s0", paragraphs: [p("a"), p("b"), p("c"), p("")] },
        { selfId: "docx_s1", paragraphs: [p("d"), p("e"), p("")] },
      ],
      counts,
    );
    expect(merged.paragraphs.map((x) => x.runs.map((r) => r.text).join(""))).toEqual([
      "a",
      "b",
      "c",
      "d",
      "e",
    ]);
  });

  it("keeps text a user added at a section's end", () => {
    const merged = mergeSectionContents([{ paragraphs: [p("x"), p("added")] }], [1]);
    expect(merged.paragraphs.length).toBe(2);
  });
});
