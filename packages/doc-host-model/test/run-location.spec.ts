/*
 * This file is part of paged (https://paged.media).
 *
 * paged is free software: you may redistribute it and/or modify it under the
 * terms of the GNU Affero General Public License, version 3, as published by
 * the Free Software Foundation, OR under the Paged Media Enterprise License
 * (PMEL), a commercial license available from And The Next GmbH. Full
 * copyright and license information is available in LICENSE.md, distributed
 * with this source code.
 *
 * paged is distributed in the hope that it will be useful, but WITHOUT ANY
 * WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the licenses for details.
 *
 *  @copyright  Copyright (c) And The Next GmbH
 *  @license    AGPL-3.0-only OR Paged Media Enterprise License (PMEL)
 */

// bodyRunLocation: which native paragraph a lowered body run landed in, per
// section story, and the text around it — the pour's own count, read back.

import { describe, expect, it } from "vitest";

import { bodyRunLocation, type LoweredDoc } from "../src";

const para = (...texts: string[]) => ({
  kind: "paragraph" as const,
  runs: texts.map((text) => ({ text })),
  sourceIndex: 0,
});

const section = (firstBlock: number, story?: number) => ({
  pageWidthPt: 612,
  pageHeightPt: 792,
  marginTopPt: 72,
  marginBottomPt: 72,
  marginLeftPt: 72,
  marginRightPt: 72,
  columns: 1,
  firstBlock,
  ...(story !== undefined ? { story } : {}),
});

describe("bodyRunLocation @feat:plugin-doc.object-model", () => {
  const ir = {
    swatches: [],
    styles: [],
    diagnostics: [],
    story: {
      blocks: [
        para("Dear ", "«Name»", ","),
        { kind: "paragraph", runs: [{ text: "split" }, { text: "here" }], sourceIndex: 1, segments: [{ at: 5 }] },
        { kind: "table", rows: 1, cols: 1, columnWidthsPt: [], cells: [] },
        para("Grüße ", "x"),
        para("second section"),
      ],
    },
    section: section(0),
    sections: [section(0, 0), section(4, 1)],
  } as unknown as LoweredDoc;

  it("counts native paragraphs: segments and table hosts each take one", () => {
    expect(bodyRunLocation(ir, 0, 1)).toEqual({
      story: 0,
      paragraph: 0,
      paragraphText: "Dear «Name»,",
      prefix: "Dear ",
      text: "«Name»",
    });
    // block 0 (1) + block 1 split once (2) + the table (1) = paragraph 4.
    expect(bodyRunLocation(ir, 3, 1)).toMatchObject({ story: 0, paragraph: 4, prefix: "Grüße ", text: "x" });
  });

  it("restarts in each section's own story", () => {
    expect(bodyRunLocation(ir, 4, 0)).toMatchObject({ story: 1, paragraph: 0, text: "second section" });
  });

  it("answers null for a table, a split paragraph, a missing block or run", () => {
    expect(bodyRunLocation(ir, 2, 0)).toBeNull();
    expect(bodyRunLocation(ir, 1, 0)).toBeNull();
    expect(bodyRunLocation(ir, 9, 0)).toBeNull();
    expect(bodyRunLocation(ir, 0, 3)).toBeNull();
  });
});
