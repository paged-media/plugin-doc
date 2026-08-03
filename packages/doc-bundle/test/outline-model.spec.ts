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

import type { LoweredDoc } from "@paged-media/doc-host-model";
import { describe, expect, it } from "vitest";

import {
  createDocStore,
  outlineEntries,
  sortedDiagnostics,
  summarize,
} from "../src/panels/outline-model.js";

function doc(overrides: Partial<LoweredDoc> = {}): LoweredDoc {
  return {
    swatches: [],
    styles: [
      {
        id: "ParagraphStyle/docx-Heading2",
        name: "docx-Heading2",
        collection: "paragraph",
        props: [],
      },
      {
        id: "ParagraphStyle/docx-p7",
        name: "heading 3", // detected by NAME, not id
        collection: "paragraph",
        props: [],
      },
      {
        id: "ParagraphStyle/docx-Default",
        name: "docx-Default",
        collection: "paragraph",
        props: [],
      },
    ],
    story: {
      blocks: [
        {
          kind: "paragraph",
          paraStyleId: "ParagraphStyle/docx-Heading2",
          runs: [{ text: "Chapter one" }],
          sourceIndex: 0,
        },
        {
          kind: "paragraph",
          paraStyleId: "ParagraphStyle/docx-Default",
          runs: [
            { text: "Body with a " },
            { text: "link", hyperlinkUrl: "https://example.com" },
          ],
          images: [{ widthPt: 10, heightPt: 10, uri: "data:image/png;base64,x" }],
          sourceIndex: 1,
        },
        {
          kind: "table",
          rows: 2,
          cols: 3,
          columnWidthsPt: [50, 50, 50],
          cells: [
            {
              row: 0,
              col: 0,
              rowSpan: 1,
              colSpan: 1,
              paragraphs: [
                {
                  paraStyleId: null,
                  runs: [{ text: "cell", hyperlinkUrl: "https://t.example" }],
                  sourceIndex: 2,
                },
              ],
            },
          ],
        },
        {
          kind: "paragraph",
          paraStyleId: "ParagraphStyle/docx-p7",
          runs: [{ text: "Named-style heading" }],
          sourceIndex: 3,
        },
      ],
    },
    section: {
      pageWidthPt: 595.3,
      pageHeightPt: 841.9,
      marginTopPt: 72,
      marginBottomPt: 72,
      marginLeftPt: 72,
      marginRightPt: 72,
      columns: 1,
    },
    diagnostics: [
      { severity: "info", tier: 3, message: "2 footnotes not inlined" },
      { severity: "error", tier: 0, message: "boom" },
      { severity: "warning", tier: 2, message: "header not placed" },
    ],
    ...overrides,
  };
}

describe("outlineEntries", () => {
  it("detects headings by style id AND by style name, tables in order", () => {
    const entries = outlineEntries(doc());
    expect(entries).toEqual([
      { kind: "heading", level: 2, text: "Chapter one", blockIndex: 0 },
      { kind: "table", rows: 2, cols: 3, blockIndex: 2 },
      { kind: "heading", level: 3, text: "Named-style heading", blockIndex: 3 },
    ]);
  });
});

describe("summarize", () => {
  it("counts paragraphs (incl. table cells), images, hyperlinks, headings", () => {
    const s = summarize(doc());
    // 3 body paragraphs + 1 cell paragraph.
    expect(s.paragraphs).toBe(4);
    expect(s.tables).toBe(1);
    expect(s.images).toBe(1);
    // one body hyperlink + one inside the table cell.
    expect(s.hyperlinks).toBe(2);
    expect(s.headings).toBe(2);
    expect(s.pageLine).toBe("595×842pt · 1 col");
  });
});

describe("sortedDiagnostics", () => {
  it("orders error → warning → info", () => {
    expect(sortedDiagnostics(doc()).map((d) => d.severity)).toEqual([
      "error",
      "warning",
      "info",
    ]);
  });
});

describe("createDocStore", () => {
  it("notifies subscribers on set and supports unsubscribe", () => {
    const store = createDocStore();
    let calls = 0;
    const off = store.subscribe(() => {
      calls += 1;
    });
    const d = { fileName: "a.docx", ir: doc(), frameId: null, storyId: null };
    store.set(d);
    expect(store.get()).toBe(d);
    expect(calls).toBe(1);
    off();
    store.set(null);
    expect(calls).toBe(1);
  });
});
