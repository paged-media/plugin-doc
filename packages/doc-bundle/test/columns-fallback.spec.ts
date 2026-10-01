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

// ADR 029 — a Word section that changes the columns mid-page lowers to
// span/split columns (protocol 64). The plugin API cannot say whether the
// engine has them, so the style batch is the probe: on a refusal the
// document is lowered with page breaks instead and reopened.
import { describe, expect, it } from "vitest";

import type { LoweredDoc } from "@paged-media/doc-host-model";

import { COLUMNS_DIAGNOSTICS_KEY, openStandalone } from "../src/open.js";

type Op = { op: string; args: Record<string, unknown> };

function doc(withSplit: boolean, sections: number): LoweredDoc {
  return {
    swatches: [],
    styles: [
      {
        id: "ParagraphStyle/docx-s1",
        name: "s1",
        collection: "paragraph",
        basedOn: null,
        props: withSplit
          ? [
              { path: "paragraphSpanColumnType", value: { type: "text", value: "SplitColumns" } },
              { path: "paragraphSpanSplitColumnCount", value: { type: "text", value: "2" } },
            ]
          : [{ path: "characterLeading", value: { type: "length", value: 12 } }],
      },
    ],
    story: { blocks: [] },
    section: {
      pageWidthPt: 360,
      pageHeightPt: 312,
      marginTopPt: 36,
      marginBottomPt: 36,
      marginLeftPt: 36,
      marginRightPt: 36,
      columns: 1,
    },
    sections: Array.from({ length: sections }, (_, k) => ({
      pageWidthPt: 360,
      pageHeightPt: 312,
      marginTopPt: 36,
      marginBottomPt: 36,
      marginLeftPt: 36,
      marginRightPt: 36,
      columns: 1,
      firstBlock: 0,
      story: k,
    })),
    diagnostics: [],
  };
}

/** An engine facade: span/split lowering until told otherwise. */
function fakeEngine() {
  let mid = true;
  const calls: string[] = [];
  const engine = {
    usesMidPageColumns: () => mid,
    setMidPageColumns(on: boolean) {
      calls.push(`setMidPageColumns(${on})`);
      mid = on;
    },
    lowered: () => doc(mid, mid ? 1 : 2),
    skeletonStories: () => (mid ? ["docx_s0"] : ["docx_s0", "docx_s1"]),
    skeletonIdml: (name: string) => new TextEncoder().encode(`${name}:${mid ? "split" : "pages"}`),
  };
  return { engine, calls };
}

/** A host whose engine knows span/split columns or not. */
function fakeHost(knowsColumns: boolean) {
  const opened: string[] = [];
  const applied: Op[] = [];
  const diagnostics = new Map<string, { severity: string; message: string }[]>();
  const refuses = (o: Op) =>
    !knowsColumns &&
    o.op === "setStyleProperty" &&
    String(o.args.path).startsWith("paragraphSpan");
  const host = {
    supports: (f: string) => f === "document.openNative@1",
    nativeDocument: {
      async open(bytes: Uint8Array) {
        opened.push(new TextDecoder().decode(bytes));
      },
    },
    document: {
      async mutate(m: Op) {
        const ops = m.op === "batch" ? (m.args.ops as Op[]) : [m];
        if (ops.some(refuses)) return { applied: false, error: { kind: "notImplemented" } };
        applied.push(...ops);
        return { applied: true, createdId: null, pageIds: [] };
      },
      async setMetadata() {},
    },
    parts: { async write() {} },
    diagnostics: {
      set: (k: string, v: { severity: string; message: string }[]) => diagnostics.set(k, v),
    },
    log: { debug() {}, info() {}, warn() {}, error() {} },
  };
  return { host, opened, applied, diagnostics };
}

describe("openStandalone and span/split columns", () => {
  it("keeps the span/split lowering on an engine that takes it", async () => {
    const { engine, calls } = fakeEngine();
    const { host, opened, applied, diagnostics } = fakeHost(true);
    const out = await openStandalone(host as never, engine as never, doc(true, 1), new Uint8Array(), "c.docx");
    expect(opened).toEqual(["c.docx:split"]);
    expect(calls).toEqual([]);
    expect(out?.midPageColumns).toBe(true);
    expect(out?.storyIds).toEqual(["docx_s0"]);
    expect(applied.some((o) => o.args.path === "paragraphSpanColumnType")).toBe(true);
    expect(diagnostics.get(COLUMNS_DIAGNOSTICS_KEY)).toEqual([]);
  });

  it("reopens with page breaks when the engine refuses them", async () => {
    const { engine, calls } = fakeEngine();
    const { host, opened, applied, diagnostics } = fakeHost(false);
    const out = await openStandalone(host as never, engine as never, doc(true, 1), new Uint8Array(), "c.docx");
    expect(opened).toEqual(["c.docx:split", "c.docx:pages"]);
    expect(calls).toEqual(["setMidPageColumns(false)"]);
    expect(out?.midPageColumns).toBe(false);
    expect(out?.storyIds).toEqual(["docx_s0", "docx_s1"]);
    expect(out?.ir.sections).toHaveLength(2);
    // The second lowering's styles went in (no span/split path anywhere).
    expect(applied.some((o) => o.args.path === "characterLeading")).toBe(true);
    expect(applied.some((o) => String(o.args.path).startsWith("paragraphSpan"))).toBe(false);
    // Both stories got their grow rule.
    expect(applied.filter((o) => o.op === "setFlowGrowRule").map((o) => o.args.storyId)).toEqual([
      "docx_s0",
      "docx_s1",
    ]);
    const d = diagnostics.get(COLUMNS_DIAGNOSTICS_KEY) ?? [];
    expect(d).toHaveLength(1);
    expect(d[0].severity).toBe("warning");
    expect(d[0].message).toContain("protocol 64");
  });

  it("a document without column changes is never reopened", async () => {
    const { engine, calls } = fakeEngine();
    engine.setMidPageColumns(true);
    calls.length = 0;
    const plain = { ...engine, usesMidPageColumns: () => false };
    const { host, opened } = fakeHost(false);
    const out = await openStandalone(host as never, plain as never, doc(false, 1), new Uint8Array(), "c.docx");
    expect(opened).toEqual(["c.docx:split"]);
    expect(out?.midPageColumns).toBe(false);
    expect(calls).toEqual([]);
  });
});
