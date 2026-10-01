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

// ADR 029 — the skeleton is IDML, which cannot carry a grow rule, so standalone
// open sets each section story's rule on the wire (protocol 64). An engine
// without the op keeps the document open and the bundle reports it.
import { describe, expect, it } from "vitest";

import { growRuleOp, OPEN_DIAGNOSTICS_KEY, setGrowRules } from "../src/open.js";

type Op = { op: string; args: Record<string, unknown> };

function fakeHost(knowsGrow: boolean) {
  const applied: Op[] = [];
  const calls: string[] = [];
  const diagnostics = new Map<string, { severity: string; message: string }[]>();
  const host = {
    document: {
      async mutate(m: Op) {
        calls.push(m.op);
        const ops = m.op === "batch" ? (m.args.ops as Op[]) : [m];
        if (!knowsGrow && ops.some((o) => o.op === "setFlowGrowRule")) {
          return { applied: false, error: { kind: "notImplemented" } };
        }
        applied.push(...ops);
        return { applied: true, createdId: null, pageIds: [] };
      },
    },
    diagnostics: { set: (k: string, v: { severity: string; message: string }[]) => diagnostics.set(k, v) },
  };
  return { host, applied, calls, diagnostics };
}

describe("setGrowRules", () => {
  it("speaks the protocol-64 wire shape (camelCase, Word's copyFrameOptions)", () => {
    expect(growRuleOp("docx_s1")).toEqual({
      op: "setFlowGrowRule",
      args: { storyId: "docx_s1", grow: true, maxPages: null, copyFrameOptions: true },
    });
  });

  it("sets every section story's rule in one batch", async () => {
    const { host, applied, calls, diagnostics } = fakeHost(true);
    expect(await setGrowRules(host as never, ["docx_s0", "docx_s1"])).toEqual([]);
    expect(calls).toEqual(["batch"]);
    expect(applied.map((o) => o.args.storyId)).toEqual(["docx_s0", "docx_s1"]);
    expect(diagnostics.get(OPEN_DIAGNOSTICS_KEY)).toEqual([]);
  });

  it("an engine without the op keeps the document and reports a warning", async () => {
    const { host, applied, diagnostics } = fakeHost(false);
    const refused = await setGrowRules(host as never, ["docx_s0", "docx_s1"]);
    expect(refused).toEqual(["docx_s0", "docx_s1"]);
    expect(applied).toEqual([]);
    const d = diagnostics.get(OPEN_DIAGNOSTICS_KEY) ?? [];
    expect(d).toHaveLength(1);
    expect(d[0].severity).toBe("warning");
    expect(d[0].message).toContain("2 of 2 Word section(s)");
  });

  it("no stories, no ops", async () => {
    const { host, calls } = fakeHost(true);
    expect(await setGrowRules(host as never, [])).toEqual([]);
    expect(calls).toEqual([]);
  });
});
