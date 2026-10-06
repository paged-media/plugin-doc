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

// Count budget for the object model (ADR 323): a 10-op batch of run writes
// on one Word document is ONE engine commit (one undo step), ONE save-back
// (two engine boots: patch, then lower the result) and a constant number of
// host reads — one tree read, one story read, one part read — never one per
// op. Measured 2026-10-06; a pin is lowered by the commit that earns it,
// never raised.

import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";

import type { BundleHost, MutationInput } from "@paged-media/plugin-api";
import { createObjectRegistry, type CoreObjects, type HeadlessHost } from "@paged-media/plugin-sdk";

import { DocEngine } from "../../src/engine";
import { createDocObjectModel, docAddress, PLUGIN_ID } from "../../src/object-model";
import { docxPresent, openDoc, requireDocx, storyTexts } from "./harness";

requireDocx("doc object model count budget");

/** Count every call on the host's `document` and `parts` doors. */
function counting(host: BundleHost): { host: BundleHost; calls: Record<string, number> } {
  const calls: Record<string, number> = {};
  const wrap = <T extends object>(target: T, prefix: string): T =>
    new Proxy(target, {
      get(t, key, receiver) {
        const v = Reflect.get(t, key, receiver) as unknown;
        if (typeof v !== "function") return v;
        return (...args: unknown[]) => {
          const name = `${prefix}.${String(key)}`;
          calls[name] = (calls[name] ?? 0) + 1;
          return (v as (...a: unknown[]) => unknown).apply(t, args);
        };
      },
    });
  const proxied = new Proxy(host, {
    get(t, key, receiver) {
      if (key === "document") return wrap(t.document, "document");
      if (key === "parts") return wrap(t.parts, "parts");
      return Reflect.get(t, key, receiver) as unknown;
    },
  });
  return { host: proxied, calls };
}

describe.skipIf(!docxPresent)("doc object model count budget @feat:plugin-doc.object-model", () => {
  let h: HeadlessHost;
  let story: string;

  beforeAll(async () => {
    const o = await openDoc("merge-letter.docx");
    h = o.h;
    story = o.opened.storyIds[0]!;
  });
  afterAll(() => h?.dispose());

  it("10 run writes: 1 commit, 2 engine boots, 1 tree + 1 story + 1 part read", async () => {
    const core = h.objects.core!;
    const commits: MutationInput[][] = [];
    const countedCore: CoreObjects = new Proxy(core, {
      get(target, key, receiver) {
        if (key === "commit") {
          return (m: readonly MutationInput[]) => {
            commits.push([...m]);
            return target.commit(m);
          };
        }
        const v = Reflect.get(target, key, receiver) as unknown;
        return typeof v === "function" ? (v as (...a: unknown[]) => unknown).bind(target) : v;
      },
    });
    const bundleHost = h.host as unknown as BundleHost;
    const { host, calls } = counting(bundleHost);
    const registry = createObjectRegistry({ core: countedCore });
    registry.register({
      plugin: PLUGIN_ID,
      model: createDocObjectModel(host, { place: async () => null, exportBytes: async () => null }),
      parts: bundleHost.parts,
    });
    // Warm the lowering memo (a read) so the budget counts the write alone.
    await registry.get(docAddress("run", "docx_f0/b0/r0"), "text");
    for (const k of Object.keys(calls)) delete calls[k];
    const boots = vi.spyOn(DocEngine, "boot");

    const ops = [0, 1, 2, 3, 4, 5, 6, 7, 8].map((j) => ({
      op: "set" as const,
      address: docAddress("run", `docx_f0/b0/r${j}`),
      path: "text",
      value: `R${j}`,
    }));
    ops.push({ op: "set", address: docAddress("run", "docx_f0/b1/r0"), path: "text", value: "S" });
    const out = await registry.batch(ops);
    expect(out).toMatchObject({ applied: true, undoSteps: 1 });
    expect(await storyTexts(h.host as unknown as BundleHost, story)).toEqual(["R0R1R2R3R4R5R6R7R8", "S«OrderNo»"]);

    if (process.env.PERF_SHOW) console.log("PERF objects.batch(10)", JSON.stringify(calls), boots.mock.calls.length);
    expect(commits).toHaveLength(1);
    expect(boots).toHaveBeenCalledTimes(2);
    expect(calls).toEqual({
      "document.tree": 1,
      "document.storyContent": 1,
      "parts.read": 1,
    });
    boots.mockRestore();
  });
});
