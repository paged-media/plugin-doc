// A refused style batch must not cost the document every style: the
// bundle falls back to op-by-op and reports what the engine refused.
import { describe, expect, it } from "vitest";

import { applyStyleOps, STYLE_DIAGNOSTICS_KEY } from "../src/pour.js";

type Op = { op: string; args: Record<string, unknown> };

/** A host whose engine refuses one style path, and refuses any batch
 *  containing it (a batch is all-or-nothing). */
function fakeHost(refusedPath: string) {
  const applied: Op[] = [];
  const diagnostics = new Map<string, unknown[]>();
  const refuses = (o: Op) => o.op === "setStyleProperty" && o.args.path === refusedPath;
  const host = {
    document: {
      async mutate(m: Op) {
        const ops = m.op === "batch" ? (m.args.ops as Op[]) : [m];
        if (ops.some(refuses)) return { applied: false, error: { kind: "notImplemented" } };
        applied.push(...ops);
        return { applied: true };
      },
    },
    diagnostics: { set: (k: string, v: unknown[]) => diagnostics.set(k, v) },
  };
  return { host, applied, diagnostics };
}

const ops: Op[] = [
  { op: "createParagraphStyle", args: { selfId: "ParagraphStyle/a" } },
  { op: "setStyleProperty", args: { styleId: "ParagraphStyle/a", path: "characterLeading" } },
  { op: "setStyleProperty", args: { styleId: "ParagraphStyle/a", path: "paragraphStartParagraph" } },
];

describe("applyStyleOps", () => {
  it("applies the catalog as one batch when the engine takes it", async () => {
    const { host, applied, diagnostics } = fakeHost("none");
    expect(await applyStyleOps(host as never, ops as never)).toEqual([]);
    expect(applied).toHaveLength(3);
    expect(diagnostics.get(STYLE_DIAGNOSTICS_KEY)).toEqual([]);
  });

  it("keeps every other style when the engine refuses one property", async () => {
    const { host, applied, diagnostics } = fakeHost("paragraphStartParagraph");
    const refused = await applyStyleOps(host as never, ops as never);
    expect(applied.map((o) => o.args.path ?? o.op)).toEqual(["createParagraphStyle", "characterLeading"]);
    expect(refused).toHaveLength(1);
    expect(refused[0]).toContain("paragraphStartParagraph on ParagraphStyle/a");
    expect(diagnostics.get(STYLE_DIAGNOSTICS_KEY)).toHaveLength(1);
  });
});
