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

// The object model (ADR 323) on the real headless host and the real docx-js
// wasm: a mail-merge letter's fields, runs and blocks through `h.objects`
// (what the Node CLI, a Boa bridge and a data binding call). A write is ONE
// undo step that changes the poured text AND the .docx the label names, and
// one `document.undo()` takes both back.

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import type { BundleHost } from "@paged-media/plugin-api";
import type { HeadlessHost } from "@paged-media/plugin-sdk";

import { DocEngine } from "../../src/engine";
import { docAddress, PANEL_FIELDS, schemaOf } from "../../src/object-model";
import { docxPresent, fixture, openDoc, requireDocx, storyTexts } from "./harness";

requireDocx("doc conformance — object model");

const value = (v: unknown) => ({ kind: "value", value: v });
const FIRST = "Dear «First Name», from «City» and «City», page 1";

describe.skipIf(!docxPresent)("doc conformance — object model (real host + real docx-js) @feat:plugin-doc.object-model", () => {
  let h: HeadlessHost;
  let host: BundleHost;
  let story: string;
  const doc = docAddress("docxDocument", "docx_f0");

  beforeEach(async () => {
    const o = await openDoc("merge-letter.docx");
    h = o.h;
    host = o.host;
    story = o.opened.storyIds[0]!;
  });
  afterEach(() => h?.dispose());

  it("the bundle boots headless: four kinds and the typed commands", async () => {
    const kinds = (await h.objects.kinds()).filter((k) => k.owner === "media.paged.doc").map((k) => k.kind);
    expect(kinds).toEqual([
      "plugin:media.paged.doc/docxDocument",
      "plugin:media.paged.doc/block",
      "plugin:media.paged.doc/run",
      "plugin:media.paged.doc/mergeField",
    ]);
    const commands = (await h.objects.commands()).map((c) => c.id);
    expect(commands).toEqual(
      expect.arrayContaining([
        "media.paged.doc.command.placeDoc",
        "media.paged.doc.command.exportDocx",
        "media.paged.doc.command.fillMergeFields",
      ]),
    );
    expect(await storyTexts(host, story)).toEqual([FIRST, "Order «OrderNo»"]);
  });

  it("documents, blocks, runs and merge fields are listed and read", async () => {
    expect(await h.objects.query("plugin:media.paged.doc/docxDocument")).toEqual([doc]);
    expect(await h.objects.get(doc, "fileName")).toEqual(value("merge-letter.docx"));
    expect(await h.objects.get(doc, "mergeFieldNames")).toEqual(value(["First Name", "City", "OrderNo"]));
    expect(await h.objects.get(doc, "stories")).toEqual(value([`story:${story}`]));
    expect(await h.objects.query("plugin:media.paged.doc/block")).toEqual([
      docAddress("block", "docx_f0/b0"),
      docAddress("block", "docx_f0/b1"),
    ]);
    expect(await h.objects.get(docAddress("block", "docx_f0/b1"), "text")).toEqual(value("Order «OrderNo»"));
    const city = await h.objects.query('plugin:media.paged.doc/mergeField[name="City"]');
    expect(city).toEqual([docAddress("mergeField", "docx_f0/b0/f1"), docAddress("mergeField", "docx_f0/b0/f2")]);
    const first = docAddress("mergeField", "docx_f0/b0/f0");
    expect(await h.objects.get(first, "value")).toEqual(value("«First Name»"));
    expect(await h.objects.get(first, "before")).toEqual(value("Mx "));
    expect(await h.objects.get(first, "runs")).toEqual(
      value([docAddress("run", "docx_f0/b0/r1"), docAddress("run", "docx_f0/b0/r2")]),
    );
    expect(await h.objects.get(docAddress("run", "docx_f0/b0/r4"), "mergeField")).toEqual(value("City"));
    const content = await h.objects.get(doc, "content");
    expect((content as { value: Uint8Array }).value).toBeInstanceOf(Uint8Array);
  });

  it("a merge field write is one undo step: poured text, .docx and label together", async () => {
    const first = docAddress("mergeField", "docx_f0/b0/f0");
    const out = await h.objects.set(first, "value", "Ada Lovelace");
    expect(out).toMatchObject({ applied: true, undoSteps: 1 });
    expect(await h.objects.get(first, "value")).toEqual(value("Ada Lovelace"));
    expect(await h.objects.get(first, "name")).toEqual(value("First Name"));
    expect((await storyTexts(host, story))[0]).toBe("Dear Ada Lovelace, from «City» and «City», page 1");
    expect((await h.objects.get(doc, "sourcePart")) as { value: string }).toMatchObject({
      value: expect.stringMatching(/^sources\/[0-9a-f]{64}\.docx$/),
    });
    // The .docx the label now names carries the value AND still the field
    // code (Word re-merges it). The emptied second result run no longer
    // lowers, so RUN ids after it move; field and block ids do not.
    const part = (await h.objects.get(doc, "content")) as { value: Uint8Array };
    const engine = await DocEngine.boot();
    try {
      engine.loadDocx(part.value);
      const p0 = engine.lowered().story.blocks[0]!;
      expect(p0.kind === "paragraph" && p0.runs.map((r) => [r.text, r.mergeField?.name ?? null])).toEqual([
        ["Dear ", null],
        ["Ada Lovelace", "First Name"],
        [", from ", null],
        ["«City»", "City"],
        [" and ", null],
        ["«City»", "City"],
        [", page ", null],
        ["1", null],
      ]);
    } finally {
      engine.dispose();
    }
    await host.document.undo();
    expect(await h.objects.get(first, "value")).toEqual(value("«First Name»"));
    expect((await storyTexts(host, story))[0]).toBe(FIRST);
    expect(await h.objects.get(doc, "sourcePart")).toEqual(value(`${story}/source.docx`));
  });

  it("fillMergeFields (typed) fills by name in one step and names what it did not know", async () => {
    const out = await h.objects.invoke("media.paged.doc.command.fillMergeFields", {
      document: doc,
      values: { City: "London", OrderNo: "A-17", Nope: "x" },
    });
    expect(out).toEqual({ filled: 3, unknown: ["Nope"] });
    expect(await storyTexts(host, story)).toEqual([
      "Dear «First Name», from London and London, page 1",
      "Order A-17",
    ]);
    await host.document.undo();
    expect(await storyTexts(host, story)).toEqual([FIRST, "Order «OrderNo»"]);
  });

  it("a run write changes one run; a bad address or path is refused", async () => {
    const run = docAddress("run", "docx_f0/b1/r0");
    expect(await h.objects.set(run, "text", "Ref ")).toMatchObject({ applied: true, undoSteps: 1 });
    expect((await storyTexts(host, story))[1]).toBe("Ref «OrderNo»");
    expect(await h.objects.set(docAddress("run", "docx_f0/b1/r9"), "text", "x")).toMatchObject({ applied: false });
    expect((await h.objects.set(run, "characterStyle", "x")).code).toBe("readOnly");
    expect((await h.objects.set(run, "text", 3)).code).toBe("invalidValue");
  });

  it("a paragraph edited on the canvas since is refused, never written over", async () => {
    await host.document.mutate({ op: "insertText", args: { storyId: story, offset: 0, text: "PS " } } as never);
    const out = await h.objects.set(docAddress("mergeField", "docx_f0/b0/f0"), "value", "Ada");
    expect(out.applied).toBe(false);
    expect(out.reason).toMatch(/edited on the canvas/);
    // Another paragraph is still writable: the check is per paragraph.
    expect((await h.objects.set(docAddress("mergeField", "docx_f0/b1/f0"), "value", "A-1")).applied).toBe(true);
  });

  it("one batch writing a document through two kinds composes (the plugin-level planner)", async () => {
    const out = await h.objects.batch([
      { op: "set", address: docAddress("run", "docx_f0/b1/r0"), path: "text", value: "Ref " },
      { op: "set", address: docAddress("mergeField", "docx_f0/b0/f0"), path: "value", value: "Ada" },
      { op: "set", address: docAddress("mergeField", "docx_f0/b1/f0"), path: "value", value: "A-17" },
    ]);
    expect(out).toMatchObject({ applied: true, undoSteps: 1 });
    expect(await storyTexts(host, story)).toEqual(["Dear Ada, from «City» and «City», page 1", "Ref A-17"]);
    expect(await h.objects.get(docAddress("mergeField", "docx_f0/b1/f0"), "value")).toEqual(value("A-17"));
    await host.document.undo();
    expect(await storyTexts(host, story)).toEqual([FIRST, "Order «OrderNo»"]);
  });

  // OPEN BUG (save-back, not the object model): the export overlays the
  // story read-back on the .docx, and the engine hands back runs it COALESCED
  // (same style), which the run diff turns into inserts and deletes; the
  // wrapped field runs cannot be deleted, so the paragraph's text comes out
  // twice. It happens to an UNEDITED document with fields. `it.fails` keeps
  // the lane green and flips red the day save-back is fixed.
  it.fails("export of an unedited letter keeps its text once (save-back overlay duplicates coalesced runs)", async () => {
    const exported = (await h.objects.invoke("media.paged.doc.command.exportDocx", { document: doc })) as {
      bytes: Uint8Array;
    };
    const engine = await DocEngine.boot();
    try {
      engine.loadDocx(exported.bytes);
      const p0 = engine.lowered().story.blocks[0]!;
      expect(p0.kind === "paragraph" && p0.runs.map((r) => r.text).join("")).toBe(FIRST);
    } finally {
      engine.dispose();
    }
  });

  it("placeDoc (typed) places a second letter embedded, and its fields are objects", async () => {
    const out = (await h.objects.invoke("media.paged.doc.command.placeDoc", {
      bytes: fixture("merge-letter.docx"),
      fileName: "second.docx",
    })) as { document: string; story: string };
    expect(out.document).toMatch(/^plugin:media\.paged\.doc\/docxDocument\//);
    expect(await h.objects.query("plugin:media.paged.doc/docxDocument")).toHaveLength(2);
    expect(await h.objects.get(out.document, "fileName")).toEqual(value("second.docx"));
    const frame = out.document.split("/").pop()!;
    const field = docAddress("mergeField", `${frame}/b1/f0`);
    expect(await h.objects.set(field, "value", "B-2")).toMatchObject({ applied: true, undoSteps: 1 });
    expect((await storyTexts(host, out.story.slice("story:".length)))[1]).toBe("Order B-2");
    // The first document is untouched.
    expect(await storyTexts(host, story)).toEqual([FIRST, "Order «OrderNo»"]);
  });

  it("a page item scope lists that frame's objects only (the panel's selection, ADR 323 §4)", async () => {
    const out = (await h.objects.invoke("media.paged.doc.command.placeDoc", {
      bytes: fixture("merge-letter.docx"),
      fileName: "scoped.docx",
    })) as { document: string };
    const frame = out.document.split("/").pop()!;
    expect(await h.objects.query(`frame:${frame} > plugin:media.paged.doc/docxDocument`)).toEqual([out.document]);
    const fields = await h.objects.query(`frame:${frame} > plugin:media.paged.doc/mergeField`);
    expect(fields.length).toBeGreaterThan(0);
    expect(fields.every((f) => f.includes(`/${frame}/`))).toBe(true);
    const block = docAddress("block", `${frame}/b1`);
    expect(await h.objects.query(`${block} > plugin:media.paged.doc/mergeField`)).toEqual([
      docAddress("mergeField", `${frame}/b1/f0`),
    ]);
    const runs = await h.objects.query(`${block} > plugin:media.paged.doc/run`);
    expect(runs.length).toBeGreaterThan(0);
    expect(runs.every((r) => r.startsWith(`plugin:media.paged.doc/run/${frame}/b1/`))).toBe(true);
  });

  it("every panel field is a row of its kind", () => {
    for (const f of PANEL_FIELDS) expect(schemaOf(f.kind).some((r) => r.path === f.path), `${f.kind}.${f.path}`).toBe(true);
  });
});
