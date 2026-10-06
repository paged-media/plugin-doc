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

// paged.doc's OBJECT MODEL (ADR 323): a placed Word document, its blocks,
// runs and mail-merge fields, addressable through `host.objects` by scripts,
// the Node CLI, data bindings and schema-driven fields.
//
//   plugin:media.paged.doc/docxDocument/<frame>
//   plugin:media.paged.doc/block/<frame>/b<i>              paragraph or table
//   plugin:media.paged.doc/block/<frame>/b<i>/i<k>         an inline picture
//   plugin:media.paged.doc/run/<frame>/b<i>/r<j>           a body run
//   plugin:media.paged.doc/run/<frame>/b<i>/c<k>/p<m>/r<j> a table-cell run
//   plugin:media.paged.doc/mergeField/<frame>/b<i>/f<o>    the o-th MERGEFIELD
//                                                          of a paragraph
//
// Ids are LOWERED story coordinates — the coordinates save-back patches by —
// so they are stable for as long as the document's block structure is (text
// writes never change it). A binding stores a selector over the field's
// NAME (`mergeField[name="First Name"]`), never an id (ADR 559).
//
// The state is what the frame's label names (doc-state.ts): the .docx part
// and the docLowered cache. A write (`run.text`, `mergeField.value`) is
// one ObjectWrite, ONE undo step:
//   · the engine (docx-js `save_object_edits`) patches the .docx — the run's
//     `w:t` only; a merge field's RESULT runs, never its field code;
//   · the poured native text changes in the same batch (deleteRange +
//     insertText at the run's place in its story, checked against the live
//     paragraph first: a paragraph edited on the canvas since is refused,
//     not written over);
//   · the new .docx and its lowering are content-addressed parts, and the
//     label names them — undo reverts the label to the previous state.
// No DOCX semantics here: which runs a field spans is the engine's answer,
// where a run was poured is doc-host-model's.

import type { LoweredDoc, LoweredParagraph } from "@paged-media/doc-host-model";
import { bodyRunLocation } from "@paged-media/doc-host-model";
import type {
  Address,
  BundleHost,
  ElementId,
  MutationInput,
  ObjectKindContribution,
  ObjectModelContribution,
  ObjectModelHandle,
  ObjectOp,
  ObjectValue,
  ObjectWrite,
  PartChange,
  PropertySchema,
  TypedCommandContribution,
} from "@paged-media/plugin-api";
import { escapeLabelValue, fromElementId, parseAddress } from "@paged-media/plugin-sdk";

import manifestJson from "../manifest.json" with { type: "json" };
import { DocEngine } from "./engine.js";
import {
  discoverDocs,
  LABEL_KEY,
  plannedState,
  readDocBytes,
  readLowered,
  type DocState,
} from "./doc-state.js";

export const PLUGIN_ID = "media.paged.doc";

type KindName = "docxDocument" | "block" | "run" | "mergeField";

/** The schema rows the manifest declares for `kind` (one declaration). */
export function schemaOf(kind: KindName): readonly PropertySchema[] {
  const k = manifestJson.contributes.objectModel.kinds.find((x) => x.kind === kind);
  if (!k) throw new Error(`paged.doc: the manifest declares no object kind "${kind}"`);
  return k.schema as unknown as readonly PropertySchema[];
}

const decl = (id: string) => {
  const c = manifestJson.contributes.objectModel.commands.find((x) => x.id === id);
  if (!c) throw new Error(`paged.doc: the manifest declares no typed command "${id}"`);
  return c as unknown as Pick<TypedCommandContribution, "id" | "title" | "args" | "result">;
};

/** The address of one of this bundle's objects. */
export const docAddress = (kind: KindName, id: string): Address => `plugin:${PLUGIN_ID}/${kind}/${id}`;

/** The schema-driven panel fields (ADR 323 / Wave 4) for the host's shared
 *  `PropertyField`: what the outline panel shows per selection. */
export const PANEL_FIELDS: ReadonlyArray<{ kind: KindName; path: string; group: string }> = [
  { kind: "docxDocument", path: "fileName", group: "Document" },
  { kind: "docxDocument", path: "blockCount", group: "Document" },
  { kind: "docxDocument", path: "mergeFieldNames", group: "Mail merge" },
  { kind: "mergeField", path: "value", group: "Mail merge" },
  { kind: "block", path: "paragraphStyle", group: "Block" },
  { kind: "run", path: "text", group: "Run" },
];

// ------------------------------------------------------------ ids

/** A parsed object id: the frame plus the lowered coordinates. */
interface DocId {
  frame: string;
  block?: number;
  image?: number;
  cell?: number;
  para?: number;
  run?: number;
  field?: number;
}

const SEG = /^(b|i|c|p|r|f)(\d+)$/;

function parseId(address: Address, kind: KindName): DocId | null {
  const a = parseAddress(address);
  if (!a || a.kind !== "plugin" || a.plugin !== PLUGIN_ID || a.objectKind !== kind) return null;
  const [frame, ...segs] = a.id.split("/");
  if (!frame) return null;
  const id: DocId = { frame };
  const key = { b: "block", i: "image", c: "cell", p: "para", r: "run", f: "field" } as const;
  for (const s of segs) {
    const m = SEG.exec(s);
    if (!m) return null;
    const k = key[m[1] as keyof typeof key];
    if (id[k] !== undefined) return null;
    id[k] = Number(m[2]);
  }
  const shape: Record<KindName, (d: DocId) => boolean> = {
    docxDocument: () => segs.length === 0,
    block: (d) => d.block !== undefined && d.cell === undefined && d.run === undefined && d.field === undefined,
    run: (d) => d.block !== undefined && d.run !== undefined && (d.cell === undefined) === (d.para === undefined),
    mergeField: (d) => d.block !== undefined && d.field !== undefined && (d.cell === undefined) === (d.para === undefined),
  };
  return shape[kind](id) ? id : null;
}

const idText = (d: DocId, tail: string) =>
  [d.frame, `b${d.block}`, ...(d.cell !== undefined ? [`c${d.cell}`, `p${d.para}`] : []), tail].join("/");

// ----------------------------------------------------- reading helpers

const val = (value: unknown): ObjectValue => ({ kind: "value", value });
const refuse = (code: "badAddress" | "unknownAddress" | "unknownPath", reason: string): ObjectValue => ({
  kind: "refused",
  code,
  reason,
});

/** A document's state and lowering by frame id (one discovery per call). */
async function docOf(host: BundleHost, frame: string): Promise<{ state: DocState; ir: LoweredDoc } | null> {
  const state = (await discoverDocs(host)).find((s) => (s.frame as { id: string }).id === frame);
  if (!state) return null;
  const ir = await readLowered(host, state);
  return ir ? { state, ir } : null;
}

/** The paragraph a run / field id names (body or table cell). */
function paragraphOf(ir: LoweredDoc, d: DocId): LoweredParagraph | null {
  const b = ir.story.blocks[d.block ?? -1];
  if (!b) return null;
  if (d.cell === undefined) return b.kind === "paragraph" ? b : null;
  if (b.kind !== "table") return null;
  return b.cells[d.cell]?.paragraphs[d.para ?? -1] ?? null;
}

/** Every paragraph with its id prefix: body paragraphs and table cells. */
function paragraphs(ir: LoweredDoc, frame: string): Array<{ d: DocId; p: LoweredParagraph }> {
  const out: Array<{ d: DocId; p: LoweredParagraph }> = [];
  ir.story.blocks.forEach((b, block) => {
    if (b.kind === "paragraph") out.push({ d: { frame, block }, p: b });
    else
      b.cells.forEach((c, cell) =>
        c.paragraphs.forEach((p, para) => out.push({ d: { frame, block, cell, para }, p })),
      );
  });
  return out;
}

/** The merge fields of a paragraph: occurrence → its result run indices. */
function fieldsOf(p: LoweredParagraph): Map<number, number[]> {
  const out = new Map<number, number[]>();
  p.runs.forEach((r, i) => {
    const o = r.mergeField?.occurrence;
    if (o === undefined) return;
    out.set(o, [...(out.get(o) ?? []), i]);
  });
  return out;
}

const textOfParagraph = (p: LoweredParagraph) => p.runs.map((r) => r.text).join("");

// ---------------------------------------------------------- writing

class Rejected extends Error {}

/** What one document's writes come to, before they are committed. */
interface DocPlan {
  state: DocState;
  ir: LoweredDoc;
  runs: Array<{ block: number; run: number; text: string }>;
  mergeFields: Array<{ block: number; occurrence: number; value: string }>;
  /** Native text writes: one per edited run, by its lowered place. */
  native: Array<{ block: number; run: number; text: string }>;
}

const utf8 = (s: string) => new TextEncoder().encode(s).length;

/** Plan and turn the per-document plans into ONE ObjectWrite. */
async function commitPlans(host: BundleHost, plans: Map<string, DocPlan>): Promise<ObjectWrite> {
  const labels: Array<{ frame: ElementId; text: string }> = [];
  const parts: PartChange[] = [];
  const mutations: MutationInput[] = [];
  for (const plan of plans.values()) {
    const bytes = await readDocBytes(host, plan.state);
    if (!bytes) throw new Rejected(`the .docx part ${plan.state.data.part} is missing — re-link the document`);
    // 1. The .docx, patched by the engine, and its new lowering.
    const engine = await DocEngine.boot();
    let saved: Uint8Array;
    try {
      engine.loadDocx(bytes);
      engine.setMidPageColumns(plan.state.data.midPageColumns ?? true);
      let out: { bytes: Uint8Array; skips: string[] };
      try {
        out = engine.saveObjectEdits({ runs: plan.runs, mergeFields: plan.mergeFields });
      } catch (err) {
        throw new Rejected(String(err instanceof Error ? err.message : err));
      }
      if (out.skips.length > 0) throw new Rejected(`save-back refused the edit: ${out.skips.join("; ")}`);
      saved = out.bytes;
    } finally {
      engine.dispose();
    }
    const after = await DocEngine.boot();
    let ir: LoweredDoc;
    try {
      after.loadDocx(saved);
      after.setMidPageColumns(plan.state.data.midPageColumns ?? true);
      ir = after.lowered();
    } finally {
      after.dispose();
    }
    // 2. The native text: each edited run's range in the LIVE story, checked.
    const stories = plan.state.data.sections ?? [];
    const edits: Array<{ story: string; start: number; end: number; text: string }> = [];
    const contents = new Map<string, Array<{ runs: Array<{ text: string }> }>>();
    for (const n of plan.native) {
      const at = bodyRunLocation(plan.ir, n.block, n.run);
      if (!at) throw new Rejected(`block ${n.block} run ${n.run} has no place in a story this document can write`);
      const storyId = stories[at.story];
      if (!storyId) throw new Rejected("the document's label names no story to write (placed by an older version)");
      let paras = contents.get(storyId);
      if (!paras) {
        const c = await host.document.storyContent(storyId);
        if (!c) throw new Rejected(`story ${storyId} cannot be read`);
        paras = c.paragraphs;
        contents.set(storyId, paras);
      }
      const live = paras[at.paragraph];
      if (!live || live.runs.map((r) => r.text).join("") !== at.paragraphText) {
        throw new Rejected(
          `the paragraph of block ${n.block} was edited on the canvas since the document was last written; ` +
            "export the document (save-back) before writing it through the object model",
        );
      }
      let start = 0;
      for (let k = 0; k < at.paragraph; k += 1) start += utf8(paras[k]!.runs.map((r) => r.text).join("")) + 1;
      start += utf8(at.prefix);
      edits.push({ story: storyId, start, end: start + utf8(at.text), text: n.text });
    }
    // Back to front, so an edit never moves the offsets of the next.
    edits.sort((a, b) => (a.story === b.story ? b.start - a.start : a.story < b.story ? -1 : 1));
    for (const e of edits) {
      if (e.end > e.start) mutations.push({ op: "deleteRange", args: { storyId: e.story, start: e.start, end: e.end } } as MutationInput);
      if (e.text.length > 0) mutations.push({ op: "insertText", args: { storyId: e.story, offset: e.start, text: e.text } } as MutationInput);
    }
    // 3. The new state's parts and the label naming them.
    const next = await plannedState(plan.state, saved, ir);
    parts.push(...next.parts);
    labels.push({ frame: plan.state.frame, text: next.labelText });
  }
  if (labels.length === 0) return { kind: "mutations", mutations };
  const [first, ...rest] = labels;
  for (const l of rest) {
    mutations.push({
      op: "setPluginMetadata",
      args: { elementId: l.frame, key: LABEL_KEY, value: escapeLabelValue(l.text), caller: PLUGIN_ID },
    } as MutationInput);
  }
  return {
    kind: "both",
    mutations,
    parts,
    host: fromElementId(first!.frame),
    labelKey: LABEL_KEY,
    labelValue: first!.text,
  };
}

/** The PLUGIN-level planner (ADR 323): every write op of a batch on any of
 *  this bundle's kinds, grouped into ONE plan per document — run and merge
 *  field writes on the same document compose into one save-back and one
 *  label — answered as ONE write. */
async function planDocOps(host: BundleHost, ops: readonly ObjectOp[]): Promise<ObjectWrite> {
  try {
    const plans = new Map<string, DocPlan>();
    const states = await discoverDocs(host);
    for (const op of ops) {
      if (op.op !== "set") throw new Rejected(`paged.doc objects take set only (${op.op} refused)`);
      const a = parseAddress(op.address);
      const kind = a && a.kind === "plugin" ? a.objectKind : "";
      if (kind !== "run" && kind !== "mergeField") {
        throw new Rejected(`${op.address}: a Word document is written through its runs and merge fields`);
      }
      const d = parseId(op.address, kind);
      if (!d) throw new Rejected(`${op.address} is not a ${kind} address`);
      if (kind === "run" ? op.path !== "text" : op.path !== "value") {
        throw new Rejected(`${kind}.${op.path} is not writable`);
      }
      if (d.cell !== undefined) throw new Rejected(`${op.address}: table-cell text is not writable through the object model yet`);
      let p = plans.get(d.frame);
      if (!p) {
        const state = states.find((s) => (s.frame as { id: string }).id === d.frame);
        if (!state) throw new Rejected(`${op.address}: no Word document on frame ${d.frame}`);
        const ir = await readLowered(host, state);
        if (!ir) throw new Rejected(`${op.address}: the document's parts are missing`);
        p = { state, ir, runs: [], mergeFields: [], native: [] };
        plans.set(d.frame, p);
      }
      const para = paragraphOf(p.ir, d);
      if (!para) throw new Rejected(`${op.address}: no such paragraph`);
      const value = op.value as string;
      if (kind === "run") {
        if (d.run! >= para.runs.length) throw new Rejected(`${op.address}: no such run`);
        p.runs.push({ block: d.block!, run: d.run!, text: value });
        p.native.push({ block: d.block!, run: d.run!, text: value });
      } else {
        const runs = fieldsOf(para).get(d.field!);
        if (!runs) throw new Rejected(`${op.address}: no such merge field`);
        p.mergeFields.push({ block: d.block!, occurrence: d.field!, value });
        runs.forEach((r, k) => p!.native.push({ block: d.block!, run: r, text: k === 0 ? value : "" }));
      }
    }
    return await commitPlans(host, plans);
  } catch (err) {
    if (err instanceof Rejected) return { kind: "rejected", reason: err.message };
    throw err;
  }
}

// ------------------------------------------------------------- kinds

function documentKind(host: BundleHost): ObjectKindContribution {
  return {
    kind: "docxDocument",
    title: "Word document",
    schema: schemaOf("docxDocument"),
    content: {
      kind: "blob",
      mime: ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
    },
    hostOf: (address) => {
      const d = parseId(address, "docxDocument");
      return d ? `frame:${d.frame}` : null;
    },
    async list() {
      return (await discoverDocs(host)).map((s) => docAddress("docxDocument", (s.frame as { id: string }).id));
    },
    async get(address, path) {
      const d = parseId(address, "docxDocument");
      if (!d) return refuse("badAddress", `${address} is not a Word document address`);
      const doc = await docOf(host, d.frame);
      if (!doc) return refuse("unknownAddress", `${address}: no Word document on frame ${d.frame}`);
      const { state, ir } = doc;
      const all = paragraphs(ir, d.frame);
      switch (path) {
        case "fileName":
          return val(state.data.fileName ?? null);
        case "frame":
          return val(fromElementId(state.frame));
        case "stories":
          return val((state.data.sections ?? []).map((s) => `story:${s}`));
        case "sourcePart":
          return val(state.data.part);
        case "blockCount":
          return val(ir.story.blocks.length);
        case "paragraphCount":
          return val(ir.story.blocks.filter((b) => b.kind === "paragraph").length);
        case "tableCount":
          return val(ir.story.blocks.filter((b) => b.kind === "table").length);
        case "imageCount":
          return val(all.reduce((n, x) => n + (x.p.images?.length ?? 0), 0));
        case "mergeFieldNames": {
          const names: string[] = [];
          for (const { p } of all) for (const r of p.runs) if (r.mergeField && !names.includes(r.mergeField.name)) names.push(r.mergeField.name);
          return val(names);
        }
        case "diagnostics":
          return val(ir.diagnostics.map((x) => x.message));
        case "content": {
          const bytes = await readDocBytes(host, state);
          return bytes ? val(bytes) : refuse("unknownAddress", `${address}: the .docx part is missing`);
        }
        default:
          return refuse("unknownPath", `docxDocument has no path "${path}"`);
      }
    },
  };
}

function blockKind(host: BundleHost): ObjectKindContribution {
  return {
    kind: "block",
    title: "Word block",
    schema: schemaOf("block"),
    hostOf: (address) => {
      const d = parseId(address, "block");
      return d ? `frame:${d.frame}` : null;
    },
    async list(query) {
      const scoped = query.within ? parseId(query.within, "docxDocument") : null;
      const out: Address[] = [];
      for (const s of await discoverDocs(host)) {
        const frame = (s.frame as { id: string }).id;
        if (scoped && scoped.frame !== frame) continue;
        const ir = await readLowered(host, s);
        ir?.story.blocks.forEach((b, i) => {
          out.push(docAddress("block", `${frame}/b${i}`));
          if (b.kind === "paragraph") (b.images ?? []).forEach((_, k) => out.push(docAddress("block", `${frame}/b${i}/i${k}`)));
        });
      }
      return out;
    },
    async get(address, path) {
      const d = parseId(address, "block");
      if (!d) return refuse("badAddress", `${address} is not a Word block address`);
      const doc = await docOf(host, d.frame);
      const b = doc?.ir.story.blocks[d.block!];
      if (!doc || !b) return refuse("unknownAddress", `${address}: no such block`);
      const image = d.image !== undefined ? (b.kind === "paragraph" ? b.images?.[d.image] : undefined) : undefined;
      if (d.image !== undefined && !image) return refuse("unknownAddress", `${address}: no such picture`);
      switch (path) {
        case "kind":
          return val(image ? "image" : b.kind);
        case "index":
          return val(d.block);
        case "text":
          if (image) return val("");
          if (b.kind === "paragraph") return val(textOfParagraph(b));
          return val(
            Array.from({ length: b.rows }, (_, r) =>
              b.cells.filter((c) => c.row === r).map((c) => c.paragraphs.map(textOfParagraph).join("\n")).join("\t"),
            ).join("\n"),
          );
        case "paragraphStyle":
          return val(!image && b.kind === "paragraph" ? (b.paraStyleId ?? null) : null);
        case "runCount":
          return val(!image && b.kind === "paragraph" ? b.runs.length : 0);
        case "rows":
          return val(!image && b.kind === "table" ? b.rows : null);
        case "columns":
          return val(!image && b.kind === "table" ? b.cols : null);
        case "widthPt":
          return val(image ? image.widthPt : null);
        case "heightPt":
          return val(image ? image.heightPt : null);
        case "document":
          return val(docAddress("docxDocument", d.frame));
        default:
          return refuse("unknownPath", `block has no path "${path}"`);
      }
    },
  };
}

function runKind(host: BundleHost): ObjectKindContribution {
  return {
    kind: "run",
    title: "Word run",
    schema: schemaOf("run"),
    hostOf: (address) => {
      const d = parseId(address, "run");
      return d ? `frame:${d.frame}` : null;
    },
    async list(query) {
      const scopedDoc = query.within ? parseId(query.within, "docxDocument") : null;
      const scopedBlock = query.within ? parseId(query.within, "block") : null;
      const out: Address[] = [];
      for (const s of await discoverDocs(host)) {
        const frame = (s.frame as { id: string }).id;
        const scope = scopedDoc ?? scopedBlock;
        if (scope && scope.frame !== frame) continue;
        const ir = await readLowered(host, s);
        for (const { d, p } of ir ? paragraphs(ir, frame) : []) {
          if (scopedBlock && scopedBlock.block !== d.block) continue;
          p.runs.forEach((_, j) => out.push(docAddress("run", idText(d, `r${j}`))));
        }
      }
      return out;
    },
    async get(address, path) {
      const d = parseId(address, "run");
      if (!d) return refuse("badAddress", `${address} is not a Word run address`);
      const doc = await docOf(host, d.frame);
      const r = doc ? paragraphOf(doc.ir, d)?.runs[d.run!] : undefined;
      if (!r) return refuse("unknownAddress", `${address}: no such run`);
      switch (path) {
        case "text":
          return val(r.text);
        case "characterStyle":
          return val(r.charStyleId ?? null);
        case "hyperlink":
          return val(r.hyperlinkUrl ?? null);
        case "mergeField":
          return val(r.mergeField?.name ?? null);
        case "block":
          return val(docAddress("block", `${d.frame}/b${d.block}`));
        default:
          return refuse("unknownPath", `run has no path "${path}"`);
      }
    },
  };
}

function mergeFieldKind(host: BundleHost): ObjectKindContribution {
  return {
    kind: "mergeField",
    title: "Merge field",
    schema: schemaOf("mergeField"),
    hostOf: (address) => {
      const d = parseId(address, "mergeField");
      return d ? `frame:${d.frame}` : null;
    },
    async list(query) {
      const scoped = query.within ? parseId(query.within, "docxDocument") : null;
      const out: Address[] = [];
      for (const s of await discoverDocs(host)) {
        const frame = (s.frame as { id: string }).id;
        if (scoped && scoped.frame !== frame) continue;
        const ir = await readLowered(host, s);
        for (const { d, p } of ir ? paragraphs(ir, frame) : []) {
          for (const o of fieldsOf(p).keys()) out.push(docAddress("mergeField", idText(d, `f${o}`)));
        }
      }
      return out;
    },
    async get(address, path) {
      const d = parseId(address, "mergeField");
      if (!d) return refuse("badAddress", `${address} is not a merge field address`);
      const doc = await docOf(host, d.frame);
      const p = doc ? paragraphOf(doc.ir, d) : null;
      const runs = p ? fieldsOf(p).get(d.field!) : undefined;
      if (!p || !runs) return refuse("unknownAddress", `${address}: no such merge field`);
      const f = p.runs[runs[0]!]!.mergeField!;
      switch (path) {
        case "name":
          return val(f.name);
        case "value":
          return val(runs.map((i) => p.runs[i]!.text).join(""));
        case "instruction":
          return val(f.instruction);
        case "before":
          return val(f.before ?? null);
        case "after":
          return val(f.after ?? null);
        case "format":
          return val(f.format ?? []);
        case "mapped":
          return val(f.mapped ?? false);
        case "writable":
          return val(d.cell === undefined && (bodyRunLocation(doc!.ir, d.block!, runs[0]!) !== null));
        case "runs":
          return val(runs.map((i) => docAddress("run", idText(d, `r${i}`))));
        case "block":
          return val(docAddress("block", `${d.frame}/b${d.block}`));
        case "document":
          return val(docAddress("docxDocument", d.frame));
        default:
          return refuse("unknownPath", `mergeField has no path "${path}"`);
      }
    },
  };
}

// ---------------------------------------------------------- commands

function typedCommands(
  host: BundleHost,
  place: (name: string, bytes: Uint8Array) => Promise<{ frame: ElementId; story: string | null } | null>,
  exportBytes: (state: DocState) => Promise<Uint8Array | null>,
): TypedCommandContribution[] {
  const cmd = <A>(id: string, handler: TypedCommandContribution<A>["handler"]): TypedCommandContribution =>
    ({ ...decl(id), handler }) as TypedCommandContribution;
  const docArg = async (ref: string): Promise<DocState> => {
    const d = parseId(ref, "docxDocument");
    const state = d ? (await discoverDocs(host)).find((s) => (s.frame as { id: string }).id === d.frame) : undefined;
    if (!state) throw new Error(`invalidValue: ${ref} is not a placed Word document`);
    return state;
  };
  return [
    cmd<{ bytes: Uint8Array; fileName: string }>("media.paged.doc.command.placeDoc", async (_ctx, args) => {
      const placed = await place(args.fileName, args.bytes);
      if (!placed) throw new Error("failed: the document could not be placed (see the log and diagnostics)");
      const frame = (placed.frame as { id: string }).id;
      return {
        document: docAddress("docxDocument", frame),
        frame: fromElementId(placed.frame),
        story: placed.story ? `story:${placed.story}` : "",
      };
    }),
    cmd<{ document: string }>("media.paged.doc.command.exportDocx", async (_ctx, args) => {
      const state = await docArg(args.document);
      const bytes = await exportBytes(state);
      if (!bytes) throw new Error(`failed: ${args.document} has no .docx part to export`);
      return { bytes, fileName: state.data.fileName ?? "document.docx" };
    }),
    cmd<{ document: string; values: Record<string, string> }>(
      "media.paged.doc.command.fillMergeFields",
      async (ctx, args) => {
        const state = await docArg(args.document);
        const frame = (state.frame as { id: string }).id;
        const fields = await ctx.objects.query(`plugin:${PLUGIN_ID}/mergeField`);
        const ops: ObjectOp[] = [];
        const seen = new Set<string>();
        for (const address of fields) {
          if (!address.startsWith(docAddress("mergeField", `${frame}/`))) continue;
          const name = await ctx.objects.get(address, "name");
          const n = name.kind === "value" ? (name.value as string) : "";
          if (!Object.prototype.hasOwnProperty.call(args.values, n)) continue;
          seen.add(n);
          ops.push({ op: "set", address, path: "value", value: args.values[n] });
        }
        if (ops.length > 0) {
          const out = await ctx.objects.batch(ops, { label: "Fill merge fields" });
          if (!out.applied) throw new Error(`failed: ${out.reason ?? "the fields were not written"}`);
        }
        return { filled: ops.length, unknown: Object.keys(args.values).filter((k) => !seen.has(k)) };
      },
    ),
  ];
}

// ------------------------------------------------------------ entry

export interface DocObjectModelDeps {
  /** Place a .docx into the current document (the bundle's embedded place). */
  place(name: string, bytes: Uint8Array): Promise<{ frame: ElementId; story: string | null } | null>;
  /** A document's bytes for export: its save-back, or the stored part. */
  exportBytes(state: DocState): Promise<Uint8Array | null>;
}

/** The bundle's object model over `host` (headless). */
export function createDocObjectModel(host: BundleHost, deps: DocObjectModelDeps): ObjectModelContribution {
  return {
    kinds: [documentKind(host), blockKind(host), runKind(host), mergeFieldKind(host)],
    commands: typedCommands(host, deps.place, deps.exportBytes),
    batch: (ops) => planDocOps(host, ops),
  };
}

/** Register it when the host has the door (plugin-api 0.2.43); `null` on an
 *  older host. */
export function contributeDocObjectModel(host: BundleHost, deps: DocObjectModelDeps): ObjectModelHandle | null {
  if (!host.supports("contribute.objectModel@1")) return null;
  return host.contribute.objectModel(createDocObjectModel(host, deps));
}
