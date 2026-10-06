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

// A placed Word document's PERSISTED state, read the one way every surface
// reads it: the frame's label names it, the container parts carry it.
//
//   label (x-paged:media.paged.doc on the host frame)  — the truth undo restores
//     { v: 1, data: { part, lowered?, blocks, sections?, fileName?, midPageColumns? } }
//   part  `part`     — the .docx bytes (the import source, or a content-addressed
//                      `sources/<sha256>.docx` an object-model write produced)
//   part  `lowered`  — the docLowered CACHE (`lowered/<sha256>.json`): the
//                      lowering of exactly those bytes. Missing (an older
//                      document, a file InDesign re-saved without our parts'
//                      cache) → the engine lowers the .docx again.
//
// Content-addressed parts are never overwritten, so an undo of the label
// lands on a state whose parts still exist.

import type { LoweredDoc } from "@paged-media/doc-host-model";
import { parseLoweredDoc } from "@paged-media/doc-host-model";
import type { BundleHost, ElementId, PartChange } from "@paged-media/plugin-api";

import { DocEngine } from "./engine.js";

/** This plugin's label key (the host derives the same from the manifest id). */
export const LABEL_KEY = "x-paged:media.paged.doc";

/** The label's `data`. */
export interface DocLabelData {
  /** The .docx part (relative to this plugin's part namespace). */
  part: string;
  /** The docLowered cache part, when one was written. */
  lowered?: string;
  /** Lowered story blocks (the outline's count). */
  blocks?: number;
  /** The story or stories the document lives in, in section order. */
  sections?: string[];
  fileName?: string;
  /** How the document was lowered (ADR 029); default true. */
  midPageColumns?: boolean;
}

/** One placed document, as the object model sees it. */
export interface DocState {
  frame: ElementId;
  data: DocLabelData;
  /** The label value as stored — the memo key. */
  labelText: string;
}

/** Every page item carrying this plugin's label, with the label. One tree
 *  read where the engine reports labels on tree rows (protocol 66+), else
 *  one label read per text frame. */
export async function discoverDocs(host: BundleHost): Promise<DocState[]> {
  const roots = await host.document.tree();
  const items: ElementId[] = [];
  const labels = new Map<string, string>();
  let reported = false;
  const walk = (nodes: ReadonlyArray<{ id?: ElementId | null; children?: unknown; pluginMetadata?: Array<{ key: string; value: string }> }> | undefined) => {
    for (const n of nodes ?? []) {
      const id = n.id as { kind?: string; id?: unknown } | null | undefined;
      if (id && typeof id.id === "string") {
        items.push(n.id as ElementId);
        if (Array.isArray(n.pluginMetadata)) {
          reported = true;
          const e = n.pluginMetadata.find((m) => m.key === LABEL_KEY);
          if (e) labels.set(id.id, e.value);
        }
      }
      walk(n.children as never);
    }
  };
  walk(roots as never);
  const out: DocState[] = [];
  for (const frame of items) {
    const key = (frame as { id: string }).id;
    let text = labels.get(key) ?? null;
    if (!reported && frame.kind === "textFrame") {
      const env = await host.document.getMetadata(frame);
      text = env ? JSON.stringify(env) : null;
    }
    const state = text ? stateOf(frame, text) : null;
    if (state) out.push(state);
  }
  return out;
}

function stateOf(frame: ElementId, labelText: string): DocState | null {
  try {
    const env = JSON.parse(labelText) as { v?: number; data?: Partial<DocLabelData> };
    if (!env || typeof env.data?.part !== "string") return null;
    return { frame, data: env.data as DocLabelData, labelText };
  } catch {
    return null;
  }
}

/** The .docx bytes of a document, or null when its part is missing. */
export async function readDocBytes(host: BundleHost, s: DocState): Promise<Uint8Array | null> {
  if (!host.supports("storage.parts@1")) return null;
  try {
    return await host.parts.read(s.data.part);
  } catch {
    return null;
  }
}

/** Lowerings already read, by label value (a label names exactly one). */
const lowerings = new Map<string, LoweredDoc>();

/** The lowering of a document's current state: the cache part, else the
 *  engine over the .docx part. Null when neither part can be read. */
export async function readLowered(host: BundleHost, s: DocState): Promise<LoweredDoc | null> {
  const memo = lowerings.get(s.labelText);
  if (memo) return memo;
  let ir: LoweredDoc | null = null;
  if (s.data.lowered && host.supports("storage.parts@1")) {
    try {
      const bytes = await host.parts.read(s.data.lowered);
      if (bytes) ir = parseLoweredDoc(new TextDecoder().decode(bytes));
    } catch {
      ir = null; // a missing or unreadable cache: lower again
    }
  }
  if (!ir) {
    const bytes = await readDocBytes(host, s);
    if (!bytes) return null;
    const engine = await DocEngine.boot();
    try {
      engine.loadDocx(bytes);
      engine.setMidPageColumns(s.data.midPageColumns ?? true);
      ir = engine.lowered();
    } finally {
      engine.dispose();
    }
  }
  if (lowerings.size > 32) lowerings.clear();
  lowerings.set(s.labelText, ir);
  return ir;
}

/** SHA-256 of `bytes`, lower-case hex. */
export async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes as Uint8Array<ArrayBuffer>);
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, "0")).join("");
}

/** A new persisted state for an edited .docx: the content-addressed .docx and
 *  docLowered parts, and the label that names them. Remembers the lowering
 *  under the new label so the next read needs no part read. */
export async function plannedState(
  s: DocState,
  bytes: Uint8Array,
  ir: LoweredDoc,
): Promise<{ parts: PartChange[]; labelText: string }> {
  const hash = await sha256(bytes);
  const part = `sources/${hash}.docx`;
  const lowered = `lowered/${hash}.json`;
  const data: DocLabelData = { ...s.data, part, lowered, blocks: ir.story.blocks.length };
  const labelText = JSON.stringify({ v: 1, data });
  lowerings.set(labelText, ir);
  return {
    parts: [
      { path: part, bytes },
      { path: lowered, bytes: new TextEncoder().encode(JSON.stringify(ir)) },
    ],
    labelText,
  };
}
