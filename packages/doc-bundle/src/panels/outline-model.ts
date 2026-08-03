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

// PURE derivations for the outline/styles panel — everything the panel shows
// comes from the RETAINED `LoweredDoc` (the engine's Tier-0 lowering) plus the
// placement record. No React, no host: unit-testable as plain functions.

import type { Diagnostic, LoweredDoc } from "@paged-media/doc-host-model";

/** One outline row: a heading-styled paragraph or a table marker. */
export type OutlineEntry =
  | { kind: "heading"; level: number; text: string; blockIndex: number }
  | { kind: "table"; rows: number; cols: number; blockIndex: number };

export interface DocSummary {
  paragraphs: number;
  tables: number;
  images: number;
  hyperlinks: number;
  headings: number;
  /** `210×297pt · 2 cols` style line from the section geometry. */
  pageLine: string;
}

/** Cap the outline so a 1000-heading document cannot melt the panel. */
export const OUTLINE_CAP = 300;
const PREVIEW_CHARS = 80;

/** `ParagraphStyle/docx-Heading1` / a style NAMED "heading 2" → level. */
function headingLevel(
  styleId: string | null | undefined,
  styleNameById: Map<string, string>,
): number | null {
  if (!styleId) return null;
  const byId = /Heading\s*(\d+)?$/i.exec(styleId);
  if (byId) return byId[1] ? Number(byId[1]) : 1;
  const name = styleNameById.get(styleId);
  if (!name) return null;
  const byName = /^heading\s*(\d+)?/i.exec(name.trim());
  if (byName) return byName[1] ? Number(byName[1]) : 1;
  return null;
}

function paragraphText(runs: ReadonlyArray<{ text: string }>): string {
  const joined = runs.map((r) => r.text).join("");
  return joined.length > PREVIEW_CHARS
    ? `${joined.slice(0, PREVIEW_CHARS - 1)}…`
    : joined;
}

export function styleNameById(ir: LoweredDoc): Map<string, string> {
  return new Map(ir.styles.map((s) => [s.id, s.name]));
}

/** Headings (indented by level) + tables, in document order, capped. */
export function outlineEntries(ir: LoweredDoc): OutlineEntry[] {
  const names = styleNameById(ir);
  const out: OutlineEntry[] = [];
  ir.story.blocks.forEach((block, blockIndex) => {
    if (out.length >= OUTLINE_CAP) return;
    if (block.kind === "table") {
      out.push({ kind: "table", rows: block.rows, cols: block.cols, blockIndex });
      return;
    }
    const level = headingLevel(block.paraStyleId, names);
    if (level != null) {
      out.push({
        kind: "heading",
        level: Math.min(Math.max(level, 1), 6),
        text: paragraphText(block.runs) || "(empty heading)",
        blockIndex,
      });
    }
  });
  return out;
}

export function summarize(ir: LoweredDoc): DocSummary {
  let paragraphs = 0;
  let tables = 0;
  let images = 0;
  let hyperlinks = 0;
  let headings = 0;
  const names = styleNameById(ir);
  for (const block of ir.story.blocks) {
    if (block.kind === "table") {
      tables += 1;
      for (const cell of block.cells) {
        paragraphs += cell.paragraphs.length;
        for (const p of cell.paragraphs) {
          images += p.images?.length ?? 0;
          hyperlinks += p.runs.filter((r) => r.hyperlinkUrl).length;
        }
      }
      continue;
    }
    paragraphs += 1;
    images += block.images?.length ?? 0;
    hyperlinks += block.runs.filter((r) => r.hyperlinkUrl).length;
    if (headingLevel(block.paraStyleId, names) != null) headings += 1;
  }
  const s = ir.section;
  const pageLine = `${Math.round(s.pageWidthPt)}×${Math.round(s.pageHeightPt)}pt · ${
    s.columns
  } col${s.columns === 1 ? "" : "s"}`;
  return { paragraphs, tables, images, hyperlinks, headings, pageLine };
}

/** Diagnostics sorted error → warning → info (the ADR-007 payload). */
export function sortedDiagnostics(ir: LoweredDoc): Diagnostic[] {
  const rank: Record<Diagnostic["severity"], number> = {
    error: 0,
    warning: 1,
    info: 2,
  };
  return [...ir.diagnostics].sort(
    (a, b) => rank[a.severity] - rank[b.severity],
  );
}

// ── A minimal document store (activate writes, the panel subscribes) ──

export interface DocPanelDoc {
  fileName: string;
  ir: LoweredDoc;
  /** The placed frame's ElementId (opaque to the panel; handed back for
   *  `host.selection.set`). Null when placement failed. */
  frameId: unknown | null;
  storyId: string | null;
}

export interface DocStore {
  get(): DocPanelDoc | null;
  set(doc: DocPanelDoc | null): void;
  subscribe(listener: () => void): () => void;
}

export function createDocStore(): DocStore {
  let current: DocPanelDoc | null = null;
  const listeners = new Set<() => void>();
  return {
    get: () => current,
    set: (doc) => {
      current = doc;
      for (const l of [...listeners]) l();
    },
    subscribe: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
