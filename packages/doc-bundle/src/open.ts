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

// Standalone open (thoughts ADR 029): a Word document becomes the WHOLE
// document, not content in a frame. The engine's skeleton (docx-skeleton, a
// minimal IDML package it writes itself — no core contact) gives each Word
// section a page of its size, its margin box as the frame, and an empty
// story. IDML cannot carry a grow rule, so after opening, each story gets
// one on the wire (`setFlowGrowRule`, protocol 64): the engine then adds
// pages while the story oversets, matching Word's own pagination (measured
// against Word's PDF export of docx_conformance::pagination_docx()). This
// file opens that skeleton, sets the rules, and pours each section's blocks
// into its story with the shared pour.

import type { LoweredDoc } from "@paged-media/doc-host-model";
import { buildStoryBlocks, buildStyleMutations, sectionBlocks } from "@paged-media/doc-host-model";
import type { BundleHost, ElementId, Mutation } from "@paged-media/plugin-api";

import type { DocEngine } from "./engine.js";
import { applyStyleOpsReporting, pourSteps } from "./pour.js";

/** What a standalone open produced: the section stories (save-back reads
 *  them back in this order), the frame carrying the binding, the lowering
 *  that was poured, and whether it uses span/split columns (save-back must
 *  lower the same way, `DocEngine.setMidPageColumns`). */
export interface Opened {
  storyIds: string[];
  frameId: ElementId;
  ir: LoweredDoc;
  midPageColumns: boolean;
}

/** Diagnostics key for what standalone open could not do on this engine. */
export const OPEN_DIAGNOSTICS_KEY = "media.paged.doc/open";

/** Diagnostics key for the mid-page column fallback (ADR 029). */
export const COLUMNS_DIAGNOSTICS_KEY = "media.paged.doc/columns";

/** The span/split column paths (core protocol 64, ADR 028 addendum) a
 *  mid-page column change lowers to. An engine that refuses any of them
 *  cannot lay the lowering out. */
export const SPAN_COLUMN_PATHS: readonly string[] = [
  "paragraphSpanColumnType",
  "paragraphSpanSplitColumnCount",
  "paragraphSplitColumnInsideGutter",
  "paragraphSplitColumnOutsideGutter",
];

/**
 * The protocol-64 grow rule for one section story (thoughts ADR 026/029):
 * generated pages after the story's last frame while it oversets, keeping
 * the section frame's options (`LeadingOffset`, zero insets, columns) —
 * Word's line-box fit on every page. The published plugin-api predates the
 * op, so it is typed here and cast at the door (as activate.ts does for
 * `storyContent`); drop the cast when the canary carrying it publishes.
 */
export function growRuleOp(storyId: string): Mutation {
  return {
    op: "setFlowGrowRule",
    args: { storyId, grow: true, maxPages: null, copyFrameOptions: true },
  } as unknown as Mutation;
}

/**
 * Set every section story's grow rule: one batch (one rebuild), op by op if
 * the batch is refused. An engine without the op (before protocol 64)
 * refuses it; the document stays open, it just does not grow, and the
 * refusal is reported (ADR-007). Returns the story ids that refused.
 */
export async function setGrowRules(host: BundleHost, storyIds: readonly string[]): Promise<string[]> {
  if (storyIds.length === 0) return [];
  const ops = storyIds.map(growRuleOp);
  const refused: string[] = [];
  let error: unknown = null;
  const whole = await host.document.mutate({ op: "batch", args: { ops } } as Mutation);
  if (!whole.applied) {
    for (let k = 0; k < ops.length; k++) {
      const one = await host.document.mutate(ops[k]);
      if (!one.applied) {
        refused.push(storyIds[k]);
        error ??= one.error;
      }
    }
  }
  host.diagnostics.set(
    OPEN_DIAGNOSTICS_KEY,
    refused.length === 0
      ? []
      : [
          {
            severity: "warning" as const,
            message:
              `This engine cannot grow pages (setFlowGrowRule, protocol 64): ` +
              `${refused.length} of ${storyIds.length} Word section(s) stay on one page, ` +
              `and text past it is overset. ${JSON.stringify(error)}`,
          },
        ],
  );
  return refused;
}

/** True when this host can open a native document (the standalone path). */
export function canOpenStandalone(host: BundleHost): boolean {
  return host.supports("document.openNative@1");
}

/**
 * Open `ir` (lowered from `source` by `engine`) as the active document.
 * Returns `null` when the host has no native-open door; the caller then
 * places the document embedded instead.
 */
export async function openStandalone(
  host: BundleHost,
  engine: DocEngine,
  ir: LoweredDoc,
  source: Uint8Array,
  name: string,
): Promise<Opened | null> {
  if (!canOpenStandalone(host)) return null;

  // 1. The skeleton: a page, frame and empty story per section (IDML).
  let storyIds = engine.skeletonStories();
  await host.nativeDocument.open(engine.skeletonIdml(name));

  // 2. Style catalog + swatches, once, before any applyStyle references them.
  //    They are also the probe for span/split columns: a Word section that
  //    changes the columns mid-page lowers to them (protocol 64). An engine
  //    that refuses them cannot lay that lowering out, and the plugin API
  //    has no flag that says so up front, so the document is lowered again
  //    with a page break at each column change (the lowering before span
  //    columns: right page content, an extra page per change) and reopened.
  let midPageColumns = engine.usesMidPageColumns();
  const refusals = await applyStyleOpsReporting(host, buildStyleMutations(ir));
  if (midPageColumns && SPAN_COLUMN_PATHS.some((p) => refusals.paths.has(p))) {
    engine.setMidPageColumns(false);
    midPageColumns = false;
    ir = engine.lowered();
    storyIds = engine.skeletonStories();
    await host.nativeDocument.open(engine.skeletonIdml(name));
    await applyStyleOpsReporting(host, buildStyleMutations(ir));
    host.diagnostics.set(COLUMNS_DIAGNOSTICS_KEY, [
      {
        severity: "warning" as const,
        message:
          "This engine cannot change columns mid-page (span/split columns, protocol 64): " +
          "each Word section that does opens a new page instead.",
      },
    ]);
  } else {
    host.diagnostics.set(COLUMNS_DIAGNOSTICS_KEY, []);
  }

  // 2b. The grow rules IDML cannot carry, before the pour oversets.
  await setGrowRules(host, storyIds);

  // 3. Each section's blocks into its own story. The engine grows each
  //    story's pages as the pour oversets its frame.
  const groups = sectionBlocks(ir);
  for (let k = 0; k < groups.length && k < storyIds.length; k++) {
    await pourSteps(host, buildStoryBlocks(groups[k], storyIds[k]));
  }

  // 4. Persist the source (the save-back baseline travels with the document)
  //    and bind it to the first section's frame.
  const frameId: ElementId = { kind: "textFrame", id: "docx_f0" };
  const partPath = `${storyIds[0] ?? "docx"}/source.docx`;
  try {
    await host.parts.write(partPath, source);
    await host.document.setMetadata(frameId, {
      v: 1,
      data: { part: partPath, blocks: ir.story.blocks.length, sections: storyIds },
    });
  } catch (err) {
    host.log.warn(`paged.doc: could not persist source part: ${String(err)}`);
  }
  return { storyIds, frameId, ir, midPageColumns };
}
