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
// document, not content in a frame. The engine's skeleton (docx-skeleton)
// gives each Word section a page of its size, its margin box as the frame,
// and a story whose grow rule lets the engine add pages while it oversets,
// matching Word's own pagination (measured against Word's PDF export of
// docx_conformance::pagination_docx()). This file opens that skeleton and
// pours each section's blocks into its story with the shared pour.

import type { LoweredDoc } from "@paged-media/doc-host-model";
import { buildStoryBlocks, buildStyleMutations, sectionBlocks } from "@paged-media/doc-host-model";
import type { BundleHost, ElementId } from "@paged-media/plugin-api";

import type { DocEngine } from "./engine.js";
import { pourSteps } from "./pour.js";

/** What a standalone open produced: the section stories (save-back reads
 *  them back in this order) and the frame carrying the binding. */
export interface Opened {
  storyIds: string[];
  frameId: ElementId;
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

  // 1. The skeleton: a page, frame and growing story per section.
  const storyIds = engine.skeletonStories();
  await host.nativeDocument.open(engine.skeletonPaged(name));

  // 2. Style catalog + swatches, once, before any applyStyle references them.
  const styleOps = buildStyleMutations(ir);
  if (styleOps.length > 0) {
    await host.document.mutate({ op: "batch", args: { ops: styleOps } });
  }

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
  return { storyIds, frameId };
}
