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

// Embedded placement — pour a lowered Word document into a NEW text frame on the
// current page, entirely through host.document.mutate. The engine + lowering are
// unit-tested (docx-lower / doc-host-model); this file is the thin host-driving
// glue. Live end-to-end placement is host-integration-verified in a later
// milestone (needs the editor + a wired NativeDocumentBackend); here it is
// written to the shipped contract and degrades honestly at each read door.

import type { LoweredDoc } from "@paged-media/doc-host-model";
import { buildDocumentMutations } from "@paged-media/doc-host-model";
import type { BundleHost, Diagnostic, ElementId, PageId } from "@paged-media/plugin-api";

/** The diagnostics key this plugin publishes under. */
export const DIAGNOSTICS_KEY = "media.paged.doc";

/** The plugin metadata namespace (the `x-paged:<id>` binding envelope). */
export const BINDING_KEY = "x-paged:media.paged.doc";

/** A frame's story id resolved via the hitTest read door. */
async function resolveStoryId(
  host: BundleHost,
  pageId: PageId,
  centre: [number, number],
): Promise<string | null> {
  const hit = await host.document.hitTest(pageId, centre);
  return hit?.storyId ?? null;
}

/**
 * Place `ir` as an embedded `wordDocument` object: create a text frame inside the
 * first page's margins, resolve its story, pour the content + styles as one
 * atomic batch, then stamp the binding + persist the source `.docx` as a part.
 * Returns the created frame id, or `null` if a read door was unavailable.
 */
export async function placeEmbedded(
  host: BundleHost,
  ir: LoweredDoc,
  source: Uint8Array,
): Promise<ElementId | null> {
  const pages = await host.document.collection<{ id: PageId }>("pages");
  const pageId = pages[0]?.id;
  if (!pageId) {
    host.log.warn("paged.doc: no page to place into");
    return null;
  }

  const s = ir.section;
  // Frame within the page margins. Bounds are [top, left, bottom, right] pts.
  const bounds: [number, number, number, number] = [
    s.marginTopPt,
    s.marginLeftPt,
    s.pageHeightPt - s.marginBottomPt,
    s.pageWidthPt - s.marginRightPt,
  ];

  const frameOutcome = await host.document.mutate({
    op: "insertTextFrame",
    args: { pageId, bounds },
  });
  if (!frameOutcome.applied || !frameOutcome.createdId) {
    host.log.warn("paged.doc: insertTextFrame was rejected by the host");
    return null;
  }
  const frameId = frameOutcome.createdId;

  const centre: [number, number] = [
    (bounds[1] + bounds[3]) / 2,
    (bounds[0] + bounds[2]) / 2,
  ];
  const storyId = await resolveStoryId(host, pageId, centre);
  if (!storyId) {
    host.log.warn(
      "paged.doc: could not resolve the frame's story (hitTest returned no story)",
    );
    return frameId;
  }

  // Style catalog + swatches + the pour, as one undo step.
  await host.document.mutate(buildDocumentMutations(ir, { storyId }));

  // Persist the source package (travels with the .paged file) + the binding.
  const partPath = `paged/media.paged.doc/${storyId}/source.docx`;
  try {
    await host.parts.write(partPath, source);
    await host.document.setMetadata(frameId, {
      v: 1,
      data: { part: partPath, blocks: ir.story.paragraphs.length },
    });
  } catch (err) {
    host.log.warn(`paged.doc: could not persist source part: ${String(err)}`);
  }

  // Surface honest diagnostics from the lowering (ADR-007).
  const diags: Diagnostic[] = ir.diagnostics.map((d) => ({
    severity: d.severity,
    message: `paged.doc: ${d.message}`,
  }));
  host.diagnostics.set(DIAGNOSTICS_KEY, diags);

  return frameId;
}
