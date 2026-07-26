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
import { buildStory, buildStyleMutations } from "@paged-media/doc-host-model";
import type { BundleHost, Diagnostic, ElementId, PageId } from "@paged-media/plugin-api";

/** A conservative story-offset advance past a table paragraph. The exact
 *  footprint is refined during editor integration (a table occupies one
 *  paragraph position in the story). */
/**
 * A table's footprint in the STYLE-range (contiguous character) space is ZERO:
 * `insertTable` pushes a host paragraph carrying the table with NO runs
 * (`Paragraph { table: Some(..), ..Default::default() }` in core's
 * `apply_insert_table`), and the contiguous space counts only `CharacterRun`
 * text. In the `insertText` address space it is 1 — that space counts a synthetic
 * inter-paragraph break, which the host paragraph still has.
 *
 * These are two DIFFERENT address spaces (see docs/status.md, "Offset
 * convention"), so they are tracked separately below.
 */
const TABLE_FOOTPRINT_STYLE = 0;
const TABLE_FOOTPRINT_TEXT = 1;

/** Extract the bare table id string from an `insertTable` outcome's createdId. */
function tableIdOf(id: ElementId | null): string | null {
  if (id && id.kind === "table") return id.id.table_id;
  return null;
}

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
/** What a successful placement produced: the host frame and (when the frame's
 *  story resolved) its story id — the address the DOC-03 read-back uses to pull
 *  the EDITED content for save-back. */
export interface PlacedDoc {
  frameId: ElementId;
  storyId: string | null;
}

export async function placeEmbedded(
  host: BundleHost,
  ir: LoweredDoc,
  source: Uint8Array,
): Promise<PlacedDoc | null> {
  // The pages collection carries `selfId`, NOT `id` — reading `.id` yielded
  // undefined on every real host, so placement always bailed with "no page to
  // place into" and nothing was ever inserted. It went unnoticed because the
  // Rust tests drive the pour directly and the bundle had never run in the
  // editor. Prefer the ACTIVE page (a designer places into the page they are
  // looking at), falling back to the first — the same resolution paged.web and
  // paged.sheet use.
  const meta = await host.document.meta();
  const pages = await host.document.collection<{ selfId: string }>("pages");
  const pageId: PageId | null = meta.activePage ?? pages[0]?.selfId ?? null;
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
    return { frameId, storyId: null };
  }

  // 1. Style catalog + swatches (must exist before applyStyle references them).
  const styleOps = buildStyleMutations(ir);
  if (styleOps.length > 0) {
    await host.document.mutate({ op: "batch", args: { ops: styleOps } });
  }

  // 2. Walk the story plan in order: text runs pour at the running offset;
  //    tables insert (the outcome mints the id) then pour their cells.
  // `styleOffset` addresses the CONTIGUOUS character space (applyStyle /
  // insertAnchoredFrame / insertHyperlink ranges); `textOffset` addresses
  // insertText's byte+synthetic-break space. They diverge as soon as the story
  // contains a table, which is why they are no longer one counter.
  let styleOffset = 0;
  let textOffset = 0;
  for (const step of buildStory(ir, storyId)) {
    if (step.kind === "text") {
      // SEQUENTIAL, not one `batch` — see RFI DOC-04. The engine keeps text
      // edits in a different lane from element edits: `Mutation::Batch`
      // translates its children to `paged_mutate::Operation`, and InsertText /
      // DeleteRange have no Operation form (they route through paged-canvas's
      // own TextOp). So a batch carrying any text op fails translation and the
      // whole thing is rejected with `notImplemented: Mutation::Batch` — which
      // is exactly what the first live run of this plugin hit: a frame
      // appeared, empty, and nothing said why.
      //
      // Cost of pouring op-by-op: the placement is no longer ONE undo step and
      // is not atomic, so a mid-pour rejection leaves a partly-poured frame.
      // We report that rather than hide it. Restore the batch when the engine
      // can carry text ops in one.
      for (const op of step.mutations(textOffset, styleOffset)) {
        const poured = await host.document.mutate(op);
        // A rejected pour used to pass unnoticed: the frame still appeared, so
        // the document looked "placed" while being empty. Say so (ADR-007 —
        // never a silent drop).
        if (!poured.applied) {
          host.log.warn(
            `paged.doc: the engine rejected a pour op (${
              (op as { op?: string }).op ?? "?"
            }): ${JSON.stringify(poured.error)}`,
          );
        }
      }
      styleOffset += step.length;
      textOffset += step.byteLength;
    } else {
      const outcome = await host.document.mutate(step.insert);
      const tableId = outcome.applied ? tableIdOf(outcome.createdId) : null;
      if (tableId) {
        await host.document.mutate(step.cells(tableId));
      }
      styleOffset += TABLE_FOOTPRINT_STYLE;
      textOffset += TABLE_FOOTPRINT_TEXT;
    }
  }

  // Persist the source package (travels with the .paged file) + the binding.
  const partPath = `paged/media.paged.doc/${storyId}/source.docx`;
  try {
    await host.parts.write(partPath, source);
    await host.document.setMetadata(frameId, {
      v: 1,
      data: { part: partPath, blocks: ir.story.blocks.length },
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

  return { frameId, storyId };
}
