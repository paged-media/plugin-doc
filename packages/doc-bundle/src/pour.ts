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

// The story pour, shared by embedded placement (place.ts) and standalone open
// (open.ts, one call per Word section story — ADR 029).

import type { StoryStep } from "@paged-media/doc-host-model";
import type { BundleHost, ElementId, Mutation } from "@paged-media/plugin-api";

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

/** Execute a story plan against the host, in order. */
export async function pourSteps(host: BundleHost, steps: readonly StoryStep[]): Promise<void> {
  // Walk the story plan in order: text runs pour at the running offset;
  //    tables insert (the outcome mints the id) then pour their cells.
  // `styleOffset` addresses the CONTIGUOUS character space (applyStyle /
  // insertAnchoredFrame / insertHyperlink ranges); `textOffset` addresses
  // insertText's byte+synthetic-break space. They diverge as soon as the story
  // contains a table, which is why they are no longer one counter.
  let styleOffset = 0;
  let textOffset = 0;
  // C-14 ADOPTED (core PR #46): `Mutation::Batch` can carry text ops on new
  // engines, so each text step pours as ONE atomic batch (one undo step per
  // step). An OLD engine rejects the whole batch (`notImplemented:
  // Mutation::Batch` — the same wire, so no capability flag distinguishes
  // them); the rejection applies NOTHING, which makes try-then-fall-back a
  // safe runtime probe. The verdict is cached for the rest of the placement.
  // (One batch for the WHOLE story stays impossible: table inserts mint ids
  // mid-pour that their cell pours need.)
  let engineBatchesText: boolean | null = null;
  for (const step of steps) {
    if (step.kind === "text") {
      const ops = step.mutations(textOffset, styleOffset);
      let sequential = engineBatchesText === false || ops.length <= 1;
      if (!sequential) {
        const outcome = await host.document.mutate({ op: "batch", args: { ops } });
        if (outcome.applied) {
          engineBatchesText = true;
        } else {
          // Pre-C-14 engine (or another whole-batch rejection): nothing was
          // applied, so the sequential pour below replays the SAME ops.
          engineBatchesText = false;
          sequential = true;
        }
      }
      if (sequential) {
        // The pre-C-14 lane (see RFI DOC-04): op-by-op, not atomic — a
        // mid-pour rejection leaves a partly-poured frame. We report that
        // rather than hide it (ADR-007 — never a silent drop).
        for (const op of ops) {
          const poured = await host.document.mutate(op);
          if (!poured.applied) {
            host.log.warn(
              `paged.doc: the engine rejected a pour op (${
                (op as { op?: string }).op ?? "?"
              }): ${JSON.stringify(poured.error)}`,
            );
          }
        }
      }
      styleOffset += step.length;
      textOffset += step.byteLength;
    } else {
      const outcome = await host.document.mutate(step.insert);
      const tableId = outcome.applied ? tableIdOf(outcome.createdId) : null;
      if (tableId) {
        // The cells are one batch: a refusal loses the whole table's text,
        // which must be said, not swallowed (ADR-007).
        const cells = await host.document.mutate(step.cells(tableId));
        if (!cells.applied) {
          host.log.warn(
            `paged.doc: the engine rejected a table's cell content: ${JSON.stringify(cells.error)}`,
          );
        }
      } else {
        host.log.warn(
          `paged.doc: the engine did not create a table: ${JSON.stringify(
            outcome.applied ? "no table id in the reply" : outcome.error,
          )}`,
        );
      }
      styleOffset += TABLE_FOOTPRINT_STYLE;
      textOffset += TABLE_FOOTPRINT_TEXT;
    }
  }

}

/** Diagnostics key for style properties the engine refused. */
export const STYLE_DIAGNOSTICS_KEY = "media.paged.doc/styles";

/**
 * Apply the style catalog. One batch first (one undo step). A batch is
 * all-or-nothing, so a single property the engine cannot set (an older
 * engine refuses many paragraph paths at style level, and the break-before
 * rule needs protocol 64) would cost the document EVERY style. On a refused
 * batch, apply the ops one by one instead, and report each refusal as a
 * warning rather than lose the rest silently (ADR-007).
 */
export async function applyStyleOps(
  host: BundleHost,
  ops: readonly Mutation[],
): Promise<string[]> {
  return (await applyStyleOpsReporting(host, ops)).messages;
}

/** What [`applyStyleOpsReporting`] found the engine refuses. */
export interface StyleRefusals {
  /** One warning per refused op (also set as diagnostics). */
  messages: string[];
  /** The property paths of the refused `setStyleProperty` ops. */
  paths: Set<string>;
}

/** [`applyStyleOps`], also saying WHICH property paths were refused (the
 *  standalone open learns from them whether the engine has span/split
 *  columns, ADR 029). */
export async function applyStyleOpsReporting(
  host: BundleHost,
  ops: readonly Mutation[],
): Promise<StyleRefusals> {
  const out: StyleRefusals = { messages: [], paths: new Set() };
  if (ops.length === 0) return out;
  const whole = await host.document.mutate({ op: "batch", args: { ops: [...ops] } } as Mutation);
  if (whole.applied) {
    host.diagnostics.set(STYLE_DIAGNOSTICS_KEY, []);
    return out;
  }
  for (const op of ops) {
    const one = await host.document.mutate(op);
    if (!one.applied) {
      const args = (op as { args?: { styleId?: string; path?: string } }).args;
      const what = args?.path ? `${args.path} on ${args.styleId ?? "?"}` : (op as { op?: string }).op ?? "?";
      out.messages.push(`This engine cannot apply ${what}: ${JSON.stringify(one.error)}`);
      if (args?.path) out.paths.add(args.path);
    }
  }
  host.diagnostics.set(
    STYLE_DIAGNOSTICS_KEY,
    out.messages.map((message) => ({ severity: "warning" as const, message })),
  );
  return out;
}

