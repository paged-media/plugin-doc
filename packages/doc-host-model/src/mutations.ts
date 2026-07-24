// paged.doc — pure Lowered IR -> host `Mutation[]`.
//
// The same role as `sheet-host-model`: a *dumb* translator. Every id in the IR
// is already a fully-formed Paged token, so this file never invents ids — it only
// maps IR nodes to `host.document.mutate(...)` ops. No engine logic lives here;
// the semantics were decided in `docx-lower` (Rust).

import type { Mutation } from "@paged-media/plugin-api";

import type { LoweredDoc, LoweredStyle } from "./lowered.js";

/**
 * Count a string's length in Unicode scalar values (code points), matching the
 * `char`-offset convention of the engine's `InsertText`/`ApplyStyle` ops
 * (`Array.from` iterates code points, so astral characters count as one).
 */
function codePointLen(s: string): number {
  return Array.from(s).length;
}

/**
 * Swatch + style-catalog mutations. Emitted before any pour so `applyStyle`
 * references resolve. Swatches precede styles (a style's `characterFillColor`
 * references a swatch); styles are already topologically ordered by `docx-lower`
 * so every `basedOn` parent precedes its child.
 */
export function buildStyleMutations(ir: LoweredDoc): Mutation[] {
  const ops: Mutation[] = [];

  for (const sw of ir.swatches) {
    ops.push({
      op: "createSwatch",
      args: {
        spec: {
          selfId: sw.id,
          name: sw.name,
          space: sw.space,
          value: sw.value,
        },
      },
    } as Mutation);
  }

  for (const style of ir.styles) {
    ops.push(createStyleOp(style));
    for (const prop of style.props) {
      ops.push({
        op: "setStyleProperty",
        args: {
          collection: style.collection,
          styleId: style.id,
          path: prop.path,
          value: prop.value,
        },
      } as Mutation);
    }
  }

  return ops;
}

function createStyleOp(style: LoweredStyle): Mutation {
  const op =
    style.collection === "paragraph"
      ? "createParagraphStyle"
      : "createCharacterStyle";
  return {
    op,
    args: {
      selfId: style.id,
      name: style.name,
      basedOn: style.basedOn ?? null,
    },
  } as Mutation;
}

/**
 * Pour the story into the frame chain rooted at `storyId`: one `insertText` of
 * the whole body (paragraphs joined by `\n`), then `applyStyle` over each
 * paragraph range and each run range. Offsets are tracked in code points.
 *
 * `applyStyle` needs its target style to already exist, so the caller must apply
 * {@link buildStyleMutations} first (or bundle both via
 * {@link buildDocumentMutations}).
 */
export function buildPourMutations(ir: LoweredDoc, storyId: string): Mutation[] {
  const ops: Mutation[] = [];
  const paragraphs = ir.story.paragraphs;

  // 1. Assemble the full text and remember each paragraph/run range.
  interface Range {
    start: number;
    end: number;
    style: string;
    scope: "paragraph" | "character";
  }
  const ranges: Range[] = [];
  let text = "";
  let offset = 0;

  paragraphs.forEach((para, pIdx) => {
    const paraStart = offset;
    for (const run of para.runs) {
      const runStart = offset;
      text += run.text;
      offset += codePointLen(run.text);
      if (run.charStyleId) {
        ranges.push({
          start: runStart,
          end: offset,
          style: run.charStyleId,
          scope: "character",
        });
      }
    }
    const paraEnd = offset;
    if (para.paraStyleId) {
      ranges.push({
        start: paraStart,
        end: paraEnd,
        style: para.paraStyleId,
        scope: "paragraph",
      });
    }
    // Paragraph separator (not after the final paragraph).
    if (pIdx < paragraphs.length - 1) {
      text += "\n";
      offset += 1;
    }
  });

  if (text.length > 0) {
    ops.push({
      op: "insertText",
      args: { storyId, offset: 0, text, cell: null },
    } as Mutation);
  }

  // 2. Apply paragraph styles first, then character styles (so a run's direct
  //    formatting wins over its paragraph's).
  for (const r of ranges.filter((r) => r.scope === "paragraph")) {
    ops.push(applyStyleOp(storyId, r.start, r.end, r.style, "paragraph"));
  }
  for (const r of ranges.filter((r) => r.scope === "character")) {
    ops.push(applyStyleOp(storyId, r.start, r.end, r.style, "character"));
  }

  return ops;
}

function applyStyleOp(
  storyId: string,
  start: number,
  end: number,
  style: string,
  scope: "paragraph" | "character",
): Mutation {
  return {
    op: "applyStyle",
    args: { storyId, start, end, style, scope },
  } as Mutation;
}

/**
 * Everything needed to realize the lowering into the story `storyId`, as one
 * atomic `batch` (a single undo step): style catalog + swatches, then the pour.
 */
export function buildDocumentMutations(
  ir: LoweredDoc,
  opts: { storyId: string },
): Mutation {
  const ops: Mutation[] = [
    ...buildStyleMutations(ir),
    ...buildPourMutations(ir, opts.storyId),
  ];
  return { op: "batch", args: { ops } } as Mutation;
}
