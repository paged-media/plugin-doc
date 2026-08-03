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

// activate(host) for paged.doc — thin glue: engine boot, importer/exporter,
// objectType + editContext, an insert command, and an outline panel. All DOCX
// semantics live in the docx-js wasm; this file only drives host surfaces.

import type {
  BundleHandle,
  BundleHost,
  ElementId,
} from "@paged-media/plugin-api";

import { DocEngine } from "./engine.js";
import { createDocStore } from "./panels/outline-model.js";
import { makeOutlinePanel } from "./panels/outline-panel.js";
import { placeEmbedded } from "./place.js";

const PANEL_ID = "media.paged.doc.panel.outline";
const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** The last-imported document: its source bytes (the preservation carrier the
 *  exporter re-emits) plus the story it was poured into — the address the DOC-03
 *  read-back uses to pull the EDITED content for save-back. */
interface LastDoc {
  fileName: string;
  source: Uint8Array;
  storyId: string | null;
}

export function activate(host: BundleHost): BundleHandle {
  let last: LastDoc | null = null;
  const disposers: Array<() => void> = [];
  // The panel's document store: the RETAINED LoweredDoc (plain JSON — the
  // engine itself is disposed after placement) + the placement record.
  const docStore = createDocStore();

  // Load bytes -> engine -> lowering -> embedded placement.
  async function ingest(name: string, bytes: Uint8Array): Promise<void> {
    const engine = await DocEngine.boot();
    try {
      engine.loadDocx(bytes);
      const ir = engine.lowered();
      const placed = await placeEmbedded(host, ir, bytes);
      last = { fileName: name, source: bytes, storyId: placed?.storyId ?? null };
      docStore.set({
        fileName: name,
        ir,
        frameId: placed?.frameId ?? null,
        storyId: placed?.storyId ?? null,
      });
      host.shell.openPanel(PANEL_ID);
      // Standalone "open as the whole canvas" needs the docx->native-bytes
      // producer (deferred); when that + host.nativeDocument.open are wired we
      // switch File/Open to a true open. Until then this is embedded placement.
      if (!host.supports("document.openNative@1")) {
        host.log.info(
          "paged.doc: host has no openNative door — placed as embedded content",
        );
      }
    } finally {
      engine.dispose();
    }
  }

  /**
   * Produce the bytes to save. M2 edited save-back (DOC-03): when the host has
   * the structured story read, pull the EDITED story back, let the engine overlay
   * it on the import baseline, diff, and write a TARGETED patch (only changed
   * `w:t`/`w:rPr` rewritten; every other part + untouched subtree byte-identical).
   *
   * Degrades honestly at every step — no read door, no story id, or a failing
   * save all fall back to the verbatim source rather than exporting something
   * wrong. (A host without the v54 read reports `readStory@1` false; the SDK's
   * reserved member is never called.)
   */
  async function exportDocx(): Promise<{
    bytes: Uint8Array;
    fileName: string;
  } | null> {
    if (!last) return null;
    const verbatim = { bytes: last.source, fileName: last.fileName };
    if (!host.supports("document.readStory@1") || !last.storyId) {
      return verbatim;
    }
    try {
      // `storyContent` is a v54 read; it postdates the PUBLISHED plugin-api
      // canary this bundle builds against (it exists on plugin-sdk main), so
      // reach it through a narrow cast — the same pattern the v52/v53 mutation
      // ops use. Drop the cast when the canary carrying it publishes. Guarded by
      // the `readStory@1` probe above, so an older host never reaches this.
      const readStory = (
        host.document as unknown as {
          storyContent(storyId: string): Promise<unknown | null>;
        }
      ).storyContent;
      const content = await readStory.call(host.document, last.storyId);
      if (!content) return verbatim;
      const engine = await DocEngine.boot();
      try {
        engine.loadDocx(last.source);
        const bytes = engine.saveEditedFromContent(content);
        // Save-back refusal feedback (ADR-007 posture): the patcher's skip
        // ledger — edits refused rather than risking corruption (gridSpan
        // column ops, non-patchable runs) — reaches the Problems panel
        // instead of dying silently in the engine.
        const skips = engine.lastSaveSkips();
        host.diagnostics.set(
          "media.paged.doc/save-back",
          skips.map((message) => ({ severity: "warning" as const, message })),
        );
        return { bytes, fileName: last.fileName };
      } finally {
        engine.dispose();
      }
    } catch (err) {
      host.log.warn(
        `paged.doc: edited save-back failed, exporting the unedited source: ${String(err)}`,
      );
      return verbatim;
    }
  }

  // Pick a .docx and place it (shared by the command + the panel button).
  async function pickAndIngest(): Promise<void> {
    if (!host.supports("shell.pickFile@1")) {
      host.log.warn("paged.doc: this host cannot pick files");
      return;
    }
    const picked = await host.shell.pickFile({ accept: [".docx", ".dotx"] });
    const file = picked[0];
    if (file) await ingest(file.name, file.bytes);
  }

  disposers.push(
    host.contribute.panel({
      id: PANEL_ID,
      ...makeOutlinePanel(
        docStore,
        () => void pickAndIngest(),
        (frameId) => void host.selection.set([frameId as ElementId]),
        () => ({ readStory: host.supports("document.readStory@1") }),
      ),
    }).dispose,
  );

  // Command: pick + place a .docx into the current document.
  disposers.push(
    host.contribute.command({
      id: "media.paged.doc.command.placeDoc",
      title: "Place Word document…",
      handler: () => void pickAndIngest(),
    }).dispose,
  );

  // Importer: File/Open + drag-drop of a .docx routes its bytes here.
  if (host.supports("contribute.importer@1")) {
    disposers.push(
      host.contribute.importer({
        id: "media.paged.doc.importer.docx",
        title: "Word document (.docx)",
        extensions: [".docx", ".dotx"],
        mimeTypes: [DOCX_MIME],
        import: ({ name, bytes }) => ingest(name, bytes),
      }).dispose,
    );
  }

  // Exporter: zero-edit save-back (verbatim carry-through of the retained OPC).
  if (host.supports("contribute.exporter@1")) {
    disposers.push(
      host.contribute.exporter({
        id: "media.paged.doc.exporter.docx",
        title: "Word document (.docx)",
        extension: ".docx",
        mimeType: DOCX_MIME,
        export: () => exportDocx(),
      }).dispose,
    );
  }

  // objectType: a frame carrying our binding envelope is a wordDocument.
  if (host.supports("contribute.objectType@1")) {
    disposers.push(
      host.contribute.objectType({
        type: "wordDocument",
        bakedFallback: "group",
        editContextType: "wordDocument",
        matches: (candidate) =>
          candidate.metadata != null &&
          typeof candidate.metadata.data?.part === "string",
      }).dispose,
    );
  }

  // editContext: double-click enters native text editing scoped to the frame;
  // the editor owns the caret (native text tools — this sidesteps C-9), so the
  // context mainly opens the doc panel and bounds the edit scope.
  if (host.supports("contribute.editContext@1")) {
    disposers.push(
      host.contribute.editContext({
        type: "wordDocument",
        entry: "doubleClick",
        panelIds: [PANEL_ID],
        onEnter: () => host.shell.openPanel(PANEL_ID),
      }).dispose,
    );
  }

  host.log.info("paged.doc activated");
  return {
    dispose() {
      for (const d of disposers) d();
    },
  };
}
