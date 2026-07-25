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

import type { BundleHandle, BundleHost } from "@paged-media/plugin-api";

import { DocEngine } from "./engine.js";
import { makeOutlinePanel } from "./panels/outline-panel.js";
import { placeEmbedded } from "./place.js";

const PANEL_ID = "media.paged.doc.panel.outline";
const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** The last-imported source, so the exporter can re-emit it verbatim (Tier-0
 *  zero-edit save-back; edited save-back is M2). */
interface LastDoc {
  fileName: string;
  source: Uint8Array;
}

export function activate(host: BundleHost): BundleHandle {
  let last: LastDoc | null = null;
  const disposers: Array<() => void> = [];

  // Load bytes -> engine -> lowering -> embedded placement.
  async function ingest(name: string, bytes: Uint8Array): Promise<void> {
    const engine = await DocEngine.boot();
    try {
      engine.loadDocx(bytes);
      const ir = engine.lowered();
      last = { fileName: name, source: bytes };
      await placeEmbedded(host, ir, bytes);
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
      ...makeOutlinePanel(() => last?.fileName ?? null, () => void pickAndIngest()),
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
        // Zero-edit passthrough. The M2 edited-save-back ENGINE exists and is
        // proven (docx-js `save_edited` byte-splices a targeted patch; see
        // docx-conformance tests/save_back.rs), but wiring it here is DEFERRED
        // (RFI DOC-03): the exporter hook gets no document handle and
        // `host.nativeDocument.readModel()` returns opaque core-native bytes this
        // isolation-clean plugin cannot diff into an EditSet. When a structured
        // whole-document read door lands, diff the edited LoweredDoc against the
        // import baseline → EditSet → `engine.save_edited(...)`.
        export: () =>
          last ? { bytes: last.source, fileName: last.fileName } : null,
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
