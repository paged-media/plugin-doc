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

// ADR 323 (object-model design §4 "Panels") — the Word document's object
// properties as a SCHEMA panel of property rows. The host renders each row
// with its schema-driven `PropertyField` over `host.objects` (widget from
// the docxDocument / block / run / mergeField schema row, one undo step per
// commit, "Bind to data…" on every field). The bundle declares fields, never
// widgets; the outline panel stays as it is.

import type { BundleHost, Disposable, PanelSchema } from "@paged-media/plugin-api";
import { propertyRows, publishObjectList } from "@paged-media/plugin-sdk";

import { PANEL_FIELDS, PLUGIN_ID } from "../object-model.js";

export const PROPERTIES_PANEL_ID = "media.paged.doc.panel.properties";

const DOCUMENT = `plugin:${PLUGIN_ID}/docxDocument`;
const BLOCK = `plugin:${PLUGIN_ID}/block`;
const RUN = `plugin:${PLUGIN_ID}/run`;
const MERGE_FIELD = `plugin:${PLUGIN_ID}/mergeField`;

/** Published bindings the lists and their property rows share. */
export const BIND = {
  mergeFields: "doc.properties.mergeFields",
  mergeField: "doc.properties.mergeField",
  blocks: "doc.properties.blocks",
  block: "doc.properties.block",
  runs: "doc.properties.runs",
  run: "doc.properties.run",
} as const;

const paths = (kind: string, group: string) =>
  PANEL_FIELDS.filter((f) => `plugin:${PLUGIN_ID}/${f.kind}` === kind && f.group === group).map((f) => f.path);

const list = (rows: string, select: string) => ({
  widget: "paged.list",
  list: {
    items: { kind: "binding" as const, bind: rows },
    labelField: "name",
    secondaryField: "secondary",
    selectionBinding: select,
  },
});

/** The panel, from `PANEL_FIELDS` (every entry is a row of its kind). */
export const DOC_PROPERTIES_PANEL: PanelSchema = {
  id: PROPERTIES_PANEL_ID,
  title: "Word properties",
  icon: "panel-canvas",
  defaultDock: "right",
  sections: [
    {
      title: "Document",
      // The selected frame's Word document (`<selected> > docxDocument`).
      rows: propertyRows(DOCUMENT, paths(DOCUMENT, "Document"), "selection"),
    },
    {
      title: "Mail merge",
      rows: [
        ...propertyRows(DOCUMENT, paths(DOCUMENT, "Mail merge"), "selection"),
        list(BIND.mergeFields, BIND.mergeField),
        ...propertyRows(MERGE_FIELD, paths(MERGE_FIELD, "Mail merge"), { bind: BIND.mergeField }),
      ],
    },
    {
      title: "Block",
      collapsible: true,
      rows: [
        list(BIND.blocks, BIND.block),
        ...propertyRows(BLOCK, paths(BLOCK, "Block"), { bind: BIND.block }),
      ],
    },
    {
      title: "Run",
      collapsible: true,
      rows: [list(BIND.runs, BIND.run), ...propertyRows(RUN, paths(RUN, "Run"), { bind: BIND.run })],
    },
  ],
};

/** Keep the lists live: the merge fields and blocks of the selected Word
 *  frame, and the runs of the picked block. */
export function publishPropertyLists(host: BundleHost): Disposable {
  const one = (sel: readonly string[]) => (sel.length === 1 ? sel[0]! : null);
  const lists = [
    publishObjectList(host, {
      rows: BIND.mergeFields,
      select: BIND.mergeField,
      selector: (sel) => (one(sel) ? `${one(sel)} > ${MERGE_FIELD}` : null),
      secondaryPath: "value",
    }),
    publishObjectList(host, {
      rows: BIND.blocks,
      select: BIND.block,
      selector: (sel) => (one(sel) ? `${one(sel)} > ${BLOCK}` : null),
      labelPath: "text",
      secondaryPath: "kind",
    }),
  ];
  // Runs follow the picked block (a published binding, not the selection).
  const runs = publishObjectList(host, {
    rows: BIND.runs,
    select: BIND.run,
    selector: () => {
      const block = host.bindings.get(BIND.block);
      return typeof block === "string" && block ? `${block} > ${RUN}` : null;
    },
    labelPath: "text",
    secondaryPath: "mergeField",
  });
  const sub = host.bindings.onDidChange((name) => {
    if (name === BIND.block) void runs.refresh();
  });
  return {
    dispose: () => {
      sub.dispose();
      runs.dispose();
      for (const l of lists) l.dispose();
    },
  };
}
