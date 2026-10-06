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

// ADR 323 — the Word properties panel declares property rows (fields, not
// widgets); every row is a row of its kind's schema and the address forms
// are well-formed (plugin-sdk validatePanelSchema).

import { describe, expect, it } from "vitest";

import { propertyRowsOf, validatePanelSchema } from "@paged-media/plugin-sdk";

import manifest from "../manifest.json" with { type: "json" };
import { PANEL_FIELDS, schemaOf } from "../src/object-model";
import { DOC_PROPERTIES_PANEL, PROPERTIES_PANEL_ID } from "../src/panels/properties-panel";

const kindOf = (q: string) => q.slice(q.lastIndexOf("/") + 1) as Parameters<typeof schemaOf>[0];

describe("Word properties panel (property rows)", () => {
  it("validates against the manifest schemas", () => {
    expect(
      validatePanelSchema(DOC_PROPERTIES_PANEL, {
        pluginId: manifest.id,
        schemaOf: (k) => (k.startsWith(`plugin:${manifest.id}/`) ? schemaOf(kindOf(k)) : undefined),
      }),
    ).toEqual([]);
  });
  it("covers every declared PANEL_FIELDS entry, each through a property row", () => {
    const rows = propertyRowsOf(DOC_PROPERTIES_PANEL).map((r) => `${kindOf(r.field.kind!)}.${r.field.path}`);
    expect(rows.sort()).toEqual(PANEL_FIELDS.map((f) => `${f.kind}.${f.path}`).sort());
  });
  it("is declared in the manifest", () => {
    expect(manifest.contributes.panels).toContain(PROPERTIES_PANEL_ID);
  });
});
