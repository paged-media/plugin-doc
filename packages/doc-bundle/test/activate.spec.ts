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

// ADR 024 — what paged.doc DECLARES about the surface it needs.
//
// This bundle had no activation test at all, which is how its edit
// context went from 2026-07 to the context-sensitivity audit declaring
// a panel and nothing else. An undeclared surface is invisible to
// review and to CI alike; these assertions make the declaration a fact
// something checks.
//
// A recording host is the whole fixture: these assertions are about
// DECLARATIONS, so they need no browser, no canvas and no DOCX.

import { describe, expect, it } from "vitest";

import type {
  BundleHost,
  Disposable,
  EditContextContribution,
} from "@paged-media/plugin-api";

import * as docBundle from "../src/activate";

const PANEL_ID = "media.paged.doc.panel.outline";

/**
 * A recording host. Deliberately hand-rolled rather than the SDK's
 * `createHeadlessHost`, which refuses to boot without a resolvable
 * `@paged-media/canvas-wasm` — correctly, since a headless host with no
 * real engine is the fiction it exists to prevent. These assertions are
 * about DECLARATIONS and need no engine at all, so paying for a wasm
 * dependency to read a contribution back would be the wrong trade.
 * Same shape paged.sheet's activation spec uses.
 */
function fakeHost() {
  const editContexts: EditContextContribution[] = [];
  const openedPanels: string[] = [];
  const noop = (): Disposable => ({ dispose() {} });
  const host = {
    manifest: docBundle.manifest,
    log: { debug() {}, info() {}, warn() {}, error() {} },
    supports: (f: string) =>
      f === "contribute.editContext@1" ||
      f === "contribute.objectType@1" ||
      f === "contribute.importer@1" ||
      f === "contribute.exporter@1",
    contribute: {
      panel: noop,
      command: noop,
      importer: noop,
      exporter: noop,
      objectType: noop,
      editContext(c: EditContextContribution): Disposable {
        editContexts.push(c);
        return noop();
      },
    },
    shell: {
      openPanel(id: string) {
        openedPanels.push(id);
      },
      closePanel() {},
      async pickFile() {
        return [];
      },
    },
  } as unknown as BundleHost;
  return { host, editContexts, openedPanels };
}

function activated() {
  const fake = fakeHost();
  docBundle.activate(fake.host);
  return fake;
}

describe("paged.doc — the declared surface (ADR 024)", () => {
  it("enters by double-click, like every other content plugin", () => {
    const ctx = activated().editContexts;
    expect(ctx.map((c) => c.type)).toContain("wordDocument");
    // K-13 — one entry gesture across the product, now a one-member
    // union in the contract. Asserted rather than assumed because this
    // is the field a new content plugin is most likely to get wrong.
    expect(ctx[0]!.entry).toBe("doubleClick");
  });

  it("declares the HOST'S text tools — the case that proves the rule is not 'plugins get an empty rail'", () => {
    // paged.doc is the interesting one. A DOCX is lowered to NATIVE
    // content — real host text frames, real stories, real styles — so
    // the editor genuinely owns the caret and the host's own text tools
    // ARE the right tools here.
    //
    // That is the opposite of paged.sheet and paged.web, where no
    // canvas tool has anything to act on and the honest declaration is
    // empty. Same rule, different answer, which is why the rule is
    // "declare what applies" and not "declare nothing".
    const ctx = activated().editContexts[0]!;
    expect(ctx.toolIds, "toolIds is DECLARED").toBeDefined();
    expect(ctx.toolIds).toContain("paged.tool.type");
    expect(ctx.toolIds!.length, "and it is NOT the empty statement").toBeGreaterThan(0);
  });

  it("declares its panel and does NOT raise it by hand", () => {
    // The declaration is the only path that can WITHHOLD a raise which
    // would displace a shared panel this context's binding providers
    // serve (ADR 023). `host.shell.openPanel` takes no options and
    // always raises, so calling it from `onEnter` — as this context
    // used to — bypasses that rule.
    const fake = activated();
    const ctx = fake.editContexts[0]!;
    expect(ctx.panelIds).toContain(PANEL_ID);
    expect(
      ctx.onEnter,
      "no onEnter hook competing with the declaration",
    ).toBeUndefined();
  });
});
