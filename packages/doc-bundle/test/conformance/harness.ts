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

// The headless lane for paged.doc: the REAL engine (canvas-wasm through the
// SDK's createHeadlessHost) and the REAL docx-js wasm, the bundle activated
// as the editor activates it. A Word document is opened the way the
// standalone open does it — the engine's page skeleton loaded as the
// document, then the bundle's own pour, styles and label (`finishStandalone`)
// — because the headless host has no native-open door (the skeleton load IS
// what that door does).
//
// The docx-js artifact is built by scripts/build-wasm.sh (gitignored).
// Without it these suites skip with a reason; REQUIRE_REAL_DOCX=1 (CI after
// the wasm build) makes a missing artifact a failure.

import { existsSync, readFileSync } from "node:fs";

import type { BundleHost } from "@paged-media/plugin-api";
import { createHeadlessHost, type HeadlessHost } from "@paged-media/plugin-sdk";

import { docBundle } from "../../src";
import { DocEngine } from "../../src/engine";
import { finishStandalone, type Opened } from "../../src/open";

const WASM = new URL("../../bin/docx_js_bg.wasm", import.meta.url);

export const docxPresent = existsSync(WASM);

/** Fail (not skip) when the lane is required and the artifact is missing. */
export function requireDocx(suite: string): void {
  if (!docxPresent && process.env.REQUIRE_REAL_DOCX === "1") {
    throw new Error(`${suite}: REQUIRE_REAL_DOCX=1 but bin/docx_js_bg.wasm is not built (bash scripts/build-wasm.sh)`);
  }
}

export const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/${name}`, import.meta.url)));

const silent = { debug() {}, info() {}, warn() {}, error() {} };

export interface OpenedDoc {
  h: HeadlessHost;
  host: BundleHost;
  opened: Opened;
}

/** A headless host with the bundle loaded and `name` opened standalone. */
export async function openDoc(name: string): Promise<OpenedDoc> {
  const bytes = fixture(name);
  const h = await createHeadlessHost({ console: silent });
  const engine = await DocEngine.boot();
  try {
    engine.loadDocx(bytes);
    const ir = engine.lowered();
    await h.load(engine.skeletonIdml(name));
    h.loadBundle(docBundle);
    const host = h.host;
    const opened = await finishStandalone(host, engine, ir, bytes, name, (idml) => h.load(idml));
    return { h, host, opened };
  } finally {
    engine.dispose();
  }
}

/** The text of every paragraph of a story, as the engine holds it. */
export async function storyTexts(host: BundleHost, storyId: string): Promise<string[]> {
  const c = await host.document.storyContent(storyId);
  return (c?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join(""));
}
