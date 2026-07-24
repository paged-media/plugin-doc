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

// The typed engine FACADE + boot. The Rust wasm (docx-js) does ALL the DOCX
// work — OPC parse, WordprocessingML -> native lowering, preservation. This is
// a thin TS shape over its snake_case wasm-bindgen methods (CLAUDE.md hard rule:
// no DOCX semantics in TS).
//
// BOOT. The artifact is the wasm-bindgen `--target web` glue (bin/docx_js.js +
// bin/docx_js_bg.wasm, produced by scripts/build-wasm.sh). We DON'T use the
// host's loadBundleWasm (it instantiates a RAW module — no wbindgen imports); we
// load the glue in the bundle realm exactly like @paged-media/canvas-wasm does,
// branching browser vs Node like plugin-sdk's wasm-loader.ts. Until the artifact
// exists the dynamic import REJECTS — bootEngine surfaces that honestly.

import type { LoweredDoc } from "@paged-media/doc-host-model";
import { parseLoweredDoc } from "@paged-media/doc-host-model";

const ENGINE_NOT_BUILT =
  "paged.doc engine wasm not built — run scripts/build-wasm.sh";

/** The wasm-bindgen `DocEngine` (snake_case surface from docx-js/src/lib.rs). */
interface WasmDocEngine {
  load_docx(bytes: Uint8Array): void;
  lowered_json(): string;
  block_count(): number;
  save_verbatim(): Uint8Array;
  free(): void;
}

/** The `--target web` glue module shape. */
interface DocWasmModule {
  default(init: {
    module_or_path: string | URL | BufferSource | WebAssembly.Module;
  }): Promise<unknown>;
  initSync(module: { module: BufferSource | WebAssembly.Module }): unknown;
  DocEngine: new () => WasmDocEngine;
}

/** A stable camelCase facade the rest of the bundle codes against. */
export class DocEngine {
  private constructor(private readonly inner: WasmDocEngine) {}

  static async boot(): Promise<DocEngine> {
    const mod = await loadModule();
    return new DocEngine(new mod.DocEngine());
  }

  /** Load a `.docx`; throws with the engine's message on a hard failure. */
  loadDocx(bytes: Uint8Array): void {
    this.inner.load_docx(bytes);
  }

  /** The Tier-0 lowering (the host-model input). */
  lowered(): LoweredDoc {
    return parseLoweredDoc(this.inner.lowered_json());
  }

  /** Number of top-level body blocks. */
  blockCount(): number {
    return this.inner.block_count();
  }

  /** Zero-edit save-back (verbatim carry-through of the retained package). */
  saveVerbatim(): Uint8Array {
    return this.inner.save_verbatim();
  }

  dispose(): void {
    this.inner.free();
  }
}

function isNode(): boolean {
  return (
    typeof process !== "undefined" &&
    process.versions != null &&
    process.versions.node != null
  );
}

async function loadModule(): Promise<DocWasmModule> {
  let mod: DocWasmModule;
  try {
    // @ts-ignore — the artifact (bin/docx_js.js, the wasm-bindgen --target web
    // glue) is produced by scripts/build-wasm.sh and is intentionally absent
    // from the source tree; the dynamic import resolves at runtime once built.
    mod = (await import("../bin/docx_js.js")) as DocWasmModule;
  } catch (cause) {
    throw new Error(ENGINE_NOT_BUILT, { cause });
  }

  if (isNode()) {
    const { readFile } = await import("node:fs/promises");
    const { fileURLToPath } = await import("node:url");
    const wasmPath = fileURLToPath(
      new URL("../bin/docx_js_bg.wasm", import.meta.url),
    );
    const bytes = await readFile(wasmPath);
    mod.initSync({
      module: new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength),
    });
  } else {
    // @ts-ignore — `?url` is a bundler affordance, untyped (the editor's
    // wasm-loading convention; tsup marks `?url` external).
    const wasmUrl = (await import("../bin/docx_js_bg.wasm?url")) as {
      default: string;
    };
    await mod.default({ module_or_path: wasmUrl.default });
  }
  return mod;
}
