#!/usr/bin/env node
// The contract-only import lint: every static import in this repo's TS source
// must come through the sanctioned plugin surface (the isolation-superset rule —
// only plugin-api, plugin-sdk, this repo's own packages, and react). No
// @paged-media/shell|client|ui|catalog, and no OTHER plugin's package.
//
// The second guarantee (CLAUDE.md hard rule, enforced by review not this lint):
// the TS side is thin glue — ALL DOCX semantics (OPC parse, WordprocessingML ->
// native lowering, preservation) live in the Rust crates (docx-js wasm).

import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, relative } from "node:path";
import process from "node:process";

const ROOT = new URL("..", import.meta.url).pathname;

const ALLOWED_PREFIXES = [
  "@paged-media/plugin-api",
  "@paged-media/plugin-sdk",
  "@paged-media/doc-", // this repo's own packages (doc-host-model)
  "@paged-media/doc", // the bundle self-reference, if any
  "react", // panels are React expert leaves (the v0 exception)
];

function walk(dir, out = []) {
  for (const name of readdirSync(dir)) {
    if (name === "node_modules" || name.startsWith(".")) continue;
    const path = join(dir, name);
    if (statSync(path).isDirectory()) walk(path, out);
    else if (/\.(ts|tsx)$/.test(name) && !/\.(spec|test)\./.test(name)) {
      out.push(path);
    }
  }
  return out;
}

const IMPORT = /(?:^|\n)\s*(?:import|export)[^"'`;]*?from\s*["']([^"']+)["']/g;

const violations = [];
for (const file of walk(join(ROOT, "packages"))) {
  if (!file.includes("/src/")) continue;
  const text = readFileSync(file, "utf8");
  IMPORT.lastIndex = 0;
  let m;
  while ((m = IMPORT.exec(text)) !== null) {
    const spec = m[1];
    if (spec.startsWith(".")) continue;
    if (ALLOWED_PREFIXES.some((p) => spec === p || spec.startsWith(`${p}/`) || spec.startsWith(p)))
      continue;
    violations.push(`${relative(ROOT, file)} → "${spec}"`);
  }
}

if (violations.length > 0) {
  console.error(
    "contract-import lint: imports outside the plugin surface " +
      "(promote to plugin-api / use an existing capability / record it):",
  );
  for (const v of violations) console.error(`  - ${v}`);
  process.exit(1);
}
console.log("contract-import lint: clean (plugin surface only)");
