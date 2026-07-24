import { defineConfig } from "tsup";

// Bundle the pure host-model INTO the published @paged-media/doc artifact (it is
// a workspace-private dep), and keep `?url` wasm imports external (a bundler
// affordance the consuming app resolves).
export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  noExternal: [/^@paged-media\/doc-host-model/],
  external: [/\?url$/, "react", "react/jsx-runtime"],
});
