// paged.doc host-model — pure Lowered IR -> host Mutation[] translation.
//
// The only @paged-media dependency is plugin-api (types). No DOM, no engine
// logic: the bundle boots the wasm engine, gets the Lowered IR JSON, and this
// package turns it into mutations the host applies via host.document.mutate.

export * from "./lowered.js";
export * from "./mutations.js";
