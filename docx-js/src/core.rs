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

//! The plain-Rust engine core behind the wasm surface.
//!
//! `DocSession` holds a loaded `.docx` (its retained bytes for the preservation
//! invariant, its parsed semantic model, the lowering, and — for M2 edited
//! save-back — the retained OPC package + the native↔OOXML provenance bindings).
//! All real work lives here so it is native-testable; `lib.rs`'s
//! `#[wasm_bindgen]` layer is a pure forwarding shim.

use docx_core::DocxDocument;
use docx_export::{
    apply_edits, build_bindings, diff, overlay_story_content, DocxBindings, EditSet, StoryContentIn,
};
use docx_import::import_docx_with_package;
use docx_lower::ir::LoweredDoc;
use docx_lower::lower;
use paged_ooxml::OpcPackage;

/// A loaded Word document session.
pub struct DocSession {
    /// The original `.docx` bytes, retained for the preservation invariant
    /// (zero-edit save-back re-emits them verbatim).
    source: Vec<u8>,
    /// The parsed semantic model.
    model: DocxDocument,
    /// The retained OPC package — the save-back patch target (`set_part` +
    /// `write`, untouched parts re-emitted byte-identical).
    package: OpcPackage,
    /// The main document part name (e.g. `word/document.xml`).
    main_part: String,
    /// The native↔OOXML provenance map for edited save-back.
    bindings: DocxBindings,
}

impl DocSession {
    /// Load a `.docx`/`.dotx` package. Returns a human-readable error string on a
    /// hard failure (unreadable container / unparseable main document).
    pub fn load(bytes: &[u8]) -> Result<DocSession, String> {
        let (model, package, main_part) =
            import_docx_with_package(bytes).map_err(|e| e.to_string())?;
        let bindings = build_bindings(&model);
        Ok(DocSession {
            source: bytes.to_vec(),
            model,
            package,
            main_part,
            bindings,
        })
    }

    /// The Tier-0 lowering (the IR the host-model turns into mutations).
    pub fn lowered(&self) -> LoweredDoc {
        lower(&self.model)
    }

    /// The lowering as a JSON string (the wasm boundary form).
    pub fn lowered_json(&self) -> String {
        serde_json::to_string(&self.lowered()).unwrap_or_else(|_| "null".to_string())
    }

    /// The parsed semantic model (exposed for tests / debugging).
    pub fn model(&self) -> &DocxDocument {
        &self.model
    }

    /// Number of top-level body blocks.
    pub fn block_count(&self) -> usize {
        self.model.body.len()
    }

    /// Zero-edit save-back: re-emit the retained source verbatim.
    pub fn save_verbatim(&self) -> Vec<u8> {
        self.source.clone()
    }

    /// M2 edited save-back: patch the given run edits into the retained package
    /// (only the changed `<w:t>`/`<w:rPr>` subtrees rewritten; every other part +
    /// untouched subtree byte-identical) and return the saved `.docx` bytes.
    /// Non-patchable edits are skipped. A cheap clone of the package keeps the
    /// session reusable for a subsequent save.
    ///
    /// DEFERRED (RFI DOC-03): the `EditSet` here is supplied by the caller
    /// (tests today). Wiring it from the LIVE editor needs a structured
    /// whole-document read door — `host.nativeDocument.readModel()` returns
    /// opaque core-native bytes this isolation-clean plugin cannot diff.
    pub fn save_edited(&self, edits: &EditSet) -> Result<Vec<u8>, String> {
        let mut package = self.package.clone();
        apply_edits(&mut package, &self.main_part, &self.bindings, edits)
            .map_err(|e| e.to_string())?;
        package.write().map_err(|e| e.to_string())
    }

    /// M2 edited save-back, driven by the DOC-03 structured read: overlay the
    /// host's read-back `StoryContent` onto the import baseline lowering, diff to
    /// an `EditSet`, and save. This is the LIVE end-to-end path (the bundle reads
    /// `host.document.storyContent(storyId)` and forwards it here) — the seam that
    /// makes edited save-back run without a hand-authored `EditSet`.
    pub fn save_edited_from_content(&self, content: &StoryContentIn) -> Result<Vec<u8>, String> {
        let baseline = self.lowered();
        let edited = overlay_story_content(&baseline, content);
        let edits = diff(&baseline, &edited, &self.bindings);
        self.save_edited(&edits)
    }
}
