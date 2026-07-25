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

//! # docx-js — the single wasm-bindgen surface for paged.doc
//!
//! Two layers: [`core::DocSession`] is plain Rust (all logic, native-testable);
//! the `#[wasm_bindgen] DocEngine` below is a pure forwarding shim — nothing
//! computes here. The bundle boots this module (`--target web`) and drives it.

pub mod core;

pub use crate::core::DocSession;

#[cfg(target_arch = "wasm32")]
mod wasm {
    use super::DocSession;
    use wasm_bindgen::prelude::*;

    /// The paged.doc engine handle exposed to the bundle.
    #[wasm_bindgen]
    pub struct DocEngine {
        session: Option<DocSession>,
    }

    #[wasm_bindgen]
    impl DocEngine {
        #[wasm_bindgen(constructor)]
        pub fn new() -> DocEngine {
            DocEngine { session: None }
        }

        /// Load a `.docx`; returns an error string on failure.
        pub fn load_docx(&mut self, bytes: &[u8]) -> Result<(), JsValue> {
            let session = DocSession::load(bytes).map_err(|e| JsValue::from_str(&e))?;
            self.session = Some(session);
            Ok(())
        }

        /// The Tier-0 lowering as a JSON string (the host-model input).
        pub fn lowered_json(&self) -> Result<String, JsValue> {
            self.session
                .as_ref()
                .map(|s| s.lowered_json())
                .ok_or_else(|| JsValue::from_str("no document loaded"))
        }

        /// Number of top-level body blocks in the loaded document.
        pub fn block_count(&self) -> usize {
            self.session.as_ref().map(|s| s.block_count()).unwrap_or(0)
        }

        /// Zero-edit save-back (verbatim carry-through of the retained package).
        pub fn save_verbatim(&self) -> Result<Vec<u8>, JsValue> {
            self.session
                .as_ref()
                .map(|s| s.save_verbatim())
                .ok_or_else(|| JsValue::from_str("no document loaded"))
        }

        /// M2 edited save-back: apply a JSON `EditSet` (keyed by lowered story
        /// `(block, run)` coordinates) as a targeted patch and return the saved
        /// `.docx` bytes. DEFERRED (RFI DOC-03): the editor cannot yet supply
        /// this — `host.nativeDocument.readModel()` hands back opaque core-native
        /// bytes, so today the `EditSet` is produced only by native tests.
        pub fn save_edited(&self, edits_json: &str) -> Result<Vec<u8>, JsValue> {
            let session = self
                .session
                .as_ref()
                .ok_or_else(|| JsValue::from_str("no document loaded"))?;
            let edits: docx_export::EditSet =
                serde_json::from_str(edits_json).map_err(|e| JsValue::from_str(&e.to_string()))?;
            session
                .save_edited(&edits)
                .map_err(|e| JsValue::from_str(&e))
        }

        /// M2 edited save-back driven by the DOC-03 read: forward the host's
        /// `StoryContent` JSON (`host.document.storyContent(storyId)`); it is
        /// overlaid on the import baseline, diffed, and saved. This is the LIVE
        /// path — live once the host injects the v54 read backend (gated by the
        /// bundle on `supports("document.readStory@1")`).
        pub fn save_edited_from_content(&self, content_json: &str) -> Result<Vec<u8>, JsValue> {
            let session = self
                .session
                .as_ref()
                .ok_or_else(|| JsValue::from_str("no document loaded"))?;
            let content: docx_export::StoryContentIn = serde_json::from_str(content_json)
                .map_err(|e| JsValue::from_str(&e.to_string()))?;
            session
                .save_edited_from_content(&content)
                .map_err(|e| JsValue::from_str(&e))
        }
    }

    impl Default for DocEngine {
        fn default() -> Self {
            DocEngine::new()
        }
    }

    #[wasm_bindgen(start)]
    fn start() {
        console_error_panic_hook::set_once();
    }
}
