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
