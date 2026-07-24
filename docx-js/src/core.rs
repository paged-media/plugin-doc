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
//! invariant, its parsed semantic model, and the lowering). All real work lives
//! here so it is native-testable; `lib.rs`'s `#[wasm_bindgen]` layer is a pure
//! forwarding shim.

use docx_core::DocxDocument;
use docx_import::import_docx;
use docx_lower::ir::LoweredDoc;
use docx_lower::lower;

/// A loaded Word document session.
pub struct DocSession {
    /// The original `.docx` bytes, retained for the preservation invariant
    /// (zero-edit save-back re-emits them verbatim).
    source: Vec<u8>,
    /// The parsed semantic model.
    model: DocxDocument,
}

impl DocSession {
    /// Load a `.docx`/`.dotx` package. Returns a human-readable error string on a
    /// hard failure (unreadable container / unparseable main document).
    pub fn load(bytes: &[u8]) -> Result<DocSession, String> {
        let model = import_docx(bytes).map_err(|e| e.to_string())?;
        Ok(DocSession {
            source: bytes.to_vec(),
            model,
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

    /// Zero-edit save-back: re-emit the retained source verbatim. (Edited
    /// save-back — projecting native changes back to WordprocessingML — is M2.)
    pub fn save_verbatim(&self) -> Vec<u8> {
        self.source.clone()
    }
}
