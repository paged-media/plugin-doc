/*
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * This file is part of paged (https://paged.media) and is additionally
 * available under the Paged Media Enterprise License (PMEL). Full
 * copyright and license information is available in LICENSE.md which is
 * distributed with this source code.
 *
 *  @copyright  Copyright (c) And The Next GmbH
 *  @license    MPL-2.0 OR Paged Media Enterprise License (PMEL)
 */

//! The typed-DOM bridge to the vendored `ooxmlsdk` crate.
//!
//! The OPC container ([`crate::opc`]) gives us the *raw bytes* of a part; this
//! module turns those bytes into `ooxmlsdk`'s code-generated typed part tree
//! (and back). We deliberately keep the container and the typed DOM separate:
//! the container owns preservation (verbatim carry-through), the typed DOM owns
//! the *semantic* read of the parts we actually understand.
//!
//! DOC-02 (the `ooxmlsdk` → `wasm32` feasibility spike) is GREEN, so this is the
//! committed semantic-parse foundation. `ooxmlsdk`'s `mce` feature is enabled at
//! the workspace level, so `mc:AlternateContent` fallback selection is handled by
//! the typed DOM at parse time.

use ooxmlsdk::sdk::SdkType;

use crate::error::{OoxmlError, Result};

/// Re-export the vendored typed part DOM so every `docx-*` crate imports the
/// OOXML schemas from ONE place (`paged_ooxml::ooxmlsdk::…`), never depending on
/// the exact crate version directly.
pub use ooxmlsdk;

/// Deserialize a part's bytes into an `ooxmlsdk` root type `T`
/// (e.g. the WordprocessingML `Document`, `Styles`, `Numbering`).
///
/// Never panics on malformed input — a rejected part surfaces as
/// [`OoxmlError::Sdk`]. `part_name` is threaded through only for error context.
pub fn parse_root<T: SdkType>(part_name: &str, bytes: &[u8]) -> Result<T> {
    T::from_bytes(bytes).map_err(|e| OoxmlError::Sdk(format!("{part_name}: {e}")))
}

/// Serialize an `ooxmlsdk` root type back to XML bytes (used by save-back in M2;
/// exercised now only for the round-trip identity harness).
pub fn serialize_root<T: SdkType>(value: &T) -> Result<Vec<u8>> {
    let mut out = Vec::with_capacity(4096);
    value
        .write_to(&mut out)
        .map_err(|e| OoxmlError::Sdk(e.to_string()))?;
    Ok(out)
}
