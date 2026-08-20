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
/// Deepest element nesting `parse_root` will hand to the SDK.
///
/// `ooxmlsdk` deserialises by recursive descent, so nesting depth in the
/// XML becomes stack depth in the process. Apache POI's corpus carries
/// `deep-table-cell.docx` for exactly this reason — it is a deliberately
/// deep table nest that has crashed parsers before, and it crashed this
/// one: `fatal runtime error: stack overflow, aborting`, SIGABRT, no
/// unwinding and nothing catchable.
///
/// A stack overflow is not a parse error. It takes the whole process
/// down, which in the editor means the wasm module dies mid-document —
/// so this has to be refused BEFORE the recursive parser sees it.
///
/// 256 is far beyond anything a real document reaches (Word's own table
/// nesting limit is 20) while still being an order of magnitude below
/// where the SDK's frames become dangerous.
const MAX_ELEMENT_DEPTH: usize = 256;

/// Reject XML nested deeper than [`MAX_ELEMENT_DEPTH`] before it reaches
/// the SDK's recursive descent. Cheap: one non-allocating scan.
fn depth_within_limit(bytes: &[u8]) -> std::result::Result<(), usize> {
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut buf = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(_)) => {
                depth += 1;
                if depth > MAX_ELEMENT_DEPTH {
                    return Err(depth);
                }
            }
            Ok(quick_xml::events::Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(quick_xml::events::Event::Eof) => return Ok(()),
            // A malformed document is the SDK's problem to report, with
            // its own better message — this guard only judges DEPTH.
            Err(_) => return Ok(()),
            _ => {}
        }
        buf.clear();
    }
}

pub fn parse_root<T: SdkType>(part_name: &str, bytes: &[u8]) -> Result<T> {
    if let Err(depth) = depth_within_limit(bytes) {
        return Err(OoxmlError::NestingTooDeep {
            part: part_name.to_string(),
            depth,
            limit: MAX_ELEMENT_DEPTH,
        });
    }
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
