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

//! Error type for the OOXML foundation.

use thiserror::Error;

/// Errors from reading / writing an OPC package.
///
/// The scanner discipline (mirrors `plugin-web`'s `web-model`): a malformed
/// package produces an `Err`, never a panic. Callers decide whether to surface
/// it as an honest diagnostic or bail.
#[derive(Debug, Error)]
pub enum OoxmlError {
    /// The container is not a readable ZIP / OPC package.
    #[error("not a valid OPC (zip) package: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// A legacy binary Word document (CFB/OLE, the pre-2007 `.doc`
    /// format), which this engine does not read.
    ///
    /// Detected by container magic BEFORE the zip reader runs, because
    /// some of these files ALSO embed an OPC package further in — real
    /// example from the corpus: a 420 KB CFB preamble followed by a
    /// complete `[Content_Types].xml` archive. The zip crate scans for
    /// the end-of-central-directory record, finds that embedded package,
    /// and happily enumerates it; main-part resolution then lands on
    /// `theme/theme/themeManager.xml` and the user sees "unexpected tag
    /// while parsing Document (expected Document, found a:themeManager)"
    /// for what is simply an old Word file. Honest degradation (ADR-007)
    /// means saying which format it is, at the door.
    #[error("legacy binary Word document (.doc, CFB/OLE) — only OOXML .docx/.dotx is supported; re-save as .docx")]
    LegacyBinaryDoc,

    /// Rich Text Format, which Word writes and users routinely save with a
    /// `.doc` extension (the corpus has one). Sniffed for the same reason
    /// as `LegacyBinaryDoc`: without it the zip reader reports "Could not
    /// find EOCD", which tells the user nothing about what they opened.
    #[error("Rich Text Format document — only OOXML .docx/.dotx is supported; re-save as .docx")]
    RichTextFormat,

    /// I/O failure while reading or writing part bytes.
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),

    /// A required part (by convention or by `_rels` resolution) was absent.
    #[error("required part not found: {0}")]
    MissingPart(String),

    /// XML in a part we parse was malformed.
    #[error("malformed xml in {part}: {detail}")]
    Xml { part: String, detail: String },

    /// The vendored `ooxmlsdk` typed-DOM (de)serializer rejected a part
    /// (see [`crate::dom`]). Carried as a string so this crate's public error
    /// does not leak `ooxmlsdk` types.
    #[error("ooxml typed-dom error: {0}")]
    Sdk(String),

    /// The package violates a structural OPC guarantee not covered above.
    #[error("opc structure error: {0}")]
    Structure(String),
}

impl OoxmlError {
    /// Build an [`OoxmlError::Xml`] from any displayable parse error, tagged
    /// with the part name being read.
    pub fn xml(part: impl Into<String>, detail: impl std::fmt::Display) -> Self {
        OoxmlError::Xml {
            part: part.into(),
            detail: detail.to_string(),
        }
    }
}

impl From<ooxmlsdk::common::SdkError> for OoxmlError {
    fn from(e: ooxmlsdk::common::SdkError) -> Self {
        OoxmlError::Sdk(e.to_string())
    }
}

/// Convenience result alias.
pub type Result<T> = std::result::Result<T, OoxmlError>;
