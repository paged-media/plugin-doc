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

//! The OPC (Open Packaging Conventions) container.
//!
//! We own the container layer over `zip` + `quick-xml` rather than delegating it
//! to `ooxmlsdk`'s `parts` feature, because `ooxmlsdk` does not document
//! byte-exact fidelity and the **preservation invariant** (§10 of the spec —
//! "Paged never destroys a document") requires verbatim carry-through of parts we
//! do not touch. `ooxmlsdk` is used only for the *typed DOM of the parts we
//! actually read/write* (see [`crate::dom`]).
//!
//! ## Preservation model (mirrors `sheet-xlsx` `opc.rs`)
//!
//! An [`OpcPackage`] holds an **ordered** `Vec<PartEntry>` preserving on-disk
//! entry order. Every part is read as [`PartEntry::Opaque`] bytes (stored
//! *decompressed*). A part is only re-serialized from a model when the caller
//! explicitly [`OpcPackage::set_part`]s it (marking it dirty). On write, an
//! untouched part re-emits its stored original bytes verbatim, so unknown parts
//! and unknown subtrees survive a round-trip. Per-part **decompressed** byte
//! identity is guaranteed; whole-file / zip-metadata identity is not (the
//! deflate stream is re-encoded — the same caveat `sheet-xlsx` documents).

use std::io::{Cursor, Read, Write};

use crate::error::{OoxmlError, Result};

/// One entry of the package, in on-disk order.
#[derive(Debug, Clone)]
pub enum PartEntry {
    /// A file part, stored decompressed. `dirty` is set when [`OpcPackage::set_part`]
    /// replaces the bytes, so the writer knows this part changed.
    File {
        /// OPC part name, e.g. `word/document.xml` (no leading slash).
        name: String,
        /// Decompressed bytes.
        bytes: Vec<u8>,
        /// True once the caller replaced these bytes.
        dirty: bool,
    },
    /// An explicit directory entry (rare in `.docx`, but preserved when present).
    Dir {
        /// Directory name, ending in `/`.
        name: String,
    },
}

impl PartEntry {
    /// The entry's name.
    pub fn name(&self) -> &str {
        match self {
            PartEntry::File { name, .. } => name,
            PartEntry::Dir { name } => name,
        }
    }
}

/// An OPC package read from `.docx`/`.dotx` bytes, preserving part order.
#[derive(Debug, Clone, Default)]
pub struct OpcPackage {
    parts: Vec<PartEntry>,
}

impl OpcPackage {
    /// Read a package from `.docx` bytes. Never panics on malformed input — a
    /// non-ZIP payload returns [`OoxmlError::Zip`].
    pub fn read(bytes: &[u8]) -> Result<Self> {
        // Container sniff BEFORE the zip reader. A legacy .doc is CFB/OLE,
        // but some of them EMBED a complete OPC package after the CFB
        // payload; the zip crate finds that by scanning backwards for the
        // EOCD and opens it, so without this guard an old Word file
        // half-imports and dies later with a misleading XML error. See
        // OoxmlError::LegacyBinaryDoc.
        const CFB_MAGIC: &[u8] = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        if bytes.starts_with(CFB_MAGIC) {
            return Err(crate::error::OoxmlError::LegacyBinaryDoc);
        }
        if bytes.starts_with(b"{\\rtf") {
            return Err(crate::error::OoxmlError::RichTextFormat);
        }
        // WinWord 2.0 — older than CFB/OLE and a different container
        // entirely, so the magic above does not catch it. Found in
        // Apache POI's corpus (`word2.doc`), which was the first time
        // this parser had seen one: it fell through to the zip reader
        // and reported "Could not find EOCD", telling the user nothing.
        const WINWORD2_MAGIC: &[u8] = &[0xDB, 0xA5, 0x2D, 0x00];
        if bytes.starts_with(WINWORD2_MAGIC) {
            return Err(crate::error::OoxmlError::LegacyBinaryDoc);
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes))?;
        let mut parts = Vec::with_capacity(archive.len());
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();
            if file.is_dir() {
                parts.push(PartEntry::Dir { name });
            } else {
                let mut buf = Vec::with_capacity(file.size() as usize);
                file.read_to_end(&mut buf)?;
                parts.push(PartEntry::File {
                    name,
                    bytes: buf,
                    dirty: false,
                });
            }
        }
        Ok(OpcPackage { parts })
    }

    /// All part names in on-disk order (files and dirs).
    pub fn names(&self) -> impl Iterator<Item = &str> {
        self.parts.iter().map(PartEntry::name)
    }

    /// All *file* part names in on-disk order.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.parts.iter().filter_map(|p| match p {
            PartEntry::File { name, .. } => Some(name.as_str()),
            PartEntry::Dir { .. } => None,
        })
    }

    /// Borrow the decompressed bytes of a file part by name.
    pub fn part(&self, name: &str) -> Option<&[u8]> {
        self.parts.iter().find_map(|p| match p {
            PartEntry::File { name: n, bytes, .. } if n == name => Some(bytes.as_slice()),
            _ => None,
        })
    }

    /// Borrow required part bytes, or [`OoxmlError::MissingPart`].
    pub fn require(&self, name: &str) -> Result<&[u8]> {
        self.part(name)
            .ok_or_else(|| OoxmlError::MissingPart(name.to_string()))
    }

    /// Replace a part's bytes, marking it dirty. If the part does not exist it is
    /// appended (a new part). Directory entries are ignored as targets.
    pub fn set_part(&mut self, name: &str, new_bytes: Vec<u8>) {
        for p in &mut self.parts {
            if let PartEntry::File {
                name: n,
                bytes,
                dirty,
            } = p
            {
                if n == name {
                    *bytes = new_bytes;
                    *dirty = true;
                    return;
                }
            }
        }
        self.parts.push(PartEntry::File {
            name: name.to_string(),
            bytes: new_bytes,
            dirty: true,
        });
    }

    /// True if any part was modified since read.
    pub fn is_dirty(&self) -> bool {
        self.parts
            .iter()
            .any(|p| matches!(p, PartEntry::File { dirty: true, .. }))
    }

    /// Re-serialize the package to `.docx` bytes. Untouched parts re-emit their
    /// stored bytes verbatim (the carry-through guarantee); modified parts emit
    /// their new bytes. Entry order is preserved.
    pub fn write(&self) -> Result<Vec<u8>> {
        let mut out = Cursor::new(Vec::new());
        {
            let mut zip = zip::ZipWriter::new(&mut out);
            let opts = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for entry in &self.parts {
                match entry {
                    PartEntry::Dir { name } => {
                        zip.add_directory(name.trim_end_matches('/'), opts)?;
                    }
                    PartEntry::File { name, bytes, .. } => {
                        zip.start_file(name, opts)?;
                        zip.write_all(bytes)?;
                    }
                }
            }
            zip.finish()?;
        }
        Ok(out.into_inner())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::OoxmlError;

    /// A minimal CFB/OLE header followed by a complete OPC package — the
    /// shape four real corpus `.doc` files have. Without the magic sniff
    /// the zip reader finds the embedded package by scanning backwards
    /// for the EOCD, enumerates it, and the failure surfaces much later
    /// as a nonsense XML error about `themeManager.xml`.
    #[test]
    fn a_legacy_doc_that_embeds_an_opc_package_is_rejected_at_the_door() {
        let mut package = OpcPackage::default();
        package.set_part("[Content_Types].xml", b"<Types/>".to_vec());
        let inner = package.write().expect("write inner package");

        let mut bytes = vec![0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        bytes.resize(512, 0);
        bytes.extend_from_slice(&inner);

        // The trap: the payload IS a readable zip if you look far enough in.
        assert!(
            zip::ZipArchive::new(Cursor::new(&bytes[..])).is_ok(),
            "test premise broken — the embedded package must be zip-readable, \
             otherwise this test would pass without the sniff"
        );
        assert!(
            matches!(OpcPackage::read(&bytes), Err(OoxmlError::LegacyBinaryDoc)),
            "a CFB container must be named as a legacy .doc, not half-opened"
        );
    }

    #[test]
    fn an_rtf_saved_as_doc_is_named_rather_than_reported_as_a_broken_zip() {
        let bytes = br"{\rtf1\ansi\deff0{\fonttbl{\f0 Times;}}\f0 hello}".to_vec();
        assert!(
            matches!(OpcPackage::read(&bytes), Err(OoxmlError::RichTextFormat)),
            "RTF must be identified; the bare zip reader calls it \
             \"Could not find EOCD\", which names nothing"
        );
    }

    #[test]
    fn genuinely_malformed_input_still_reports_as_a_container_failure() {
        // The sniffs must not swallow everything — random bytes are still
        // a zip error, not a format we claim to recognise.
        assert!(matches!(
            OpcPackage::read(b"not a container at all"),
            Err(OoxmlError::Zip(_))
        ));
    }
}
