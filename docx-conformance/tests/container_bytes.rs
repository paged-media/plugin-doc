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

//! The container bytes `OpcPackage::write` produces, read back field by field.
//!
//! An edited save re-zips the package, so the ZIP metadata is part of what a
//! user gets back. It must be a function of the parts alone: no wall-clock
//! timestamp, no build-host `version made by`, no extra fields, data
//! descriptors or ZIP64 records. These are the values every save has carried
//! so far; a `zip` upgrade that changes one of them fails here first.
//!
//! And the reader must not rename a part: an entry name stored as UTF-8
//! without the UTF-8 flag keeps its bytes through a save.

// The `__feat__` suffix is the state-registry attribution convention.
#![allow(non_snake_case)]

use docx_conformance::{memo_docx, zip_parts};
use paged_ooxml::OpcPackage;

fn u16_at(b: &[u8], at: usize) -> u16 {
    u16::from_le_bytes([b[at], b[at + 1]])
}

fn u32_at(b: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([b[at], b[at + 1], b[at + 2], b[at + 3]])
}

/// One entry's header fields: (central, local).
#[derive(Debug)]
struct Entry {
    name: Vec<u8>,
    made_by: u16,
    needed: u16,
    flags: u16,
    method: u16,
    time: u16,
    date: u16,
    central_extra: u16,
    comment: u16,
    external: u32,
    local_needed: u16,
    local_flags: u16,
    local_time: u16,
    local_date: u16,
    local_extra: u16,
}

/// Walk the central directory from the end record, and each entry's local
/// header from the offset the central directory gives.
fn entries(zip: &[u8]) -> (Vec<Entry>, u16) {
    let eocd = zip
        .windows(4)
        .rposition(|w| w == b"PK\x05\x06")
        .expect("end of central directory");
    let count = u16_at(zip, eocd + 10) as usize;
    let mut at = u32_at(zip, eocd + 16) as usize;
    let archive_comment = u16_at(zip, eocd + 20);
    let mut out = Vec::new();
    for _ in 0..count {
        assert_eq!(&zip[at..at + 4], b"PK\x01\x02", "central header");
        let name_len = u16_at(zip, at + 28) as usize;
        let extra = u16_at(zip, at + 30);
        let comment = u16_at(zip, at + 32);
        let local = u32_at(zip, at + 42) as usize;
        assert_eq!(&zip[local..local + 4], b"PK\x03\x04", "local header");
        out.push(Entry {
            name: zip[at + 46..at + 46 + name_len].to_vec(),
            made_by: u16_at(zip, at + 4),
            needed: u16_at(zip, at + 6),
            flags: u16_at(zip, at + 8),
            method: u16_at(zip, at + 10),
            time: u16_at(zip, at + 12),
            date: u16_at(zip, at + 14),
            central_extra: extra,
            comment,
            external: u32_at(zip, at + 38),
            local_needed: u16_at(zip, local + 4),
            local_flags: u16_at(zip, local + 6),
            local_time: u16_at(zip, local + 10),
            local_date: u16_at(zip, local + 12),
            local_extra: u16_at(zip, local + 28),
        });
        at += 46 + name_len + extra as usize + comment as usize;
    }
    (out, archive_comment)
}

/// 1980-01-01 00:00 in MS-DOS date/time.
const DOS_DATE_1980: u16 = (1 << 5) | 1;

#[test]
fn a_saved_package_carries_fixed_zip_metadata__feat__plugin_doc_opc_foundation() {
    let mut pkg = OpcPackage::read(&memo_docx()).unwrap();
    // A directory entry and a non-ASCII name, besides the memo's own parts.
    pkg.set_part("word/m\u{e4}rz.xml", b"<x/>".to_vec());
    let dir = {
        let mut cursor = std::io::Cursor::new(Vec::new());
        let mut z = zip::ZipWriter::new(&mut cursor);
        z.add_directory("media", zip::write::SimpleFileOptions::DEFAULT)
            .unwrap();
        z.finish().unwrap();
        cursor.into_inner()
    };
    let dir_pkg = OpcPackage::read(&dir).unwrap();
    let saved = [pkg.write().unwrap(), dir_pkg.write().unwrap()];

    for zip in &saved {
        assert!(
            !zip.windows(4)
                .any(|w| w == b"PK\x06\x06" || w == b"PK\x07\x08"),
            "no ZIP64 end record and no data descriptor"
        );
        let (entries, archive_comment) = entries(zip);
        assert_eq!(archive_comment, 0);
        for e in &entries {
            let name = String::from_utf8_lossy(&e.name);
            let is_dir = e.name.ends_with(b"/");
            let utf8 = if e.name.is_ascii() { 0 } else { 1 << 11 };
            // Deflated files and directories need 2.0; made by Unix (3).
            assert_eq!(e.made_by, 0x0314, "{name}: version made by");
            assert_eq!((e.needed, e.local_needed), (20, 20), "{name}");
            assert_eq!((e.flags, e.local_flags), (utf8, utf8), "{name}: flags");
            assert_eq!(e.method, if is_dir { 0 } else { 8 }, "{name}: method");
            assert_eq!((e.time, e.local_time), (0, 0), "{name}: time");
            assert_eq!(
                (e.date, e.local_date),
                (DOS_DATE_1980, DOS_DATE_1980),
                "{name}: date"
            );
            assert_eq!((e.central_extra, e.local_extra), (0, 0), "{name}: extra");
            assert_eq!(e.comment, 0, "{name}: comment");
            let mode = if is_dir { 0o40755 } else { 0o100644 };
            assert_eq!(e.external, mode << 16, "{name}: external attributes");
        }
    }
    let (dir_entries, _) = entries(&saved[1]);
    assert_eq!(
        dir_entries[0].name, b"media/",
        "the directory entry survives"
    );
}

#[test]
fn an_unflagged_utf8_entry_name_keeps_its_bytes__feat__plugin_doc_opc_foundation() {
    let name = "word/m\u{e4}rz.xml";
    let mut zip = zip_parts(&[("[Content_Types].xml", b"<Types/>"), (name, b"<x/>")]);
    // Clear the UTF-8 flag, as some writers do for UTF-8 names.
    let headers: Vec<(usize, usize)> = zip
        .windows(4)
        .enumerate()
        .filter_map(|(i, w)| match w {
            b"PK\x03\x04" => Some((i, 6)),
            b"PK\x01\x02" => Some((i, 8)),
            _ => None,
        })
        .collect();
    for (at, flags) in headers {
        zip[at + flags + 1] &= !0x08;
    }
    let (read_back, _) = entries(&zip);
    assert!(read_back.iter().all(|e| e.flags & (1 << 11) == 0));

    let pkg = OpcPackage::read(&zip).unwrap();
    assert_eq!(pkg.part(name), Some(&b"<x/>"[..]), "read as UTF-8");
    let (saved, _) = entries(&pkg.write().unwrap());
    assert_eq!(
        saved[1].name,
        name.as_bytes(),
        "the stored name is unchanged"
    );
}
