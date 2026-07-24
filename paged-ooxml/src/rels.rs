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

//! Minimal `_rels` relationship-graph reader.
//!
//! We only need enough of the relationship graph to *locate* the parts Tier-0
//! reads: the main WordprocessingML document, and the styles/numbering parts it
//! references. Parsing is a tolerant `quick-xml` scan over `<Relationship
//! Id=… Type=… Target=…/>` elements — a scanner, not a validator; malformed
//! input yields an empty set rather than a panic.

use quick_xml::events::Event;
use quick_xml::Reader;

/// One `<Relationship>` from a `.rels` part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    /// `Id` attribute.
    pub id: String,
    /// `Type` attribute (a relationship-type URI).
    pub rel_type: String,
    /// `Target` attribute (usually relative to the part's base).
    pub target: String,
    /// `TargetMode` attribute if present (`External` for external targets).
    pub target_mode: Option<String>,
}

/// The parsed relationships of one `.rels` part.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    /// Relationships in document order.
    pub items: Vec<Relationship>,
}

impl Relationships {
    /// Parse a `.rels` part's bytes. Never panics; returns whatever it could scan.
    pub fn parse(bytes: &[u8]) -> Self {
        let mut reader = Reader::from_reader(bytes);
        reader.config_mut().trim_text(false);
        let mut items = Vec::new();
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if local_name(e.name().as_ref()) == b"Relationship" {
                        let mut id = String::new();
                        let mut rel_type = String::new();
                        let mut target = String::new();
                        let mut target_mode = None;
                        for attr in e.attributes().flatten() {
                            let val = attr
                                .decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map(|c| c.into_owned())
                                .unwrap_or_default();
                            match attr.key.as_ref() {
                                b"Id" => id = val,
                                b"Type" => rel_type = val,
                                b"Target" => target = val,
                                b"TargetMode" => target_mode = Some(val),
                                _ => {}
                            }
                        }
                        items.push(Relationship {
                            id,
                            rel_type,
                            target,
                            target_mode,
                        });
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        Relationships { items }
    }

    /// The relationship with the given `Id` (e.g. an `r:embed` image ref).
    pub fn by_id(&self, id: &str) -> Option<&Relationship> {
        self.items.iter().find(|r| r.id == id)
    }

    /// The first relationship whose `Type` ends with `suffix`
    /// (e.g. `"/officeDocument"`, `"/styles"`, `"/numbering"`).
    pub fn by_type_suffix(&self, suffix: &str) -> Option<&Relationship> {
        self.items.iter().find(|r| r.rel_type.ends_with(suffix))
    }
}

/// Strip an XML namespace prefix from a qualified name (`a:b` -> `b`).
fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|&b| b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

/// Join an OPC `Target` (relative to `base_dir`) into a package part name.
///
/// `base_dir` is the directory of the source part (e.g. `word` for
/// `word/document.xml`; empty for the package root). Handles `../` and a leading
/// `/` (absolute-from-root). Returns a normalized `a/b/c.xml` part name.
pub fn resolve_target(base_dir: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut segments: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').filter(|s| !s.is_empty()).collect()
    };
    for seg in target.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

/// The `.rels` part name for a given part (`word/document.xml` ->
/// `word/_rels/document.xml.rels`; `` -> `_rels/.rels`).
pub fn rels_part_name(part: &str) -> String {
    match part.rfind('/') {
        Some(i) => format!("{}/_rels/{}.rels", &part[..i], &part[i + 1..]),
        None if part.is_empty() => "_rels/.rels".to_string(),
        None => format!("_rels/{part}.rels"),
    }
}

/// The directory of a part name (`word/document.xml` -> `word`; `x.xml` -> ``).
pub fn part_dir(part: &str) -> &str {
    match part.rfind('/') {
        Some(i) => &part[..i],
        None => "",
    }
}
