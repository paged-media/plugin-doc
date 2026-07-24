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

//! `[Content_Types].xml` — the OPC content-type table.
//!
//! Extension defaults + per-part overrides, in document order. We keep this as a
//! *read* model this pass (it is needed to resolve a part's content type and to
//! know which overrides exist); the writer re-emits the part verbatim, so no
//! re-serialization is required until save-back grows new parts (M2+).

use crate::error::{OoxmlError, Result};

/// The reserved OPC part name of the content-type table.
pub const CONTENT_TYPES_PART: &str = "[Content_Types].xml";

/// Parsed `[Content_Types].xml`: extension defaults + per-part overrides.
#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    /// `<Default Extension=".." ContentType="..">` in document order.
    pub defaults: Vec<(String, String)>,
    /// `<Override PartName="/.." ContentType="..">` in document order.
    pub overrides: Vec<(String, String)>,
}

impl ContentTypes {
    /// Parse the content-type table. Only `Default`/`Override` elements matter;
    /// anything else is ignored (tolerant scan, never panics).
    pub fn parse(xml: &[u8]) -> Result<ContentTypes> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut ct = ContentTypes::default();
        let mut buf = Vec::new();
        loop {
            let ev = reader
                .read_event_into(&mut buf)
                .map_err(|e| OoxmlError::xml(CONTENT_TYPES_PART, e))?;
            match ev {
                Event::Empty(e) | Event::Start(e) => match e.local_name().as_ref() {
                    b"Default" => {
                        if let (Some(ext), Some(ty)) =
                            (attr(&e, b"Extension")?, attr(&e, b"ContentType")?)
                        {
                            ct.defaults.push((ext, ty));
                        }
                    }
                    b"Override" => {
                        if let (Some(pn), Some(ty)) =
                            (attr(&e, b"PartName")?, attr(&e, b"ContentType")?)
                        {
                            ct.overrides.push((pn, ty));
                        }
                    }
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(ct)
    }

    /// Resolve the content type of a part (leading-slash absolute name, e.g.
    /// `/word/document.xml`): an explicit override wins, else the extension
    /// default, else `None`.
    pub fn content_type_of(&self, absolute_part_name: &str) -> Option<&str> {
        if let Some((_, ty)) = self
            .overrides
            .iter()
            .find(|(pn, _)| pn == absolute_part_name)
        {
            return Some(ty);
        }
        let ext = absolute_part_name.rsplit('.').next()?;
        self.defaults
            .iter()
            .find(|(e, _)| e.eq_ignore_ascii_case(ext))
            .map(|(_, ty)| ty.as_str())
    }

    /// True if an explicit `<Override>` exists for `absolute_part_name`.
    pub fn has_override(&self, absolute_part_name: &str) -> bool {
        self.overrides
            .iter()
            .any(|(pn, _)| pn == absolute_part_name)
    }
}

/// Read an attribute's UTF-8 value by local name (namespace prefix ignored).
pub(crate) fn attr(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Result<Option<String>> {
    for a in e.attributes() {
        let a = a.map_err(|e| OoxmlError::xml(CONTENT_TYPES_PART, e))?;
        if a.key.local_name().as_ref() == key {
            let v = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .map_err(|e| OoxmlError::xml(CONTENT_TYPES_PART, e))?;
            return Ok(Some(v.into_owned()));
        }
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_defaults_and_overrides_and_resolves() {
        let xml = br#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
        let ct = ContentTypes::parse(xml).unwrap();
        assert_eq!(ct.defaults.len(), 2);
        assert_eq!(ct.overrides.len(), 1);
        assert!(ct.has_override("/word/document.xml"));
        assert!(ct
            .content_type_of("/word/document.xml")
            .unwrap()
            .contains("wordprocessingml.document.main"));
        // resolves by extension default when no override
        assert!(ct
            .content_type_of("/word/_rels/document.xml.rels")
            .is_some());
    }

    #[test]
    fn tolerates_garbage_without_panicking() {
        assert!(
            ContentTypes::parse(b"not xml at all <<<").is_ok()
                || ContentTypes::parse(b"<<<").is_err()
        );
    }
}
