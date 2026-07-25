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

//! Fragment emitters: `RunProps -> <w:rPr>…</w:rPr>` and text `-> <w:t>…</w:t>`.
//! The exact inverse of `docx-import`'s `apply_run_property_choice`, in WML
//! `CT_RPr` child order. `w:`-prefixed element names resolve against the
//! `xmlns:w` decl on the root element (copied verbatim in the surrounding bytes).

use docx_core::{RunProps, VertAlign};
use quick_xml::escape::escape;

/// A run's `<w:t>` element with the new text, always `xml:space="preserve"` so
/// leading/trailing spaces survive.
pub fn render_wt(text: &str) -> Vec<u8> {
    format!("<w:t xml:space=\"preserve\">{}</w:t>", escape(text)).into_bytes()
}

/// A `<w:rPr>` element for the run's effective direct properties. `rstyle` is a
/// REAL Word style id (`<w:rStyle>`), emitted first per the schema; synthesized
/// styles must already have been projected into `props`, never passed here.
pub fn render_rpr(props: &RunProps, rstyle: Option<&str>) -> Vec<u8> {
    let mut s = String::from("<w:rPr>");
    if let Some(id) = rstyle {
        s.push_str(&format!("<w:rStyle w:val=\"{}\"/>", escape(id)));
    }
    if let Some(font) = &props.font {
        s.push_str(&format!("<w:rFonts w:ascii=\"{}\"/>", escape(font)));
    }
    toggle(&mut s, "b", props.bold);
    toggle(&mut s, "i", props.italic);
    toggle(&mut s, "caps", props.caps);
    toggle(&mut s, "smallCaps", props.small_caps);
    toggle(&mut s, "strike", props.strike);
    if let Some(color) = &props.color {
        s.push_str(&format!("<w:color w:val=\"{}\"/>", escape(color)));
    }
    if let Some(pos) = props.baseline_half_pts {
        s.push_str(&format!("<w:position w:val=\"{pos}\"/>"));
    }
    if let Some(sz) = props.size_half_pts {
        s.push_str(&format!("<w:sz w:val=\"{sz}\"/>"));
    }
    match props.underline {
        Some(true) => s.push_str("<w:u w:val=\"single\"/>"),
        Some(false) => s.push_str("<w:u w:val=\"none\"/>"),
        None => {}
    }
    if let Some(va) = &props.vert_align {
        let v = match va {
            VertAlign::Baseline => "baseline",
            VertAlign::Superscript => "superscript",
            VertAlign::Subscript => "subscript",
        };
        s.push_str(&format!("<w:vertAlign w:val=\"{v}\"/>"));
    }
    s.push_str("</w:rPr>");
    s.into_bytes()
}

/// Emit a boolean toggle property: `Some(true)` ⇒ `<w:NAME/>`, `Some(false)` ⇒
/// `<w:NAME w:val="false"/>`, `None` ⇒ omitted (inherit). Matches the `on()`
/// reading in `docx-import` (absent `w:val` ⇒ true).
fn toggle(s: &mut String, name: &str, v: Option<bool>) {
    match v {
        Some(true) => s.push_str(&format!("<w:{name}/>")),
        Some(false) => s.push_str(&format!("<w:{name} w:val=\"false\"/>")),
        None => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn wt_preserves_space_and_escapes() {
        assert_eq!(
            render_wt(" a & b "),
            b"<w:t xml:space=\"preserve\"> a &amp; b </w:t>"
        );
    }

    #[test]
    fn rpr_bold_off_keeps_only_color() {
        // The slice's edit: bold toggled off, color kept.
        let props = RunProps {
            color: Some("FF0000".into()),
            ..Default::default()
        };
        assert_eq!(
            render_rpr(&props, None),
            b"<w:rPr><w:color w:val=\"FF0000\"/></w:rPr>".to_vec()
        );
    }

    #[test]
    fn rpr_child_order_and_toggles() {
        let props = RunProps {
            bold: Some(true),
            italic: Some(false),
            size_half_pts: Some(24),
            underline: Some(true),
            ..Default::default()
        };
        assert_eq!(
            String::from_utf8(render_rpr(&props, Some("Emphasis"))).unwrap(),
            "<w:rPr><w:rStyle w:val=\"Emphasis\"/><w:b/><w:i w:val=\"false\"/>\
             <w:sz w:val=\"24\"/><w:u w:val=\"single\"/></w:rPr>"
        );
    }
}
