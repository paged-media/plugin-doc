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

use docx_core::{Justification, ParaProps, RunProps, VertAlign};
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

/// A `<w:pPr>` element for a paragraph's applied style + direct formatting —
/// the paragraph twin of [`render_rpr`], in WML `CT_PPr` child order. `pstyle` is
/// a REAL Word style id; synthesized paragraph styles must already have been
/// projected into `props`.
pub fn render_ppr(props: &ParaProps, pstyle: Option<&str>) -> Vec<u8> {
    let mut s = String::from("<w:pPr>");
    if let Some(id) = pstyle {
        s.push_str(&format!("<w:pStyle w:val=\"{}\"/>", escape(id)));
    }
    if props.keep_next == Some(true) {
        s.push_str("<w:keepNext/>");
    }
    if props.keep_lines == Some(true) {
        s.push_str("<w:keepLines/>");
    }
    // <w:spacing> carries before/after together.
    if props.space_before.is_some() || props.space_after.is_some() {
        s.push_str("<w:spacing");
        if let Some(v) = props.space_before {
            s.push_str(&format!(" w:before=\"{v}\""));
        }
        if let Some(v) = props.space_after {
            s.push_str(&format!(" w:after=\"{v}\""));
        }
        s.push_str("/>");
    }
    // <w:ind> carries the indents together; a hanging indent is emitted as
    // `w:hanging` (a positive twip value), never as a negative firstLine.
    if props.left_indent.is_some()
        || props.right_indent.is_some()
        || props.first_line_indent.is_some()
        || props.hanging_indent.is_some()
    {
        s.push_str("<w:ind");
        if let Some(v) = props.left_indent {
            s.push_str(&format!(" w:left=\"{v}\""));
        }
        if let Some(v) = props.right_indent {
            s.push_str(&format!(" w:right=\"{v}\""));
        }
        if let Some(v) = props.hanging_indent {
            s.push_str(&format!(" w:hanging=\"{v}\""));
        } else if let Some(v) = props.first_line_indent {
            s.push_str(&format!(" w:firstLine=\"{v}\""));
        }
        s.push_str("/>");
    }
    if !props.tabs.is_empty() {
        s.push_str("<w:tabs>");
        for t in &props.tabs {
            let val = t.alignment.as_deref().unwrap_or("left");
            s.push_str(&format!(
                "<w:tab w:val=\"{}\" w:pos=\"{}\"/>",
                escape(val),
                t.position
            ));
        }
        s.push_str("</w:tabs>");
    }
    if let Some(j) = props.justification {
        let v = match j {
            Justification::Left | Justification::Start => "left",
            Justification::Center => "center",
            Justification::Right | Justification::End => "right",
            Justification::Both => "both",
            Justification::Distribute => "distribute",
        };
        s.push_str(&format!("<w:jc w:val=\"{v}\"/>"));
    }
    s.push_str("</w:pPr>");
    s.into_bytes()
}

/// A whole `<w:r>` element: optional `<w:rPr>` + the text. Used by the
/// structural inserts (Increment 2).
pub fn render_run(text: &str, props: &RunProps, rstyle: Option<&str>) -> Vec<u8> {
    let mut out = b"<w:r>".to_vec();
    if props != &RunProps::default() || rstyle.is_some() {
        out.extend_from_slice(&render_rpr(props, rstyle));
    }
    out.extend_from_slice(&render_wt(text));
    out.extend_from_slice(b"</w:r>");
    out
}

/// A whole `<w:p>` element carrying one run, with an optional `<w:pStyle>`.
pub fn render_paragraph(
    text: &str,
    props: &RunProps,
    rstyle: Option<&str>,
    para_style: Option<&str>,
) -> Vec<u8> {
    let mut out = b"<w:p>".to_vec();
    if let Some(ps) = para_style {
        out.extend_from_slice(
            format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", escape(ps)).as_bytes(),
        );
    }
    out.extend_from_slice(&render_run(text, props, rstyle));
    out.extend_from_slice(b"</w:p>");
    out
}

/// A whole `<w:tr>` element with one `<w:tc>` per `cells` entry. Each cell holds
/// a single paragraph + run (a `<w:tc>` MUST contain at least one block-level
/// child, so an empty cell still gets a `<w:p>`).
pub fn render_table_row(cells: &[String]) -> Vec<u8> {
    let mut out = b"<w:tr>".to_vec();
    for text in cells {
        out.extend_from_slice(b"<w:tc><w:p>");
        if !text.is_empty() {
            out.extend_from_slice(&render_run(text, &RunProps::default(), None));
        }
        out.extend_from_slice(b"</w:p></w:tc>");
    }
    out.extend_from_slice(b"</w:tr>");
    out
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
