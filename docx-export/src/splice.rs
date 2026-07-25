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

//! The byte-level targeted patcher. quick-xml is used ONLY as a locator: a single
//! streaming pass computes the source byte ranges of the `<w:t>`/`<w:rPr>` to
//! rewrite, and every byte outside those ranges is copied verbatim from `src`.
//! Untouched subtrees (and every other part, via `OpcPackage`) are therefore
//! byte-identical **by construction** — no serializer, no re-emission. The
//! ooxmlsdk `write_to` path is never touched (the 8 MiB wasm-budget guard).

use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader;

use crate::rpr::render_wt;

/// A run to patch, addressed by source ordinals + the pre-rendered replacements.
#[derive(Debug, Clone)]
pub struct ResolvedTarget {
    /// The run's paragraph ordinal (direct `<w:p>` child of `<w:body>`).
    pub para_ord: u32,
    /// The run's ordinal (direct `<w:r>` child of the `<w:p>`).
    pub run_ord: u32,
    /// New run text (`Some` ⇒ rewrite the run's `<w:t>`).
    pub new_text: Option<String>,
    /// Pre-rendered `<w:rPr>…</w:rPr>` bytes (`Some` ⇒ replace/insert the rPr).
    pub new_rpr: Option<Vec<u8>>,
}

fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|&b| b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

/// Rewrite `word/document.xml`'s bytes, splicing only the targeted runs'
/// `<w:t>`/`<w:rPr>` and copying everything else verbatim. Malformed input stops
/// the scan and flushes the remainder unchanged (never panics).
pub fn patch_document_xml(src: &[u8], targets: &[ResolvedTarget]) -> Vec<u8> {
    let mut reader = Reader::from_reader(src);
    reader.config_mut().trim_text(false);

    let mut out: Vec<u8> = Vec::with_capacity(src.len() + 128);
    let mut cursor: usize = 0; // next source byte not yet flushed to `out`
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut p_ord: i64 = -1;
    let mut r_ord: i64 = -1;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) | Err(_) => break,
            Ok(Event::Start(e)) => {
                let ln = local_name(e.name().as_ref()).to_vec();
                let parent = stack.last().map(Vec::as_slice);
                if ln == b"p" && parent == Some(b"body".as_ref()) {
                    p_ord += 1;
                    r_ord = -1;
                }
                if ln == b"r" && parent == Some(b"p".as_ref()) {
                    r_ord += 1;
                    if let Some(t) = find_target(targets, p_ord, r_ord) {
                        // reader is positioned just after the `<w:r …>` start tag.
                        let run_open_end = reader.buffer_position() as usize;
                        splice_run(&mut reader, src, t, run_open_end, &mut out, &mut cursor);
                        continue; // run fully consumed — do not push onto the stack
                    }
                }
                stack.push(ln);
            }
            Ok(Event::Empty(e)) => {
                let ln = local_name(e.name().as_ref()).to_vec();
                let parent = stack.last().map(Vec::as_slice);
                if ln == b"p" && parent == Some(b"body".as_ref()) {
                    p_ord += 1;
                    r_ord = -1;
                }
                if ln == b"r" && parent == Some(b"p".as_ref()) {
                    r_ord += 1; // an empty `<w:r/>` has no text/rPr to patch
                }
            }
            Ok(Event::End(_)) => {
                stack.pop();
            }
            Ok(_) => {}
        }
    }

    out.extend_from_slice(&src[cursor..]);
    out
}

fn find_target(targets: &[ResolvedTarget], p: i64, r: i64) -> Option<&ResolvedTarget> {
    targets
        .iter()
        .find(|t| t.para_ord as i64 == p && t.run_ord as i64 == r)
}

/// Walk one targeted `<w:r>`'s direct children, splicing its `<w:rPr>` and/or
/// `<w:t>` and copying the rest verbatim. On return the reader is positioned just
/// after `</w:r>`, and `cursor` is advanced past every spliced hole.
fn splice_run(
    reader: &mut Reader<&[u8]>,
    src: &[u8],
    t: &ResolvedTarget,
    run_open_end: usize,
    out: &mut Vec<u8>,
    cursor: &mut usize,
) {
    // `Some` while an rPr replacement/insertion is still owed. It is placed at
    // the existing `<w:rPr>` if present, else just before the first `<w:t>`, else
    // right after the `<w:r>` open tag (schema: rPr is the run's first child).
    let mut rpr_pending: Option<&[u8]> = t.new_rpr.as_deref();
    let mut text_done = false;

    loop {
        let child_start = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Eof) | Err(_) => return,
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let ln = local_name(&name).to_vec();
                if ln == b"rPr" && t.new_rpr.is_some() {
                    let _ = reader.read_to_end(QName(&name));
                    let end = reader.buffer_position() as usize;
                    out.extend_from_slice(&src[*cursor..child_start]);
                    out.extend_from_slice(t.new_rpr.as_deref().unwrap());
                    *cursor = end;
                    rpr_pending = None;
                    continue;
                }
                if ln == b"t" && t.new_text.is_some() {
                    let _ = reader.read_to_end(QName(&name));
                    let end = reader.buffer_position() as usize;
                    // A pending rPr must land before the text (schema order).
                    if let Some(rpr) = rpr_pending.take() {
                        out.extend_from_slice(&src[*cursor..child_start]);
                        out.extend_from_slice(rpr);
                        *cursor = child_start;
                    }
                    out.extend_from_slice(&src[*cursor..child_start]);
                    if !text_done {
                        out.extend_from_slice(&render_wt(t.new_text.as_deref().unwrap()));
                        text_done = true;
                    }
                    // First `<w:t>` carries the whole new text; any later `<w:t>`
                    // in the same run is dropped (collapsed).
                    *cursor = end;
                    continue;
                }
                // A child we don't touch — skip its subtree, leave bytes verbatim.
                let _ = reader.read_to_end(QName(&name));
            }
            Ok(Event::Empty(e)) => {
                let ln = local_name(e.name().as_ref()).to_vec();
                if ln == b"rPr" && t.new_rpr.is_some() {
                    let end = reader.buffer_position() as usize;
                    out.extend_from_slice(&src[*cursor..child_start]);
                    out.extend_from_slice(t.new_rpr.as_deref().unwrap());
                    *cursor = end;
                    rpr_pending = None;
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"r" {
                    // No rPr element existed — insert it right after `<w:r>`.
                    if let Some(rpr) = rpr_pending.take() {
                        out.extend_from_slice(&src[*cursor..run_open_end]);
                        out.extend_from_slice(rpr);
                        *cursor = run_open_end;
                    }
                    return;
                }
            }
            Ok(_) => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // Two paragraphs: p0 has a plain run, p1 has a bold run with an rPr.
    const DOC: &[u8] = br#"<?xml version="1.0"?><w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t xml:space="preserve">Hello</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>bold red</w:t></w:r></w:p></w:body></w:document>"#;

    #[test]
    fn text_change_rewrites_only_the_wt_and_is_byte_identical_elsewhere() {
        let targets = vec![ResolvedTarget {
            para_ord: 0,
            run_ord: 0,
            new_text: Some("World".into()),
            new_rpr: None,
        }];
        let out = patch_document_xml(DOC, &targets);
        let expected = String::from_utf8(DOC.to_vec())
            .unwrap()
            .replace(">Hello<", ">World<");
        assert_eq!(String::from_utf8(out).unwrap(), expected);
    }

    #[test]
    fn prop_change_replaces_only_the_rpr_leaving_text_verbatim() {
        // Drop bold, keep color — the slice's edit shape.
        let new_rpr = crate::rpr::render_rpr(
            &docx_core::RunProps {
                color: Some("FF0000".into()),
                ..Default::default()
            },
            None,
        );
        let targets = vec![ResolvedTarget {
            para_ord: 1,
            run_ord: 0,
            new_text: None,
            new_rpr: Some(new_rpr),
        }];
        let out = String::from_utf8(patch_document_xml(DOC, &targets)).unwrap();
        let expected = String::from_utf8(DOC.to_vec()).unwrap().replace(
            r#"<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>"#,
            r#"<w:rPr><w:color w:val="FF0000"/></w:rPr>"#,
        );
        assert_eq!(out, expected);
        // The run's text is untouched.
        assert!(out.contains(">bold red<"));
    }

    #[test]
    fn ordinals_target_the_right_paragraph_and_run() {
        // Editing (p1, r0) must NOT touch p0's run.
        let targets = vec![ResolvedTarget {
            para_ord: 1,
            run_ord: 0,
            new_text: Some("BOLD".into()),
            new_rpr: None,
        }];
        let out = String::from_utf8(patch_document_xml(DOC, &targets)).unwrap();
        assert!(out.contains(">Hello<"), "p0 run untouched");
        assert!(out.contains(">BOLD<"));
        assert!(!out.contains(">bold red<"));
    }

    #[test]
    fn no_targets_is_byte_identical() {
        assert_eq!(patch_document_xml(DOC, &[]), DOC);
    }

    #[test]
    fn rpr_inserted_when_run_has_none() {
        let src = br#"<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#;
        let new_rpr = crate::rpr::render_rpr(
            &docx_core::RunProps {
                bold: Some(true),
                ..Default::default()
            },
            None,
        );
        let targets = vec![ResolvedTarget {
            para_ord: 0,
            run_ord: 0,
            new_text: None,
            new_rpr: Some(new_rpr),
        }];
        let out = String::from_utf8(patch_document_xml(src, &targets)).unwrap();
        assert!(out.contains("<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>"));
    }
}
