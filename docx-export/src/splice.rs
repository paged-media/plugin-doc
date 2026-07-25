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

use std::collections::BTreeMap;

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
    /// Increment 2 — drop the whole `<w:r>` subtree instead of editing it.
    pub delete: bool,
    /// Increment 2 — pre-rendered `<w:r>…</w:r>` fragments to emit immediately
    /// after this run's `</w:r>`.
    pub insert_after: Vec<Vec<u8>>,
    /// Increment 3 — when set, `run_ord` counts `<w:r>` inside the n-th
    /// `<w:hyperlink>` / `<w:fldSimple>` child instead of the paragraph's direct
    /// `<w:r>` children.
    pub wrapper: Option<crate::bindings::Wrapper>,
}

impl ResolvedTarget {
    /// A target that only edits (no structural change).
    pub fn edit(para_ord: u32, run_ord: u32) -> Self {
        ResolvedTarget {
            para_ord,
            run_ord,
            new_text: None,
            new_rpr: None,
            delete: false,
            insert_after: Vec::new(),
            wrapper: None,
        }
    }
}

/// A TABLE-CELL run to patch: the `w:tbl`/`w:tr`/`w:tc`/`w:p` path + the run
/// ordinal within that cell paragraph, plus the replacements.
#[derive(Debug, Clone)]
pub struct ResolvedCellTarget {
    pub table_ord: u32,
    pub row: u32,
    pub cell: u32,
    pub para: u32,
    pub run_ord: u32,
    pub new_text: Option<String>,
    pub new_rpr: Option<Vec<u8>>,
}

/// A ROW-level structural action, addressed by `(table_ord, row)`.
#[derive(Debug, Clone, Default)]
pub struct ResolvedRowTarget {
    /// Drop the whole `<w:tr>` subtree.
    pub delete: bool,
    /// Pre-rendered `<w:tr>…</w:tr>` fragments to emit after this row's `</w:tr>`.
    pub insert_after: Vec<Vec<u8>>,
}

/// A paragraph-level structural action, addressed by `<w:p>` ordinal.
#[derive(Debug, Clone, Default)]
pub struct ResolvedParaTarget {
    /// Drop the whole `<w:p>` subtree.
    pub delete: bool,
    /// Pre-rendered `<w:p>…</w:p>` fragments to emit after this paragraph's
    /// `</w:p>`.
    pub insert_after: Vec<Vec<u8>>,
    /// Pre-rendered `<w:r>…</w:r>` fragments to emit at the START of this
    /// paragraph's content (used by `InsertRun { run: None }`).
    pub prepend_runs: Vec<Vec<u8>>,
    /// Increment 3 — pre-rendered `<w:pPr>…</w:pPr>` bytes replacing (or, when
    /// the paragraph has none, inserted as) the paragraph's first child.
    pub new_ppr: Option<Vec<u8>>,
}

/// Are we inside a table cell? (The body-paragraph counters must ignore cell
/// content, and vice versa.)
fn in_cell(stack: &[Vec<u8>]) -> bool {
    stack.iter().any(|n| n.as_slice() == b"tc")
}

fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|&b| b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

/// Run-edits-only convenience over [`patch_document_xml_full`] (tests).
#[cfg(test)]
pub fn patch_document_xml(src: &[u8], targets: &[ResolvedTarget]) -> Vec<u8> {
    patch_document_xml_all(src, targets, &BTreeMap::new(), &[])
}

/// Body-only convenience over [`patch_document_xml_all`] (tests): run targets +
/// paragraph-level structural actions keyed by `<w:p>` ordinal.
#[cfg(test)]
pub fn patch_document_xml_full(
    src: &[u8],
    targets: &[ResolvedTarget],
    paras: &BTreeMap<u32, ResolvedParaTarget>,
) -> Vec<u8> {
    patch_document_xml_all(src, targets, paras, &[])
}

/// As [`patch_document_xml_full`], plus TABLE-CELL run targets (their own
/// `w:tbl`/`w:tr`/`w:tc`/`w:p` locator path).
///
/// NOTE: the cell counters assume tables are not NESTED (a `<w:tbl>` inside a
/// `<w:tc>`); a nested table's rows would be counted against the outer table.
/// Nested-table cell content is therefore not patched — the bindings only ever
/// address top-level tables.
#[cfg(test)]
pub fn patch_document_xml_all(
    src: &[u8],
    targets: &[ResolvedTarget],
    paras: &BTreeMap<u32, ResolvedParaTarget>,
    cells: &[ResolvedCellTarget],
) -> Vec<u8> {
    patch_document_xml_rows(src, targets, paras, cells, &BTreeMap::new())
}

/// As [`patch_document_xml_all`], plus ROW-level structural actions keyed by
/// `(table_ord, row)`.
pub fn patch_document_xml_rows(
    src: &[u8],
    targets: &[ResolvedTarget],
    paras: &BTreeMap<u32, ResolvedParaTarget>,
    cells: &[ResolvedCellTarget],
    rows: &BTreeMap<(u32, u32), ResolvedRowTarget>,
) -> Vec<u8> {
    let mut reader = Reader::from_reader(src);
    reader.config_mut().trim_text(false);

    let mut out: Vec<u8> = Vec::with_capacity(src.len() + 128);
    let mut cursor: usize = 0; // next source byte not yet flushed to `out`
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut p_ord: i64 = -1;
    let mut r_ord: i64 = -1;
    // Table-cell locator counters (see the nesting note above).
    let mut tbl_ord: i64 = -1;
    let mut tr_ord: i64 = -1;
    let mut tc_ord: i64 = -1;
    let mut cp_ord: i64 = -1;
    let mut cr_ord: i64 = -1;
    // The paragraph currently open at body level, if it carries actions.
    let mut open_para: Option<(u32, ResolvedParaTarget)> = None;
    let mut prepended = false;
    // Whether the open paragraph's `<w:pPr>` action has been discharged.
    let mut ppr_done = false;
    // Wrapper locator (`<w:hyperlink>` / `<w:fldSimple>` children of a body
    // paragraph, and the `<w:r>` index inside the open wrapper).
    let mut hl_ord: i64 = -1;
    let mut fld_ord: i64 = -1;
    let mut wrap_run_ord: i64 = -1;
    let mut open_wrapper: Option<crate::bindings::Wrapper> = None;
    let mut open_row: Option<ResolvedRowTarget> = None;

    loop {
        // The byte offset of the `<` that begins the event we are about to read
        // — the anchor every structural splice needs.
        let event_start = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Eof) | Err(_) => break,
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let ln = local_name(&name).to_vec();
                let parent = stack.last().map(Vec::as_slice);
                if ln == b"p" && parent == Some(b"body".as_ref()) {
                    p_ord += 1;
                    r_ord = -1;
                    hl_ord = -1;
                    fld_ord = -1;
                    open_wrapper = None;
                    let para_start = event_start;
                    if let Some(pt) = paras.get(&(p_ord as u32)) {
                        if pt.delete {
                            // Drop the whole `<w:p>` subtree.
                            let _ = reader.read_to_end(QName(&name));
                            let end = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..para_start]);
                            // Any paragraphs to add still land here.
                            for frag in &pt.insert_after {
                                out.extend_from_slice(frag);
                            }
                            cursor = end;
                            continue;
                        }
                        open_para = Some((p_ord as u32, pt.clone()));
                        prepended = false;
                        ppr_done = false;
                    } else {
                        open_para = None;
                    }
                }
                // Increment 3 — wrapper elements whose runs carry their own address.
                if parent == Some(b"p".as_ref()) && !in_cell(&stack) {
                    if ln == b"hyperlink" {
                        hl_ord += 1;
                        wrap_run_ord = -1;
                        open_wrapper = Some(crate::bindings::Wrapper::Hyperlink(hl_ord as u32));
                    } else if ln == b"fldSimple" {
                        fld_ord += 1;
                        wrap_run_ord = -1;
                        open_wrapper = Some(crate::bindings::Wrapper::Field(fld_ord as u32));
                    }
                }
                // A `<w:r>` inside an open wrapper resolves on the wrapper path.
                if ln == b"r" && open_wrapper.is_some() && parent != Some(b"p".as_ref()) {
                    wrap_run_ord += 1;
                    if let Some(t) = targets.iter().find(|t| {
                        t.wrapper == open_wrapper
                            && t.para_ord as i64 == p_ord
                            && t.run_ord as i64 == wrap_run_ord
                    }) {
                        let run_open_end = reader.buffer_position() as usize;
                        splice_run(&mut reader, src, t, run_open_end, &mut out, &mut cursor);
                        continue;
                    }
                }
                // Increment 3 — replace a targeted paragraph's `<w:pPr>`.
                if ln == b"pPr" && parent == Some(b"p".as_ref()) && !in_cell(&stack) {
                    if let Some((_, pt)) = open_para.as_ref() {
                        if let Some(ppr) = pt.new_ppr.clone() {
                            let _ = reader.read_to_end(QName(&name));
                            let end = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..event_start]);
                            out.extend_from_slice(&ppr);
                            cursor = end;
                            ppr_done = true;
                            continue;
                        }
                    }
                }
                // --- table-cell locator path (tbl → tr → tc → p → r) ---
                if ln == b"tbl" && parent == Some(b"body".as_ref()) {
                    tbl_ord += 1;
                    tr_ord = -1;
                }
                if ln == b"tr" && parent == Some(b"tbl".as_ref()) {
                    tr_ord += 1;
                    tc_ord = -1;
                    open_row = rows.get(&(tbl_ord as u32, tr_ord as u32)).cloned();
                    if let Some(rt) = open_row.as_ref() {
                        if rt.delete {
                            let _ = reader.read_to_end(QName(&name));
                            let end = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..event_start]);
                            for frag in &rt.insert_after {
                                out.extend_from_slice(frag);
                            }
                            cursor = end;
                            open_row = None;
                            continue;
                        }
                    }
                }
                if ln == b"tc" && parent == Some(b"tr".as_ref()) {
                    tc_ord += 1;
                    cp_ord = -1;
                }
                if ln == b"p" && parent == Some(b"tc".as_ref()) {
                    cp_ord += 1;
                    cr_ord = -1;
                }
                if ln == b"r" && parent == Some(b"p".as_ref()) && in_cell(&stack) {
                    cr_ord += 1;
                    if let Some(ct) = cells.iter().find(|c| {
                        c.table_ord as i64 == tbl_ord
                            && c.row as i64 == tr_ord
                            && c.cell as i64 == tc_ord
                            && c.para as i64 == cp_ord
                            && c.run_ord as i64 == cr_ord
                    }) {
                        let t = ResolvedTarget {
                            para_ord: 0,
                            run_ord: 0,
                            new_text: ct.new_text.clone(),
                            new_rpr: ct.new_rpr.clone(),
                            delete: false,
                            insert_after: Vec::new(),
                            wrapper: None,
                        };
                        let run_open_end = reader.buffer_position() as usize;
                        splice_run(&mut reader, src, &t, run_open_end, &mut out, &mut cursor);
                        continue;
                    }
                }
                if ln == b"r" && parent == Some(b"p".as_ref()) && !in_cell(&stack) {
                    r_ord += 1;
                    // A pending pPr insert + prepends land before the first run.
                    if let Some((_, pt)) = open_para.as_ref() {
                        if !ppr_done {
                            if let Some(ppr) = pt.new_ppr.as_deref() {
                                let at = event_start;
                                out.extend_from_slice(&src[cursor..at]);
                                out.extend_from_slice(ppr);
                                cursor = at;
                            }
                            ppr_done = true;
                        }
                        if !prepended && !pt.prepend_runs.is_empty() {
                            let at = event_start;
                            out.extend_from_slice(&src[cursor..at]);
                            for frag in &pt.prepend_runs {
                                out.extend_from_slice(frag);
                            }
                            cursor = at;
                            prepended = true;
                        }
                    }
                    if let Some(t) = find_target(targets, p_ord, r_ord) {
                        let run_start = event_start;
                        if t.delete {
                            let _ = reader.read_to_end(QName(&name));
                            let end = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..run_start]);
                            for frag in &t.insert_after {
                                out.extend_from_slice(frag);
                            }
                            cursor = end;
                            continue;
                        }
                        // reader is positioned just after the `<w:r …>` start tag.
                        let run_open_end = reader.buffer_position() as usize;
                        splice_run(&mut reader, src, t, run_open_end, &mut out, &mut cursor);
                        if !t.insert_after.is_empty() {
                            let after = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..after]);
                            for frag in &t.insert_after {
                                out.extend_from_slice(frag);
                            }
                            cursor = after;
                        }
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
                    hl_ord = -1;
                    fld_ord = -1;
                    open_wrapper = None;
                }
                if ln == b"r" && parent == Some(b"p".as_ref()) {
                    r_ord += 1; // an empty `<w:r/>` has no text/rPr to patch
                }
            }
            Ok(Event::End(e)) => {
                let ln = local_name(e.name().as_ref()).to_vec();
                stack.pop();
                if ln == b"hyperlink" || ln == b"fldSimple" {
                    open_wrapper = None;
                }
                if ln == b"tr" {
                    if let Some(rt) = open_row.take() {
                        if !rt.insert_after.is_empty() {
                            let after = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..after]);
                            for frag in &rt.insert_after {
                                out.extend_from_slice(frag);
                            }
                            cursor = after;
                        }
                    }
                }
                // A body paragraph just closed — emit any paragraphs queued to
                // follow it (the reader is positioned just past `</w:p>`).
                if ln == b"p" && stack.last().map(Vec::as_slice) == Some(b"body".as_ref()) {
                    if let Some((_, pt)) = open_para.take() {
                        if !pt.insert_after.is_empty() {
                            let after = reader.buffer_position() as usize;
                            out.extend_from_slice(&src[cursor..after]);
                            for frag in &pt.insert_after {
                                out.extend_from_slice(frag);
                            }
                            cursor = after;
                        }
                    }
                }
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
        .find(|t| t.wrapper.is_none() && t.para_ord as i64 == p && t.run_ord as i64 == r)
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
        let targets = {
            let mut t = ResolvedTarget::edit(0, 0);
            t.new_text = Some("World".into());
            vec![t]
        };
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
        let targets = {
            let mut t = ResolvedTarget::edit(1, 0);
            t.new_rpr = Some(new_rpr);
            vec![t]
        };
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
        let targets = {
            let mut t = ResolvedTarget::edit(1, 0);
            t.new_text = Some("BOLD".into());
            vec![t]
        };
        let out = String::from_utf8(patch_document_xml(DOC, &targets)).unwrap();
        assert!(out.contains(">Hello<"), "p0 run untouched");
        assert!(out.contains(">BOLD<"));
        assert!(!out.contains(">bold red<"));
    }

    #[test]
    fn multi_wt_run_collapses_into_one() {
        // A run whose text is split across several `<w:t>` children (Word does
        // this after edits/spell-check). `LoweredRun.text` is the CONCATENATION,
        // so replacing the run's text must replace ALL of them — the first `<w:t>`
        // carries the new text and the rest are dropped, never duplicated or left
        // stale.
        let src = br#"<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Hel</w:t><w:t>lo</w:t><w:t> there</w:t></w:r></w:p></w:body></w:document>"#;
        let mut t = ResolvedTarget::edit(0, 0);
        t.new_text = Some("Replaced".into());
        let out = String::from_utf8(patch_document_xml(src, &[t])).unwrap();
        assert!(
            out.contains("<w:r><w:t xml:space=\"preserve\">Replaced</w:t></w:r>"),
            "one <w:t> carries the whole new text:\n{out}"
        );
        assert!(!out.contains(">Hel<"), "stale fragment gone");
        assert!(!out.contains(">lo<"), "stale fragment gone");
        assert!(!out.contains("> there<"), "stale fragment gone");
        assert_eq!(out.matches("<w:t").count(), 1, "exactly one <w:t> remains");
    }

    #[test]
    fn no_targets_is_byte_identical() {
        assert_eq!(patch_document_xml(DOC, &[]), DOC);
    }

    #[test]
    fn delete_run_drops_the_whole_subtree() {
        let mut t = ResolvedTarget::edit(1, 0);
        t.delete = true;
        let out = String::from_utf8(patch_document_xml(DOC, &[t])).unwrap();
        assert!(!out.contains("bold red"), "run's text gone");
        assert!(!out.contains("<w:b/>"), "run's rPr gone");
        assert!(
            out.contains("<w:p></w:p>"),
            "the paragraph remains, now empty"
        );
        assert!(out.contains(">Hello<"), "the other paragraph is untouched");
    }

    #[test]
    fn insert_run_after_places_the_fragment() {
        let mut t = ResolvedTarget::edit(0, 0);
        t.insert_after.push(crate::rpr::render_run(
            "added",
            &docx_core::RunProps::default(),
            None,
        ));
        let out = String::from_utf8(patch_document_xml(DOC, &[t])).unwrap();
        assert!(
            out.contains(
                "<w:t xml:space=\"preserve\">Hello</w:t></w:r><w:r><w:t xml:space=\"preserve\">added</w:t></w:r>"
            ),
            "new run follows the existing one:\n{out}"
        );
    }

    #[test]
    fn delete_and_insert_paragraphs() {
        use std::collections::BTreeMap;
        let mut paras: BTreeMap<u32, ResolvedParaTarget> = BTreeMap::new();
        // Delete p0; append a new paragraph after p1.
        paras.entry(0).or_default().delete = true;
        paras
            .entry(1)
            .or_default()
            .insert_after
            .push(crate::rpr::render_paragraph(
                "new para",
                &docx_core::RunProps::default(),
                None,
                None,
            ));
        let out = String::from_utf8(patch_document_xml_full(DOC, &[], &paras)).unwrap();
        assert!(!out.contains(">Hello<"), "p0 deleted");
        assert!(out.contains(">bold red<"), "p1 survives");
        assert!(
            out.contains("</w:p><w:p><w:r><w:t xml:space=\"preserve\">new para</w:t></w:r></w:p>"),
            "new paragraph appended after p1:\n{out}"
        );
        assert!(out.contains("</w:body>"), "document structure intact");
    }

    #[test]
    fn prepend_run_lands_before_the_first_run() {
        use std::collections::BTreeMap;
        let mut paras: BTreeMap<u32, ResolvedParaTarget> = BTreeMap::new();
        paras
            .entry(0)
            .or_default()
            .prepend_runs
            .push(crate::rpr::render_run(
                "first! ",
                &docx_core::RunProps::default(),
                None,
            ));
        let out = String::from_utf8(patch_document_xml_full(DOC, &[], &paras)).unwrap();
        assert!(
            out.contains("<w:p><w:r><w:t xml:space=\"preserve\">first! </w:t></w:r><w:r>"),
            "prepended run precedes the original:\n{out}"
        );
        assert!(out.contains(">Hello<"), "original run kept");
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
        let targets = {
            let mut t = ResolvedTarget::edit(0, 0);
            t.new_rpr = Some(new_rpr);
            vec![t]
        };
        let out = String::from_utf8(patch_document_xml(src, &targets)).unwrap();
        assert!(out.contains("<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>"));
    }
}
