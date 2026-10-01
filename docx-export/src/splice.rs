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

use crate::rpr::{render_text, RunSpecials, SpecialSource};
use docx_core::{LINE_BREAK, SOFT_HYPHEN};

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
    /// How a refusal names this run (`"run edit skipped: block 3 run 1"`).
    pub label: String,
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
            label: format!("run edit skipped: paragraph {para_ord} run {run_ord}"),
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
    /// How a refusal names this run.
    pub label: String,
}

/// A COLUMN-level structural action on one table. Applied to the `<w:tblGrid>`
/// AND to the matching `<w:tc>` of every row, so the grid and the rows stay
/// consistent (a bare cell insert/remove would leave the row inconsistent with
/// `tblGrid`, which is why this is a column op rather than a cell op).
#[derive(Debug, Clone)]
pub enum ColumnAction {
    /// Drop grid column `col` — its `<w:gridCol>` and each row's `col`-th `<w:tc>`.
    Delete { col: u32 },
    /// Add a column after `after_col`: a copy of that `<w:gridCol>` plus a fresh
    /// `<w:tc>` carrying `text` in every row.
    Insert { after_col: u32, text: String },
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
#[cfg(test)]
pub fn patch_document_xml_rows(
    src: &[u8],
    targets: &[ResolvedTarget],
    paras: &BTreeMap<u32, ResolvedParaTarget>,
    cells: &[ResolvedCellTarget],
    rows: &BTreeMap<(u32, u32), ResolvedRowTarget>,
) -> Vec<u8> {
    patch_document_xml_cols(src, targets, paras, cells, rows, &BTreeMap::new()).0
}

/// As [`patch_document_xml_rows`], plus COLUMN actions keyed by table ordinal.
/// Returns the patched part and the run edits it refused (each under its
/// target's label, with the reason) — those runs are left byte-identical.
#[allow(clippy::too_many_arguments)]
pub fn patch_document_xml_cols(
    src: &[u8],
    targets: &[ResolvedTarget],
    paras: &BTreeMap<u32, ResolvedParaTarget>,
    cells: &[ResolvedCellTarget],
    rows: &BTreeMap<(u32, u32), ResolvedRowTarget>,
    columns: &BTreeMap<u32, ColumnAction>,
) -> (Vec<u8>, Vec<String>) {
    let mut refusals: Vec<String> = Vec::new();
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
    // Column locator: the `<w:gridCol>` index inside the open `<w:tblGrid>`.
    let mut gc_ord: i64 = -1;
    // `<w:tbl>` nesting depth. A NESTED table's rows/cells sit under their own
    // `<w:tbl>`, so without this they would be counted against the OUTER table
    // and an edit aimed at an outer row would land inside the nested one.
    // Only depth 1 (a body-level table) is addressed.
    let mut tbl_depth: i32 = 0;

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
                        splice_run(
                            &mut reader,
                            src,
                            t,
                            run_open_end,
                            &mut out,
                            &mut cursor,
                            &mut refusals,
                        );
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
                if ln == b"tbl" {
                    tbl_depth += 1;
                    if tbl_depth == 1 {
                        tbl_ord += 1;
                        tr_ord = -1;
                    }
                }
                if ln == b"tblGrid" && parent == Some(b"tbl".as_ref()) && tbl_depth == 1 {
                    gc_ord = -1;
                }
                if ln == b"tr" && parent == Some(b"tbl".as_ref()) && tbl_depth == 1 {
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
                if ln == b"tc" && parent == Some(b"tr".as_ref()) && tbl_depth == 1 {
                    tc_ord += 1;
                    cp_ord = -1;
                    // Column ops: the grid is uniform here (guarded upstream), so
                    // the `<w:tc>` index IS the grid column.
                    if let Some(action) = columns.get(&(tbl_ord as u32)) {
                        match action {
                            ColumnAction::Delete { col } if *col as i64 == tc_ord => {
                                let _ = reader.read_to_end(QName(&name));
                                let end = reader.buffer_position() as usize;
                                out.extend_from_slice(&src[cursor..event_start]);
                                cursor = end;
                                continue;
                            }
                            ColumnAction::Insert { after_col, text }
                                if *after_col as i64 == tc_ord =>
                            {
                                let _ = reader.read_to_end(QName(&name));
                                let end = reader.buffer_position() as usize;
                                // Copy the reference cell verbatim, then append a
                                // fresh one carrying the new column's text.
                                out.extend_from_slice(&src[cursor..end]);
                                out.extend_from_slice(b"<w:tc><w:p>");
                                if !text.is_empty() {
                                    out.extend_from_slice(&crate::rpr::render_run(
                                        text,
                                        &docx_core::RunProps::default(),
                                        None,
                                    ));
                                }
                                out.extend_from_slice(b"</w:p></w:tc>");
                                cursor = end;
                                continue;
                            }
                            _ => {}
                        }
                    }
                }
                if ln == b"p" && parent == Some(b"tc".as_ref()) && tbl_depth == 1 {
                    cp_ord += 1;
                    cr_ord = -1;
                }
                if ln == b"r" && parent == Some(b"p".as_ref()) && in_cell(&stack) && tbl_depth == 1
                {
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
                            label: ct.label.clone(),
                        };
                        let run_open_end = reader.buffer_position() as usize;
                        splice_run(
                            &mut reader,
                            src,
                            &t,
                            run_open_end,
                            &mut out,
                            &mut cursor,
                            &mut refusals,
                        );
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
                        splice_run(
                            &mut reader,
                            src,
                            t,
                            run_open_end,
                            &mut out,
                            &mut cursor,
                            &mut refusals,
                        );
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
                if ln == b"gridCol" && parent == Some(b"tblGrid".as_ref()) && tbl_depth == 1 {
                    gc_ord += 1;
                    if let Some(action) = columns.get(&(tbl_ord as u32)) {
                        let end = reader.buffer_position() as usize;
                        match action {
                            ColumnAction::Delete { col } if *col as i64 == gc_ord => {
                                out.extend_from_slice(&src[cursor..event_start]);
                                cursor = end;
                            }
                            ColumnAction::Insert { after_col, .. }
                                if *after_col as i64 == gc_ord =>
                            {
                                // Duplicate this `<w:gridCol>` (keeps its width).
                                out.extend_from_slice(&src[cursor..end]);
                                out.extend_from_slice(&src[event_start..end]);
                                cursor = end;
                            }
                            _ => {}
                        }
                    }
                }
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
                if ln == b"tbl" {
                    tbl_depth -= 1;
                }
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
    (out, refusals)
}

fn find_target(targets: &[ResolvedTarget], p: i64, r: i64) -> Option<&ResolvedTarget> {
    targets
        .iter()
        .find(|t| t.wrapper.is_none() && t.para_ord as i64 == p && t.run_ord as i64 == r)
}

/// What a run child is to the run's text.
enum Special {
    /// The import turned it into this ONE character of the text:
    /// `<w:br/>` with no `w:type` or `w:type="textWrapping"` and `<w:cr/>`
    /// (U+2028), `<w:tab/>` and `<w:ptab>` (`\t`), `<w:noBreakHyphen/>`
    /// (U+2011), `<w:softHyphen/>` (U+00AD), a `<w:sym>` with a Unicode
    /// equivalent (that character, [`docx_core::symbol_char`]).
    Char(char),
    /// It has no character in the text (a `<w:sym>` with no equivalent), so
    /// an edited text cannot say where it goes.
    Unplaceable,
    /// Not text at all (a page or column `w:br` is a pagination
    /// instruction; an `rPr`, a drawing, …).
    No,
}

fn attr(e: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    e.attributes()
        .flatten()
        .find(|a| local_name(a.key.as_ref()) == name)
        .map(|a| String::from_utf8_lossy(&a.value).into_owned())
}

fn special(e: &quick_xml::events::BytesStart<'_>) -> Special {
    match local_name(e.name().as_ref()) {
        b"cr" => Special::Char(LINE_BREAK),
        b"tab" | b"ptab" => Special::Char('\t'),
        b"noBreakHyphen" => Special::Char('\u{2011}'),
        b"softHyphen" => Special::Char(SOFT_HYPHEN),
        b"sym" => match docx_core::symbol_char(
            attr(e, b"font").as_deref(),
            &attr(e, b"char").unwrap_or_default(),
        ) {
            Some(c) => Special::Char(c),
            None => Special::Unplaceable,
        },
        b"br" if attr(e, b"type").is_none_or(|t| t == "textWrapping") => Special::Char(LINE_BREAK),
        _ => Special::No,
    }
}

fn is_special_char(e: &quick_xml::events::BytesStart<'_>) -> bool {
    matches!(special(e), Special::Char(_))
}

/// The text content of a `<w:t>` (the reader just past its start tag), as
/// far as special characters go.
fn wt_text(reader: &mut Reader<&[u8]>) -> String {
    let mut out = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Text(t)) => {
                if let Ok(s) = t.decode() {
                    out.push_str(&s);
                }
            }
            Ok(Event::CData(t)) => {
                if let Ok(s) = t.decode() {
                    out.push_str(&s);
                }
            }
            Ok(Event::GeneralRef(r)) => {
                // A character reference may be a special character (`&#9;`);
                // the predefined entities never are, so they are not needed.
                if let Ok(Some(c)) = r.resolve_char_ref() {
                    out.push(c);
                }
            }
            Ok(Event::End(_)) | Ok(Event::Eof) | Err(_) => return out,
            Ok(_) => {}
        }
    }
}

/// Read the children of the run whose open tag ends at `run_open_end`: how
/// the run writes each special character of its text ([`RunSpecials`]), or
/// why its text cannot be rewritten safely.
///
/// The new text is written once, where the run's first text-bearing child
/// (`<w:t>` or a special element) was, and every later one is dropped. So a
/// child that is neither, sitting BETWEEN two of them (a `<w:drawing>`, a
/// page break, a `<w:footnoteReference>`), would move behind the text: that
/// is refused. `<w:lastRenderedPageBreak/>` is only Word's layout cache, so
/// it may move. And content the import drops ([`Special::Unplaceable`]) is refused
/// wherever it sits.
fn scan_run(src: &[u8], run_open_end: usize) -> Result<RunSpecials<'_>, String> {
    let body = &src[run_open_end..];
    let mut reader = Reader::from_reader(body);
    reader.config_mut().trim_text(false);
    // The run's text-bearing children in order: a `<w:t>`'s text, or a
    // special element's character and bytes. Which characters count as
    // special is only known at the end (a `<w:sym>`'s character may also
    // stand literally in an earlier `<w:t>`), so the specials are built
    // from this list.
    enum Piece<'a> {
        Text(String),
        Element(char, &'a [u8]),
    }
    let mut pieces: Vec<Piece<'_>> = Vec::new();
    let mut first_text: Option<usize> = None;
    let mut last_text: Option<usize> = None;
    let mut opaque: Vec<(usize, String)> = Vec::new();
    let mut k = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let (e, empty) = match reader.read_event() {
            Ok(Event::Start(e)) => (e.into_owned(), false),
            Ok(Event::Empty(e)) => (e.into_owned(), true),
            Ok(Event::End(_)) | Ok(Event::Eof) | Err(_) => break,
            Ok(_) => continue,
        };
        k += 1;
        let name = e.name().as_ref().to_vec();
        let ln = local_name(&name).to_vec();
        let mut text_bearing = false;
        if ln == b"t" {
            if !empty {
                pieces.push(Piece::Text(wt_text(&mut reader)));
            }
            text_bearing = true;
        } else {
            if !empty {
                let _ = reader.read_to_end(QName(&name));
            }
            let bytes = &body[start..reader.buffer_position() as usize];
            let kind = special(&e);
            if let Special::Char(c) = kind {
                pieces.push(Piece::Element(c, bytes));
                text_bearing = true;
            } else if matches!(kind, Special::Unplaceable) {
                return Err(format!(
                    "it holds a <w:{}> with no character in the text ({}), so the \
                     edited text has no place for it",
                    String::from_utf8_lossy(&ln),
                    String::from_utf8_lossy(bytes)
                ));
            } else if ln != b"rPr" && ln != b"lastRenderedPageBreak" {
                opaque.push((k, String::from_utf8_lossy(&name).into_owned()));
            }
        }
        if text_bearing {
            first_text.get_or_insert(k);
            last_text = Some(k);
        }
    }
    if let (Some(first), Some(last)) = (first_text, last_text) {
        if let Some((_, name)) = opaque.iter().find(|(k, _)| first < *k && *k < last) {
            return Err(format!(
                "its <{name}> sits inside the run's text, and the edited text has no \
                 place for it"
            ));
        }
    }
    let tracked = |c: char| {
        crate::rpr::word_element(c).is_some()
            || pieces
                .iter()
                .any(|p| matches!(p, Piece::Element(e, _) if *e == c))
    };
    let mut specials = RunSpecials::default();
    for piece in &pieces {
        match piece {
            Piece::Text(t) => {
                for c in t.chars().filter(|c| tracked(*c)) {
                    specials.push(c, SpecialSource::Literal);
                }
            }
            Piece::Element(c, bytes) => specials.push(*c, SpecialSource::Element(bytes)),
        }
    }
    Ok(specials)
}

/// Copy the rest of the run verbatim: consume its children up to `</w:r>`
/// without moving `cursor`.
fn skip_run(reader: &mut Reader<&[u8]>) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let _ = reader.read_to_end(QName(&name));
            }
            Ok(Event::End(_)) | Ok(Event::Eof) | Err(_) => return,
            Ok(_) => {}
        }
    }
}

/// A run has its `<w:rPr>` as its FIRST child (schema). When the run has
/// none and one is owed, it goes in before the first child that is not one,
/// whatever that child is (a `<w:lastRenderedPageBreak/>`, a `<w:tab/>`).
fn place_rpr_first(
    ln: &[u8],
    src: &[u8],
    child_start: usize,
    out: &mut Vec<u8>,
    cursor: &mut usize,
    rpr_pending: &mut Option<&[u8]>,
) {
    if ln == b"rPr" {
        return;
    }
    if let Some(rpr) = rpr_pending.take() {
        out.extend_from_slice(&src[*cursor..child_start]);
        out.extend_from_slice(rpr);
        *cursor = child_start;
    }
}

/// Walk one targeted `<w:r>`'s direct children, splicing its `<w:rPr>` and/or
/// `<w:t>` and copying the rest verbatim. On return the reader is positioned just
/// after `</w:r>`, and `cursor` is advanced past every spliced hole.
///
/// A new text replaces ALL of the run's `<w:t>` AND its special children
/// (line breaks, tabs, non-breaking hyphens — the text carries those as
/// characters): it is written once, where the first of them was (or before
/// `</w:r>` when the run had none), with the run's original elements
/// re-emitted in order for its special characters ([`render_text`]). A run
/// whose text cannot be rewritten that way ([`scan_run`]) is left untouched
/// and the refusal is recorded in `refusals` under the target's label.
fn splice_run(
    reader: &mut Reader<&[u8]>,
    src: &[u8],
    t: &ResolvedTarget,
    run_open_end: usize,
    out: &mut Vec<u8>,
    cursor: &mut usize,
    refusals: &mut Vec<String>,
) {
    // `Some` while an rPr replacement/insertion is still owed. It is placed at
    // the existing `<w:rPr>` if present, else just before the text, else
    // right after the `<w:r>` open tag (schema: rPr is the run's first child).
    let mut rpr_pending: Option<&[u8]> = t.new_rpr.as_deref();
    let mut text_done = false;
    let originals = if t.new_text.is_some() {
        match scan_run(src, run_open_end) {
            Ok(o) => o,
            Err(reason) => {
                refusals.push(format!("{}: {reason}", t.label));
                skip_run(reader);
                return;
            }
        }
    } else {
        RunSpecials::default()
    };
    let rendered = || render_text(t.new_text.as_deref().unwrap_or(""), &originals);

    loop {
        let child_start = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Eof) | Err(_) => return,
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let ln = local_name(&name).to_vec();
                place_rpr_first(&ln, src, child_start, out, cursor, &mut rpr_pending);
                if ln == b"rPr" && t.new_rpr.is_some() {
                    let _ = reader.read_to_end(QName(&name));
                    let end = reader.buffer_position() as usize;
                    out.extend_from_slice(&src[*cursor..child_start]);
                    out.extend_from_slice(t.new_rpr.as_deref().unwrap());
                    *cursor = end;
                    rpr_pending = None;
                    continue;
                }
                if t.new_text.is_some() && (ln == b"t" || is_special_char(&e)) {
                    let _ = reader.read_to_end(QName(&name));
                    let end = reader.buffer_position() as usize;
                    out.extend_from_slice(&src[*cursor..child_start]);
                    // A pending rPr must land before the text (schema order).
                    if let Some(rpr) = rpr_pending.take() {
                        out.extend_from_slice(rpr);
                    }
                    if !text_done {
                        out.extend_from_slice(&rendered());
                        text_done = true;
                    }
                    // The first text-bearing child carries the whole new text;
                    // every later one in the same run is dropped (collapsed).
                    *cursor = end;
                    continue;
                }
                // A child we don't touch — skip its subtree, leave bytes verbatim.
                let _ = reader.read_to_end(QName(&name));
            }
            Ok(Event::Empty(e)) => {
                let ln = local_name(e.name().as_ref()).to_vec();
                let end = reader.buffer_position() as usize;
                place_rpr_first(&ln, src, child_start, out, cursor, &mut rpr_pending);
                if ln == b"rPr" && t.new_rpr.is_some() {
                    out.extend_from_slice(&src[*cursor..child_start]);
                    out.extend_from_slice(t.new_rpr.as_deref().unwrap());
                    *cursor = end;
                    rpr_pending = None;
                } else if t.new_text.is_some() && (ln == b"t" || is_special_char(&e)) {
                    out.extend_from_slice(&src[*cursor..child_start]);
                    if let Some(rpr) = rpr_pending.take() {
                        out.extend_from_slice(rpr);
                    }
                    if !text_done {
                        out.extend_from_slice(&rendered());
                        text_done = true;
                    }
                    *cursor = end;
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
                    // No text-bearing child to write the new text at: it goes
                    // last.
                    if t.new_text.is_some() && !text_done {
                        out.extend_from_slice(&src[*cursor..child_start]);
                        out.extend_from_slice(&rendered());
                        *cursor = child_start;
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

    fn patch_one(src: &[u8], text: &str) -> String {
        let mut t = ResolvedTarget::edit(0, 0);
        t.new_text = Some(text.into());
        String::from_utf8(patch_document_xml(src, &[t])).unwrap()
    }

    #[test]
    fn line_breaks_patch_back_as_w_br_in_place() {
        // U+2028 in the edited text is Word's `<w:br/>`; the run's OWN
        // break elements are re-emitted verbatim (the `w:clear`, the `w:cr`).
        let src = br#"<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>one</w:t><w:br w:clear="all"/><w:t>two</w:t><w:cr/><w:t>three</w:t></w:r></w:p></w:body></w:document>"#;
        let out = patch_one(src, "one\u{2028}TWO\u{2028}three");
        assert!(
            out.contains(r#"<w:r><w:t xml:space="preserve">one</w:t><w:br w:clear="all"/><w:t xml:space="preserve">TWO</w:t><w:cr/><w:t xml:space="preserve">three</w:t></w:r>"#),
            "{out}"
        );
        // A break added by the edit is a plain `<w:br/>`.
        let out = patch_one(src, "one\u{2028}two\u{2028}three\u{2028}four");
        assert!(out.contains(r#"<w:t xml:space="preserve">three</w:t><w:br/><w:t xml:space="preserve">four</w:t></w:r>"#), "{out}");
        // A break removed by the edit is gone, the others keep their bytes.
        let out = patch_one(src, "one\u{2028}two three");
        assert!(out.contains(r#"<w:r><w:t xml:space="preserve">one</w:t><w:br w:clear="all"/><w:t xml:space="preserve">two three</w:t></w:r>"#), "{out}");
        assert!(!out.contains("<w:cr/>"), "{out}");
    }

    #[test]
    fn line_breaks_at_run_edges_and_page_breaks() {
        // A break before the first `<w:t>` and a run of only breaks; a page
        // break is pagination, not text, and is left where it is.
        let src = br#"<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:br/><w:t>x</w:t><w:br w:type="page"/></w:r><w:r><w:br/><w:br w:type="textWrapping"/></w:r></w:p></w:body></w:document>"#;
        let mut a = ResolvedTarget::edit(0, 0);
        a.new_text = Some("\u{2028}y".into());
        let mut b = ResolvedTarget::edit(0, 1);
        b.new_text = Some("\u{2028}z\u{2028}".into());
        let out = String::from_utf8(patch_document_xml(src, &[a, b])).unwrap();
        assert!(
            out.contains(
                r#"<w:r><w:br/><w:t xml:space="preserve">y</w:t><w:br w:type="page"/></w:r>"#
            ),
            "{out}"
        );
        assert!(
            out.contains(r#"<w:r><w:br/><w:t xml:space="preserve">z</w:t><w:br w:type="textWrapping"/></w:r>"#),
            "{out}"
        );
    }

    #[test]
    fn tabs_patch_back_as_w_tab_in_place() {
        // A tab is `\t` in the text and `<w:tab/>` in the run: the edit keeps
        // the element where it was and writes NO tab character into `<w:t>`.
        let src = br#"<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p></w:body></w:document>"#;
        let out = patch_one(src, "a\tB");
        assert!(
            out.contains(r#"<w:r><w:t xml:space="preserve">a</w:t><w:tab/><w:t xml:space="preserve">B</w:t></w:r>"#),
            "{out}"
        );
        // A tab removed by the edit is gone; one added is `<w:tab/>`.
        assert!(patch_one(src, "ab").contains(r#"<w:r><w:t xml:space="preserve">ab</w:t></w:r>"#));
        assert!(patch_one(src, "a\tb\tc").contains(
            r#"<w:t xml:space="preserve">b</w:t><w:tab/><w:t xml:space="preserve">c</w:t></w:r>"#
        ));
    }

    #[test]
    fn content_inside_the_text_refuses_the_edit_and_keeps_the_run() {
        // A footnote reference between two pieces of text would move behind
        // the edited text; a `w:sym` with no Unicode equivalent (Wingdings'
        // Windows logo) has no character to place at all.
        for run in [
            r#"<w:r><w:t>a</w:t><w:footnoteReference w:id="1"/><w:t>b</w:t></w:r>"#,
            r#"<w:r><w:sym w:font="Wingdings" w:char="F0FF"/><w:t>b</w:t></w:r>"#,
            r#"<w:r><w:t>a</w:t><w:br w:type="page"/><w:t>b</w:t></w:r>"#,
        ] {
            let src = format!(
                r#"<w:document xmlns:w="urn:w"><w:body><w:p>{run}</w:p></w:body></w:document>"#
            );
            let mut t = ResolvedTarget::edit(0, 0);
            t.new_text = Some("xy".into());
            t.new_rpr = Some(b"<w:rPr><w:b/></w:rPr>".to_vec());
            let (out, refused) = patch_document_xml_cols(
                src.as_bytes(),
                &[t],
                &BTreeMap::new(),
                &[],
                &BTreeMap::new(),
                &BTreeMap::new(),
            );
            assert_eq!(out, src.as_bytes(), "{run}: the run keeps its bytes");
            assert_eq!(refused.len(), 1, "{run}");
            assert!(
                refused[0].starts_with("run edit skipped: paragraph 0 run 0: "),
                "{refused:?}"
            );
        }
        // Word's layout cache may move; content outside the text stays put.
        let src = br#"<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:lastRenderedPageBreak/><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t><w:footnoteReference w:id="1"/></w:r></w:p></w:body></w:document>"#;
        let mut t = ResolvedTarget::edit(0, 0);
        t.new_text = Some("xy".into());
        t.new_rpr = Some(b"<w:rPr><w:b/></w:rPr>".to_vec());
        let out = String::from_utf8(patch_document_xml(src, &[t])).unwrap();
        // The rPr goes FIRST (schema), ahead of the layout cache.
        assert!(
            out.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:lastRenderedPageBreak/><w:t xml:space="preserve">xy</w:t><w:lastRenderedPageBreak/><w:footnoteReference w:id="1"/></w:r>"#),
            "{out}"
        );
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
