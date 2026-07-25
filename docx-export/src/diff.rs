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

//! `diff(base, edited) -> EditSet`: the seam the live readback door will target
//! (it will re-lower the edited native model to a `LoweredDoc` and diff it against
//! the import-time baseline). Increment 1 covers STRUCTURE-PRESERVING edits (same
//! block/run counts); a count mismatch skips that block (structural insert/delete
//! is a later increment).
//!
//! Runs are compared on their RESOLVED effective formatting, not their lowered
//! style-id strings: synthesized `docx-auto-cN` ids are positional and renumber
//! between two independent lowerings, so only the inverted `RunProps`/`rStyle`
//! comparison is stable.

use docx_core::{RunProps, VertAlign};
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue, StyleProp};

use crate::bindings::DocxBindings;
use crate::edit::{CellRunEdit, EditSet, ParaEdit, RunEdit, StructuralEdit};

/// Diff two lowerings into the edits needed to turn `base` into `edited`.
pub fn diff(base: &LoweredDoc, edited: &LoweredDoc, bindings: &DocxBindings) -> EditSet {
    let mut runs = Vec::new();
    let mut structural: Vec<StructuralEdit> = Vec::new();
    let mut cells: Vec<CellRunEdit> = Vec::new();
    let mut paragraphs: Vec<ParaEdit> = Vec::new();

    // Increment 2 — paragraph-level structure. Blocks beyond the edited story's
    // length were deleted; blocks the edited story adds at the end are appended
    // after the last shared paragraph block.
    if edited.story.blocks.len() < base.story.blocks.len() {
        for (block_idx, bb) in base
            .story
            .blocks
            .iter()
            .enumerate()
            .skip(edited.story.blocks.len())
        {
            if matches!(bb, LoweredBlock::Paragraph(_)) {
                structural.push(StructuralEdit::DeleteParagraph { block: block_idx });
            }
        }
    } else if edited.story.blocks.len() > base.story.blocks.len() {
        // Anchor appended paragraphs after the LAST baseline paragraph block.
        let anchor = base
            .story
            .blocks
            .iter()
            .enumerate()
            .rev()
            .find(|(_, b)| matches!(b, LoweredBlock::Paragraph(_)))
            .map(|(i, _)| i);
        if let Some(anchor) = anchor {
            for eb in edited.story.blocks.iter().skip(base.story.blocks.len()) {
                if let LoweredBlock::Paragraph(ep) = eb {
                    let text: String = ep.runs.iter().map(|r| r.text.as_str()).collect();
                    let (props, rstyle) = ep
                        .runs
                        .first()
                        .map(|r| effective_props(r.char_style_id.as_deref(), edited, bindings))
                        .unwrap_or_default();
                    structural.push(StructuralEdit::InsertParagraph {
                        block: anchor,
                        text,
                        props,
                        para_style: ep.para_style_id.clone(),
                        rstyle,
                    });
                }
            }
        }
    }

    for (block_idx, (bb, eb)) in base
        .story
        .blocks
        .iter()
        .zip(&edited.story.blocks)
        .enumerate()
    {
        // Tables: diff cell content (text/style per cell paragraph run).
        if let (LoweredBlock::Table(bt), LoweredBlock::Table(et)) = (bb, eb) {
            for (cell_idx, (bc, ec)) in bt.cells.iter().zip(&et.cells).enumerate() {
                for (para_idx, (bpara, epara)) in
                    bc.paragraphs.iter().zip(&ec.paragraphs).enumerate()
                {
                    if bpara.runs.len() != epara.runs.len() {
                        continue; // cell structure changed — deferred
                    }
                    for (run_idx, (br, er)) in bpara.runs.iter().zip(&epara.runs).enumerate() {
                        let mut edit = CellRunEdit {
                            block: block_idx,
                            cell: cell_idx,
                            para: para_idx,
                            run: run_idx,
                            ..Default::default()
                        };
                        let mut changed = false;
                        if br.text != er.text {
                            edit.new_text = Some(er.text.clone());
                            changed = true;
                        }
                        let (bprops, brs) =
                            effective_props(br.char_style_id.as_deref(), base, bindings);
                        let (eprops, ers) =
                            effective_props(er.char_style_id.as_deref(), edited, bindings);
                        if bprops != eprops || brs != ers {
                            edit.new_props = Some(eprops);
                            edit.rstyle = Some(ers);
                            changed = true;
                        }
                        if changed {
                            cells.push(edit);
                        }
                    }
                }
            }
            continue;
        }
        let (LoweredBlock::Paragraph(bp), LoweredBlock::Paragraph(ep)) = (bb, eb) else {
            continue; // a paragraph↔table swap — structural, deferred
        };
        if bp.runs.len() != ep.runs.len() {
            // Increment 2 — a run was inserted or removed. Align the two run
            // lists by their (text, style) identity and emit structural ops.
            structural.extend(align_runs(block_idx, bp, ep, edited, bindings));
            continue;
        }
        for (run_idx, (br, er)) in bp.runs.iter().zip(&ep.runs).enumerate() {
            let mut edit = RunEdit {
                block: block_idx,
                run: run_idx,
                ..Default::default()
            };
            let mut changed = false;

            if br.text != er.text {
                edit.new_text = Some(er.text.clone());
                changed = true;
            }

            let (bp_props, bp_rstyle) =
                effective_props(br.char_style_id.as_deref(), base, bindings);
            let (ep_props, ep_rstyle) =
                effective_props(er.char_style_id.as_deref(), edited, bindings);
            if bp_props != ep_props || bp_rstyle != ep_rstyle {
                edit.new_props = Some(ep_props);
                edit.rstyle = Some(ep_rstyle);
                changed = true;
            }

            if changed {
                runs.push(edit);
            }
        }

        // Increment 3 — the paragraph's own `<w:pPr>` (style + direct formatting).
        let (b_props, b_style) = effective_para_props(bp.para_style_id.as_deref(), base, bindings);
        let (e_props, e_style) =
            effective_para_props(ep.para_style_id.as_deref(), edited, bindings);
        if b_props != e_props || b_style != e_style {
            paragraphs.push(ParaEdit {
                block: block_idx,
                new_props: e_props,
                pstyle: Some(e_style),
            });
        }
    }
    EditSet {
        runs,
        structural,
        cells,
        paragraphs,
    }
}

/// The effective DIRECT paragraph formatting a lowered paragraph-style token maps
/// to, plus the real `<w:pStyle>` it should carry — the paragraph twin of
/// [`effective_props`].
fn effective_para_props(
    token: Option<&str>,
    doc: &LoweredDoc,
    bindings: &DocxBindings,
) -> (docx_core::ParaProps, Option<String>) {
    let Some(token) = token else {
        return (docx_core::ParaProps::default(), None);
    };
    if let Some(style_id) = bindings.para_token_to_style_id.get(token) {
        return (docx_core::ParaProps::default(), Some(style_id.clone()));
    }
    if let Some(style) = doc.styles.iter().find(|s| s.id == token) {
        let props = invert_para_props(&style.props);
        let pstyle = style
            .based_on
            .as_deref()
            .and_then(|b| bindings.para_token_to_style_id.get(b).cloned());
        return (props, pstyle);
    }
    (docx_core::ParaProps::default(), None)
}

/// Invert a synthesized paragraph style's `StyleProp`s back to `ParaProps` — the
/// reverse of `docx-lower`'s `para_props` (points → twips, 20 twips per point).
fn invert_para_props(props: &[StyleProp]) -> docx_core::ParaProps {
    use docx_core::Justification as J;
    let tw = |pt: f32| (pt * 20.0).round() as i32;
    let mut p = docx_core::ParaProps::default();
    for sp in props {
        match (sp.path.as_str(), &sp.value) {
            ("paragraphJustification", PropValue::Text(v)) => {
                p.justification = match v.as_str() {
                    "CenterAlign" => Some(J::Center),
                    "RightAlign" => Some(J::Right),
                    "LeftJustified" => Some(J::Both),
                    "FullyJustified" => Some(J::Distribute),
                    _ => Some(J::Left),
                };
            }
            ("paragraphLeftIndent", PropValue::Length(v)) => p.left_indent = Some(tw(*v)),
            ("paragraphRightIndent", PropValue::Length(v)) => p.right_indent = Some(tw(*v)),
            ("paragraphFirstLineIndent", PropValue::Length(v)) => {
                // A NEGATIVE first-line indent is Word's hanging indent.
                if *v < 0.0 {
                    p.hanging_indent = Some(tw(-*v));
                } else {
                    p.first_line_indent = Some(tw(*v));
                }
            }
            ("paragraphSpaceBefore", PropValue::Length(v)) => p.space_before = Some(tw(*v)),
            ("paragraphSpaceAfter", PropValue::Length(v)) => p.space_after = Some(tw(*v)),
            ("paragraphKeepWithNext", PropValue::Length(v)) => p.keep_next = Some(*v > 0.0),
            ("paragraphKeepLinesTogether", PropValue::Bool(b)) => p.keep_lines = Some(*b),
            ("paragraphTabStops", PropValue::TabStops(stops)) => {
                p.tabs = stops
                    .iter()
                    .map(|t| docx_core::TabStop {
                        position: tw(t.position),
                        alignment: t.alignment.clone(),
                        leader: None,
                    })
                    .collect();
            }
            _ => {}
        }
    }
    p
}

/// Align a paragraph's baseline runs against its edited runs by (text, style)
/// identity and emit the insert/delete ops that turn one into the other. A
/// classic LCS: matched runs are left alone, unmatched baseline runs are
/// deleted, unmatched edited runs are inserted after the preceding match (or at
/// the paragraph start).
fn align_runs(
    block: usize,
    bp: &docx_lower::ir::LoweredParagraph,
    ep: &docx_lower::ir::LoweredParagraph,
    edited: &LoweredDoc,
    bindings: &DocxBindings,
) -> Vec<StructuralEdit> {
    let key = |r: &docx_lower::ir::LoweredRun| (r.text.clone(), r.char_style_id.clone());
    let a: Vec<_> = bp.runs.iter().map(key).collect();
    let b: Vec<_> = ep.runs.iter().map(key).collect();

    // LCS table over the two run lists.
    let (n, m) = (a.len(), b.len());
    let mut lcs = vec![vec![0usize; m + 1]; n + 1];
    for i in (0..n).rev() {
        for j in (0..m).rev() {
            lcs[i][j] = if a[i] == b[j] {
                lcs[i + 1][j + 1] + 1
            } else {
                lcs[i + 1][j].max(lcs[i][j + 1])
            };
        }
    }

    let mut ops = Vec::new();
    let (mut i, mut j) = (0usize, 0usize);
    // The last baseline run index that survives — new runs anchor after it.
    let mut anchor: Option<usize> = None;
    while i < n && j < m {
        if a[i] == b[j] {
            anchor = Some(i);
            i += 1;
            j += 1;
        } else if lcs[i + 1][j] >= lcs[i][j + 1] {
            ops.push(StructuralEdit::DeleteRun { block, run: i });
            i += 1;
        } else {
            ops.push(insert_run_op(block, anchor, &ep.runs[j], edited, bindings));
            j += 1;
        }
    }
    while i < n {
        ops.push(StructuralEdit::DeleteRun { block, run: i });
        i += 1;
    }
    while j < m {
        ops.push(insert_run_op(block, anchor, &ep.runs[j], edited, bindings));
        j += 1;
    }
    ops
}

fn insert_run_op(
    block: usize,
    after: Option<usize>,
    run: &docx_lower::ir::LoweredRun,
    edited: &LoweredDoc,
    bindings: &DocxBindings,
) -> StructuralEdit {
    let (props, rstyle) = effective_props(run.char_style_id.as_deref(), edited, bindings);
    StructuralEdit::InsertRun {
        block,
        run: after,
        text: run.text.clone(),
        props,
        rstyle,
    }
}

/// The effective DIRECT character formatting a run's lowered style token maps to,
/// plus the real `w:rStyle` id it should carry. A real Word style ⇒ rStyle + no
/// direct props; a synthesized style ⇒ its inverted props + (its `basedOn`'s real
/// style, if any); no style ⇒ cleared formatting.
fn effective_props(
    token: Option<&str>,
    doc: &LoweredDoc,
    bindings: &DocxBindings,
) -> (RunProps, Option<String>) {
    let Some(token) = token else {
        return (RunProps::default(), None);
    };
    if let Some(style_id) = bindings.char_token_to_style_id.get(token) {
        return (RunProps::default(), Some(style_id.clone()));
    }
    if let Some(style) = doc.styles.iter().find(|s| s.id == token) {
        let props = invert_props(&style.props, doc);
        let rstyle = style
            .based_on
            .as_deref()
            .and_then(|b| bindings.char_token_to_style_id.get(b).cloned());
        return (props, rstyle);
    }
    (RunProps::default(), None)
}

/// Invert a synthesized character style's `StyleProp`s back to `RunProps` — the
/// exact reverse of `docx-lower`'s `run_props`.
fn invert_props(props: &[StyleProp], doc: &LoweredDoc) -> RunProps {
    let mut r = RunProps::default();
    for p in props {
        match (p.path.as_str(), &p.value) {
            ("characterFontFamily", PropValue::Text(f)) => r.font = Some(f.clone()),
            ("characterFontStyle", PropValue::Text(s)) => {
                let (b, i) = match s.as_str() {
                    "Bold Italic" => (true, true),
                    "Bold" => (true, false),
                    "Italic" => (false, true),
                    _ => (false, false), // "Regular"
                };
                r.bold = Some(b);
                r.italic = Some(i);
            }
            ("characterFontSize", PropValue::Length(pt)) => {
                r.size_half_pts = Some((pt * 2.0).round() as u32);
            }
            ("characterFillColor", PropValue::ColorRef(id)) => r.color = swatch_hex(id, doc),
            ("characterUnderline", PropValue::Bool(u)) => r.underline = Some(*u),
            ("characterStrikethru", PropValue::Bool(s)) => r.strike = Some(*s),
            ("characterPosition", PropValue::Text(pos)) => {
                r.vert_align = match pos.as_str() {
                    "Superscript" => Some(VertAlign::Superscript),
                    "Subscript" => Some(VertAlign::Subscript),
                    _ => None,
                };
            }
            ("characterCase", PropValue::Text(c)) => match c.as_str() {
                "SmallCaps" => r.small_caps = Some(true),
                "AllCaps" => r.caps = Some(true),
                _ => {}
            },
            ("characterBaselineShift", PropValue::Length(pt)) => {
                r.baseline_half_pts = Some((pt * 2.0).round() as i32);
            }
            _ => {}
        }
    }
    r
}

/// Resolve a swatch id to its `RRGGBB` hex (the inverse of `swatch_for`).
fn swatch_hex(id: &str, doc: &LoweredDoc) -> Option<String> {
    let sw = doc.swatches.iter().find(|s| s.id == id)?;
    if sw.value.len() < 3 {
        return None;
    }
    Some(format!(
        "{:02X}{:02X}{:02X}",
        sw.value[0] as u8, sw.value[1] as u8, sw.value[2] as u8
    ))
}
