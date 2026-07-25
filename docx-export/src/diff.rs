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
use crate::edit::{EditSet, RunEdit};

/// Diff two lowerings into the edits needed to turn `base` into `edited`.
pub fn diff(base: &LoweredDoc, edited: &LoweredDoc, bindings: &DocxBindings) -> EditSet {
    let mut runs = Vec::new();
    for (block_idx, (bb, eb)) in base
        .story
        .blocks
        .iter()
        .zip(&edited.story.blocks)
        .enumerate()
    {
        let (LoweredBlock::Paragraph(bp), LoweredBlock::Paragraph(ep)) = (bb, eb) else {
            continue; // table (or a paragraph↔table swap) — structural, deferred
        };
        if bp.runs.len() != ep.runs.len() {
            continue; // run inserted/removed — structural, deferred
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
    }
    EditSet { runs }
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
