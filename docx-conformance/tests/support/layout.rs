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

//! The engine's layout rule for the section fixtures (`continuous_docx()`,
//! `columns_docx()`), where every paragraph is one 12 pt line: enough of
//! core's flow to say where each line lands, so a test can hold the
//! lowering against Word's page map without the engine.
//!
//! - Every native story starts on a new page of its first section's
//!   geometry; its frame is that margin box in the story's frame columns
//!   (`SectionPlacement::frame`); the story grows pages with the same frame.
//!   A line fits only if its whole 12 pt box does (`LeadingOffset`).
//! - Ordinary paragraphs fill a column top-down, then the next column of
//!   the band, then the next page. `NextColumn` opens the next column (or
//!   page) unless the paragraph already opens one. A left indent shifts x.
//! - Span and split columns, as core lays them out (`span_columns.rs`,
//!   measured against InDesign 2025, ADR 028 addendum): text above a SPAN
//!   that starts at the top of its band and fits is balanced over the
//!   columns by line count (`ceil(lines / k)` each); the span sits at the
//!   full width below the deepest line above it; the text after it starts
//!   below it and fills the columns in turn. Consecutive SPLIT paragraphs
//!   with the same count, inside and outside gutter form one block in the
//!   current column (core `d4311c7`: a block ends where any of the three
//!   changes), `k` sub-columns `(column − 2 × outside − (k − 1) × inside)
//!   / k` wide, balanced by line count; the next paragraph or block starts
//!   below the deepest sub-column, a block boundary spaced by
//!   `max(SpaceAfter + SpaceBefore, ending block's SpanColumnMinSpaceAfter,
//!   starting block's SpanColumnMinSpaceBefore)` (0 in the fixtures, which
//!   this model asserts); a block that does not fit fills its sub-columns to
//!   the bottom and continues, balanced again, in the next column or page.
//!
//! A first line's top sits 2.03 pt below its 12 pt row in Word's PDF (the
//! glyph box, not the line box).
#![allow(dead_code)]

use std::collections::HashMap;

use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue};
use docx_lower::sections::SectionPlacement;

/// One laid-out line: its label, x and top (pt from the page's top-left),
/// as Word's maps record them.
#[derive(Debug, Clone, PartialEq)]
pub struct Line {
    pub label: String,
    pub x: f64,
    pub top: f64,
}

/// A page: its size and its lines, in the order Word's PDF lists them
/// (column by column).
pub type Page = (Vec<f64>, Vec<Line>);

/// A paragraph style property, resolved through the lowered `basedOn` chain
/// (nearest first), as the engine cascades it.
pub fn resolved(ir: &LoweredDoc, style: Option<&str>, path: &str) -> Option<PropValue> {
    let by_id: HashMap<&str, _> = ir.styles.iter().map(|s| (s.id.as_str(), s)).collect();
    let mut next = style;
    let mut depth = 0;
    while let Some(id) = next {
        let s = by_id.get(id)?;
        if let Some(p) = s.props.iter().rev().find(|p| p.path == path) {
            return Some(p.value.clone());
        }
        next = s.based_on.as_deref();
        depth += 1;
        assert!(depth < 32, "basedOn cycle");
    }
    None
}

pub fn length(ir: &LoweredDoc, style: Option<&str>, path: &str) -> f32 {
    match resolved(ir, style, path) {
        Some(PropValue::Length(v)) => v,
        _ => 0.0,
    }
}

pub fn text(ir: &LoweredDoc, style: Option<&str>, path: &str) -> Option<String> {
    match resolved(ir, style, path) {
        Some(PropValue::Text(t)) => Some(t),
        _ => None,
    }
}

pub fn label(block: &LoweredBlock) -> String {
    let LoweredBlock::Paragraph(p) = block else {
        panic!("the fixture has no tables")
    };
    let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
    text.split_whitespace().next().unwrap_or("").to_string()
}

/// How a paragraph sits, read from its resolved style.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Kind {
    Single,
    Span,
    Split { k: usize, inside: f64, outside: f64 },
}

pub fn kind(ir: &LoweredDoc, style: Option<&str>) -> Kind {
    match text(ir, style, "paragraphSpanColumnType").as_deref() {
        Some("SpanColumns") => {
            assert_eq!(
                text(ir, style, "paragraphSpanSplitColumnCount").as_deref(),
                Some("All")
            );
            Kind::Span
        }
        Some("SplitColumns") => Kind::Split {
            k: text(ir, style, "paragraphSpanSplitColumnCount")
                .expect("a split count")
                .parse()
                .expect("a whole number"),
            inside: f64::from(length(ir, style, "paragraphSplitColumnInsideGutter")),
            outside: f64::from(length(ir, style, "paragraphSplitColumnOutsideGutter")),
        },
        Some(other) if other != "SingleColumn" => panic!("{other}"),
        _ => Kind::Single,
    }
}

const PITCH: f64 = 12.0;
const GLYPH: f64 = 2.03;

/// The flow state inside one story.
struct Flow<'a> {
    pages: &'a mut Vec<Page>,
    size: Vec<f64>,
    left: f64,
    top: f64,
    rows: usize,
    cols: usize,
    col_w: f64,
    gutter: f64,
    /// The row the current band starts at (0 on a fresh page; below the
    /// last span otherwise).
    band_top: usize,
    /// The column being filled, and its next free row.
    col: usize,
    row: usize,
    /// One past the deepest row used in the band so far.
    deepest: Option<usize>,
}

impl Flow<'_> {
    fn new_page(&mut self) {
        self.pages.push((self.size.clone(), Vec::new()));
        (self.band_top, self.col, self.row, self.deepest) = (0, 0, 0, None);
    }

    fn next_column(&mut self) {
        if self.col + 1 < self.cols {
            self.col += 1;
            self.row = self.band_top;
        } else {
            self.new_page();
        }
    }

    fn at_band_start(&self) -> bool {
        self.col == 0 && self.row == self.band_top && self.deepest.is_none()
    }

    fn col_x(&self, c: usize) -> f64 {
        self.left + c as f64 * (self.col_w + self.gutter)
    }

    fn put(&mut self, label: String, x: f64, row: usize) {
        let top = self.top + row as f64 * PITCH + GLYPH;
        self.pages
            .last_mut()
            .unwrap()
            .1
            .push(Line { label, x, top });
        self.deepest = Some(self.deepest.map_or(row + 1, |d| d.max(row + 1)));
    }
}

/// Lay out `ir` (placed by `placed`, one entry per section) by the rule in
/// the module docs.
pub fn lay_out(ir: &LoweredDoc, placed: &[SectionPlacement]) -> Vec<Page> {
    assert_eq!(placed.len(), ir.sections.len());
    let mut pages: Vec<Page> = Vec::new();
    // (first section, end block) per story.
    let mut groups: Vec<(usize, usize)> = Vec::new();
    for (k, s) in ir.sections.iter().enumerate() {
        if k == 0 || ir.sections[k - 1].story != s.story {
            groups.push((k, 0));
        }
    }
    let starts: Vec<usize> = groups
        .iter()
        .map(|(k, _)| ir.sections[*k].first_block)
        .collect();
    for (i, g) in groups.iter_mut().enumerate() {
        g.1 = starts.get(i + 1).copied().unwrap_or(ir.story.blocks.len());
    }
    for (k, end) in groups {
        let sec = &ir.sections[k];
        let frame = placed[k].frame;
        let cols = frame.count as usize;
        let gutter = f64::from(frame.gutter_pt);
        let width = f64::from(sec.page_width_pt - sec.margin_left_pt - sec.margin_right_pt);
        let body = f64::from(sec.page_height_pt - sec.margin_top_pt - sec.margin_bottom_pt);
        let mut flow = Flow {
            pages: &mut pages,
            size: vec![f64::from(sec.page_width_pt), f64::from(sec.page_height_pt)],
            left: f64::from(sec.margin_left_pt),
            top: f64::from(sec.margin_top_pt),
            rows: (body / PITCH).floor() as usize,
            cols,
            col_w: (width - gutter * (cols as f64 - 1.0)) / cols as f64,
            gutter,
            band_top: 0,
            col: 0,
            row: 0,
            deepest: None,
        };
        flow.new_page();
        let blocks = &ir.story.blocks[sec.first_block..end];
        let style = |b: &LoweredBlock| match b {
            LoweredBlock::Paragraph(p) => p.para_style_id.clone(),
            LoweredBlock::Table(_) => panic!("the fixture has no tables"),
        };
        let kinds: Vec<Kind> = blocks
            .iter()
            .map(|b| {
                let k = kind(ir, style(b).as_deref());
                // A span in one column, or a split into one, is ordinary.
                match k {
                    Kind::Span if cols < 2 => Kind::Single,
                    Kind::Split { k: 1, .. } => Kind::Single,
                    k => k,
                }
            })
            .collect();
        let mut p = 0;
        while p < blocks.len() {
            let st = style(&blocks[p]);
            let st = st.as_deref();
            match kinds[p] {
                Kind::Single => {
                    if flow.at_band_start() {
                        let s = (p..blocks.len())
                            .find(|&s| kinds[s] != Kind::Single)
                            .unwrap_or(blocks.len());
                        if s < blocks.len() && kinds[s] == Kind::Span {
                            let per = (s - p).div_ceil(cols);
                            if flow.band_top + per <= flow.rows {
                                for (j, b) in blocks[p..s].iter().enumerate() {
                                    let x = flow.col_x(j / per)
                                        + f64::from(length(
                                            ir,
                                            style(b).as_deref(),
                                            "paragraphLeftIndent",
                                        ));
                                    flow.put(label(b), x, flow.band_top + j % per);
                                }
                                flow.col = cols; // the band is spent
                                p = s;
                                continue;
                            }
                        }
                    }
                    if text(ir, st, "paragraphStartParagraph").as_deref() == Some("NextColumn")
                        && flow.row > flow.band_top
                    {
                        flow.next_column();
                    }
                    if flow.row >= flow.rows {
                        flow.next_column();
                    }
                    let x = flow.col_x(flow.col) + f64::from(length(ir, st, "paragraphLeftIndent"));
                    let row = flow.row;
                    flow.put(label(&blocks[p]), x, row);
                    flow.row += 1;
                    p += 1;
                }
                Kind::Span => {
                    let mut at = flow.deepest.unwrap_or(flow.band_top);
                    if at >= flow.rows {
                        flow.new_page();
                        at = 0;
                    }
                    let x = flow.left + f64::from(length(ir, st, "paragraphLeftIndent"));
                    flow.put(label(&blocks[p]), x, at);
                    (flow.band_top, flow.col, flow.row, flow.deepest) = (at + 1, 0, at + 1, None);
                    p += 1;
                }
                Kind::Split { k, inside, outside } => {
                    let end = (p..blocks.len())
                        .find(|&e| kinds[e] != kinds[p])
                        .unwrap_or(blocks.len());
                    if p > 0 && matches!(kinds[p - 1], Kind::Split { .. }) {
                        // A block boundary: the engine's spacing rule, in
                        // whole rows (the fixtures space nothing).
                        let (prev, this) = (style(&blocks[p - 1]), style(&blocks[p]));
                        let gap = (length(ir, prev.as_deref(), "paragraphSpaceAfter")
                            + length(ir, this.as_deref(), "paragraphSpaceBefore"))
                        .max(length(
                            ir,
                            prev.as_deref(),
                            "paragraphSpanColumnMinSpaceAfter",
                        ))
                        .max(length(
                            ir,
                            this.as_deref(),
                            "paragraphSpanColumnMinSpaceBefore",
                        ));
                        assert_eq!(gap, 0.0, "the model spaces split blocks by whole rows only");
                    }
                    let sub_w = (flow.col_w - 2.0 * outside - (k as f64 - 1.0) * inside) / k as f64;
                    let mut from = p;
                    while from < end {
                        if flow.row >= flow.rows {
                            flow.next_column();
                        }
                        let top = flow.row;
                        let room = flow.rows - top;
                        let left = end - from;
                        let per = left.div_ceil(k);
                        let per = if per <= room { per } else { room };
                        let take = (per * k).min(left);
                        let x0 = flow.col_x(flow.col) + outside;
                        for (j, b) in blocks[from..from + take].iter().enumerate() {
                            let x = x0 + (j / per) as f64 * (sub_w + inside);
                            flow.put(label(b), x, top + j % per);
                        }
                        from += take;
                        flow.row = if from < end {
                            flow.rows // full: continue in the next column
                        } else {
                            top + per
                        };
                    }
                    p = end;
                }
            }
        }
    }
    pages
}

/// Word's page map from a fixture's JSON.
pub fn word_pages(map: &serde_json::Value) -> Vec<Page> {
    map["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|p| {
            let size = p["size_pt"]
                .as_array()
                .unwrap()
                .iter()
                .map(|v| v.as_f64().unwrap())
                .collect();
            let lines = p["lines"]
                .as_array()
                .unwrap()
                .iter()
                .map(|l| Line {
                    label: l["label"].as_str().unwrap().to_string(),
                    x: l["x"].as_f64().unwrap(),
                    top: l["top"].as_f64().unwrap(),
                })
                .collect();
            (size, lines)
        })
        .collect()
}

/// A page's labels, in reading order.
pub fn labels(p: &Page) -> Vec<String> {
    p.1.iter().map(|l| l.label.clone()).collect()
}

/// Every line of `ours` where Word has it: label, x within 0.2 pt and top
/// within 0.25 pt. Word's PDF lists a page's text column by column; ours is
/// sorted the same way (by x, then top) before comparing. Word's glyph tops
/// drift up to 0.2 pt from the exact 12 pt grid over a 20-line page (PDF
/// rounding), and its later columns start up to 0.17 pt right of the
/// computed edge, hence the tolerances.
pub fn assert_page(ours: &Page, word: &Page, what: &str) {
    assert_eq!(ours.0, word.0, "{what}: page size");
    let sorted = |p: &Page| -> Vec<Line> {
        let mut v = p.1.clone();
        v.sort_by(|a, b| {
            let col = |l: &Line| (l.x / 10.0).round() as i64;
            col(a).cmp(&col(b)).then(a.top.partial_cmp(&b.top).unwrap())
        });
        v
    };
    let (a, b) = (sorted(ours), sorted(word));
    let names = |v: &[Line]| v.iter().map(|l| l.label.clone()).collect::<Vec<_>>();
    assert_eq!(names(&a), names(&b), "{what}: lines");
    for (a, b) in a.iter().zip(&b) {
        assert!(
            (a.x - b.x).abs() < 0.2 && (a.top - b.top).abs() < 0.25,
            "{what}: {} at ({}, {}), Word ({}, {})",
            a.label,
            a.x,
            a.top,
            b.x,
            b.top
        );
    }
}
