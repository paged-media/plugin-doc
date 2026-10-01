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

//! thoughts ADR 029 — `continuous` (and `nextColumn`) sections as Word lays
//! them out. `continuous_docx()` is the document Word paginated;
//! `fixtures/continuous.word.json` is its answer (`scripts/word-continuous-probe.sh`).
//!
//! What is pinned here:
//! - the import carries every section's kind, page, margins and columns;
//! - the lowering puts each section in the story Word's layout calls for
//!   (`docx_lower::sections`): joined when Word continues it on the same
//!   page and the native model can say so, a new story otherwise, with a
//!   diagnostic for every place it differs from Word;
//! - a joined section with other left/right margins carries the difference
//!   as paragraph indents, a `nextColumn` section a `NextColumn` rule;
//! - laid out by the engine's rule on the skeleton's frames (a story per
//!   page-starting section, its margin box, its columns, 12 pt lines), the
//!   lowering reproduces Word's page map — label, x and top of every line —
//!   except where it is documented not to (mid-page column changes; a later
//!   page's top margin), and those differences are pinned too.
//!
//! `__feat__` names link these tests to the Cockpit feature row.
#![allow(non_snake_case)]

use std::collections::HashMap;

use docx_conformance::{continuous_docx, CONTINUOUS_CASES};
use docx_core::SectionKind;
use docx_import::import_docx;
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue};

fn word_map() -> serde_json::Value {
    serde_json::from_str(include_str!("../fixtures/continuous.word.json"))
        .expect("continuous.word.json")
}

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&continuous_docx()).expect("import"))
}

/// A paragraph style property, resolved through the lowered `basedOn` chain
/// (nearest first), as the engine cascades it.
fn resolved(ir: &LoweredDoc, style: Option<&str>, path: &str) -> Option<PropValue> {
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

fn length(ir: &LoweredDoc, style: Option<&str>, path: &str) -> f32 {
    match resolved(ir, style, path) {
        Some(PropValue::Length(v)) => v,
        _ => 0.0,
    }
}

fn label(block: &LoweredBlock) -> String {
    let LoweredBlock::Paragraph(p) = block else {
        panic!("the fixture has no tables")
    };
    let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
    text.split_whitespace().next().unwrap_or("").to_string()
}

/// The story index of every section, by its label.
fn stories(ir: &LoweredDoc) -> Vec<(&'static str, usize)> {
    CONTINUOUS_CASES
        .iter()
        .zip(&ir.sections)
        .map(|(c, s)| (c.label, s.story))
        .collect()
}

#[test]
fn every_section_imports_with_its_kind_and_geometry__feat__plugin_doc_word_pagination() {
    let doc = import_docx(&continuous_docx()).expect("import");
    assert_eq!(doc.sections.len(), CONTINUOUS_CASES.len());
    let mut first = 0;
    for (sec, case) in doc.sections.iter().zip(CONTINUOUS_CASES) {
        let kind = match case.kind {
            "nextPage" => SectionKind::NextPage,
            "continuous" => SectionKind::Continuous,
            "nextColumn" => SectionKind::NextColumn,
            other => panic!("{other}"),
        };
        assert_eq!(sec.kind, kind, "{}", case.label);
        assert_eq!(
            (sec.page_width, sec.page_height),
            case.page,
            "{}",
            case.label
        );
        let (t, r, b, l) = case.margins;
        assert_eq!(
            (
                sec.margin_top,
                sec.margin_right,
                sec.margin_bottom,
                sec.margin_left
            ),
            (t, r, b, l),
            "{}",
            case.label
        );
        assert_eq!(sec.columns, case.columns, "{}", case.label);
        assert_eq!(sec.first_block, first, "{}", case.label);
        first += case.lines as usize;
    }
}

/// Which sections share a story, and what is said about the rest.
#[test]
fn continuous_sections_join_the_story_word_continues__feat__plugin_doc_word_pagination() {
    let ir = lowered();
    assert_eq!(
        stories(&ir),
        vec![
            // (a) same geometry: one story.
            ("A1", 0),
            ("A2", 0),
            ("A3", 0),
            // (b) the column count changes: each its own story (a new page).
            ("B1", 1),
            ("B2", 2),
            ("B3", 3),
            ("B4", 4),
            // (d) nextColumn: joins.
            ("D1", 5),
            ("D2", 5),
            // (c1) other left/right margins: joins (as indents).
            ("C1", 6),
            ("C2", 6),
            ("C3", 6),
            // (c2) other top/bottom margins: joins (later pages differ).
            ("E1", 7),
            ("E2", 7),
            // (c3) another page size: Word itself starts a new page.
            ("F1", 8),
            ("F2", 9),
            ("G1", 10),
        ]
    );

    // Every difference from Word is diagnosed, nothing else.
    let warnings: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.severity == "warning")
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warnings.len(), 3, "{warnings:?}");
    for (w, n) in warnings
        .iter()
        .zip(["section 5 ", "section 6 ", "section 7 "])
    {
        assert!(
            w.starts_with(n) && w.contains("column") && w.contains("new page"),
            "{w}"
        );
    }
    let margins: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.message.contains("top/bottom margins"))
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(margins.len(), 1, "{margins:?}");
    assert!(margins[0].starts_with("section 14 "), "{}", margins[0]);
    assert!(
        !ir.diagnostics
            .iter()
            .any(|d| d.message.contains("section 16 ")),
        "another page size is Word's own new page: no loss to report"
    );
}

/// The joined sections' paragraphs: `nextColumn` carries the rule, other
/// left/right margins carry indents equal to Word's measured shift.
#[test]
fn joined_sections_carry_word_s_column_start_and_margins__feat__plugin_doc_word_pagination() {
    let ir = lowered();
    let word = word_map();
    let x_of = |want: &str| -> f64 {
        word["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|p| p["lines"].as_array().unwrap())
            .find(|l| l["label"] == want)
            .unwrap_or_else(|| panic!("{want} in Word's map"))["x"]
            .as_f64()
            .unwrap()
    };
    let base_x = x_of("C1-01");
    for block in &ir.story.blocks {
        let LoweredBlock::Paragraph(p) = block else {
            unreachable!()
        };
        let l = label(block);
        let style = p.para_style_id.as_deref();
        let rule = match resolved(&ir, style, "paragraphStartParagraph") {
            Some(PropValue::Text(t)) if t != "Anywhere" => Some(t),
            _ => None,
        };
        assert_eq!(
            rule.as_deref(),
            (l == "D2-01").then_some("NextColumn"),
            "{l}"
        );
        let (left, right) = (
            length(&ir, style, "paragraphLeftIndent"),
            length(&ir, style, "paragraphRightIndent"),
        );
        if l.starts_with("C2-") {
            // Word: x 108.07 against 36.02 — the 1.5 in left margin, mid-page.
            assert!(
                (f64::from(left) - (x_of(&l) - base_x)).abs() < 0.05,
                "{l}: left indent {left} vs Word's shift"
            );
            assert_eq!(right, 36.0, "{l}: 1 in right margin against 0.5 in");
        } else {
            assert_eq!((left, right), (0.0, 0.0), "{l}");
        }
        // The exact 12 pt grid holds for the indented paragraphs too.
        assert_eq!(length(&ir, style, "characterLeading"), 12.0, "{l}");
    }
}

/// One laid-out line: its page, label, x and top (pt from the page's
/// top-left), as Word's map records them.
#[derive(Debug, Clone, PartialEq)]
struct Line {
    label: String,
    x: f64,
    top: f64,
}

/// Lay the lowering out with the engine's rule on the skeleton's frames:
/// every story starts on a new page of its first section's geometry, its
/// frame is that margin box in that many columns (36 pt gap), lines are
/// 12 pt and fill a column top-down (a line fits only if its whole 12 pt box
/// does), the story grows pages with the same frame, `NextColumn` opens the
/// next column (or page) unless the paragraph already opens one, and a
/// paragraph's left indent shifts its x. A first line's top sits 2.03 pt
/// below the frame's top in Word's PDF (the glyph box, not the line box).
fn lay_out(ir: &LoweredDoc) -> Vec<(Vec<f64>, Vec<Line>)> {
    const PITCH: f64 = 12.0;
    const GLYPH: f64 = 2.03;
    let mut pages: Vec<(Vec<f64>, Vec<Line>)> = Vec::new();
    let groups = {
        let mut g: Vec<(usize, usize)> = Vec::new(); // (section, end block)
        for (k, s) in ir.sections.iter().enumerate() {
            if k == 0 || ir.sections[k - 1].story != s.story {
                g.push((k, 0));
            }
        }
        let starts: Vec<usize> = g.iter().map(|(k, _)| ir.sections[*k].first_block).collect();
        for (i, e) in g.iter_mut().enumerate() {
            e.1 = starts.get(i + 1).copied().unwrap_or(ir.story.blocks.len());
        }
        g
    };
    for (k, end) in groups {
        let sec = &ir.sections[k];
        let cols = sec.columns as usize;
        let width = f64::from(sec.page_width_pt - sec.margin_left_pt - sec.margin_right_pt);
        let col_w = (width - 36.0 * (cols as f64 - 1.0)) / cols as f64;
        let body = f64::from(sec.page_height_pt - sec.margin_top_pt - sec.margin_bottom_pt);
        let per_col = (body / PITCH).floor() as usize;
        let size = vec![f64::from(sec.page_width_pt), f64::from(sec.page_height_pt)];
        pages.push((size.clone(), Vec::new()));
        let (mut col, mut line) = (0usize, 0usize);
        for block in &ir.story.blocks[sec.first_block..end] {
            let LoweredBlock::Paragraph(p) = block else {
                unreachable!()
            };
            let style = p.para_style_id.as_deref();
            let next = |col: &mut usize, line: &mut usize, pages: &mut Vec<_>| {
                if *col + 1 < cols {
                    (*col, *line) = (*col + 1, 0);
                } else {
                    pages.push((size.clone(), Vec::new()));
                    (*col, *line) = (0, 0);
                }
            };
            if let Some(PropValue::Text(t)) = resolved(ir, style, "paragraphStartParagraph") {
                if t == "NextColumn" && line > 0 {
                    next(&mut col, &mut line, &mut pages);
                }
            }
            if line == per_col {
                next(&mut col, &mut line, &mut pages);
            }
            let x = f64::from(sec.margin_left_pt)
                + col as f64 * (col_w + 36.0)
                + f64::from(length(ir, style, "paragraphLeftIndent"));
            let top = f64::from(sec.margin_top_pt) + line as f64 * PITCH + GLYPH;
            pages.last_mut().unwrap().1.push(Line {
                label: label(block),
                x,
                top,
            });
            line += 1;
        }
    }
    pages
}

fn word_pages() -> Vec<(Vec<f64>, Vec<Line>)> {
    word_map()["pages"]
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

fn assert_page(ours: &(Vec<f64>, Vec<Line>), word: &(Vec<f64>, Vec<Line>), what: &str) {
    assert_eq!(ours.0, word.0, "{what}: page size");
    // Word's PDF lists a page's text column by column; ours fills in the
    // same order. Word's glyph tops drift up to 0.2 pt from the exact 12 pt
    // grid over a 20-line page (PDF rounding), and its second column starts
    // 0.13 pt right of 36 + 126 + 36 (both columns are 126 pt either way),
    // hence the tolerances.
    let labels = |p: &(Vec<f64>, Vec<Line>)| -> Vec<String> {
        p.1.iter().map(|l| l.label.clone()).collect()
    };
    assert_eq!(labels(ours), labels(word), "{what}: lines");
    for (a, b) in ours.1.iter().zip(&word.1) {
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

/// The page map: every Word page is reproduced line for line (label, x,
/// top) except the two documented differences, which are pinned as they are.
#[test]
fn the_lowering_reproduces_words_page_map__feat__plugin_doc_word_pagination() {
    let ir = lowered();
    let ours = lay_out(&ir);
    let word = word_pages();
    assert_eq!(word.len(), 10);

    // (a) Word pages 1–2: the invisible boundary, A3 running on.
    assert_page(&ours[0], &word[0], "Word page 1");
    assert_page(&ours[1], &word[1], "Word page 2");

    // (b) Word page 3 holds B1, the balanced two-column B2, B3 and B4 in
    // one page. The engine cannot change the column count mid-page: each
    // section opens a page, and B2 (one frame, two columns, nothing to
    // balance against) fills its first column. Pinned as the known loss.
    let labels = |p: &(Vec<f64>, Vec<Line>)| -> Vec<String> {
        p.1.iter().map(|l| l.label.clone()).collect()
    };
    let b: Vec<Vec<String>> = ours[2..6].iter().map(labels).collect();
    let section =
        |s: &str, n: u32| -> Vec<String> { (1..=n).map(|i| format!("{s}-{i:02}")).collect() };
    assert_eq!(
        b,
        vec![
            section("B1", 3),
            section("B2", 9),
            section("B3", 3),
            section("B4", 6)
        ]
    );
    assert!(ours[3].1.iter().all(|l| (l.x - 36.0).abs() < 0.1));
    let mut word_b = labels(&word[2]);
    word_b.sort();
    assert_eq!(
        word_b,
        b.concat().into_iter().collect::<Vec<_>>(),
        "same lines"
    );

    // From here on, ours is three pages ahead.
    let shift = 3;
    // (d) nextColumn: D2 opens column 2 of the same page.
    assert_page(&ours[3 + shift], &word[3], "Word page 4");
    // (c1) other left/right margins, applied mid-page and undone.
    assert_page(&ours[4 + shift], &word[4], "Word page 5");
    // (c2) other top/bottom margins: the section continues on Word's page
    // under the old margins…
    assert_page(&ours[5 + shift], &word[5], "Word page 6");
    // …but Word's next page uses the NEW 1 in top margin; the native story's
    // grown page keeps the first section's 0.5 in (the diagnosed loss). The
    // same lines land on it, 36 pt higher.
    let (o, w) = (&ours[6 + shift], &word[6]);
    assert_eq!(labels(o), labels(w));
    for (a, b) in o.1.iter().zip(&w.1) {
        assert!((b.top - a.top - 36.0).abs() < 0.25, "{}", a.label);
    }
    // (c3) another page size: Word's own new page, reproduced.
    assert_page(&ours[7 + shift], &word[7], "Word page 8");
    assert_page(&ours[8 + shift], &word[8], "Word page 9");
    assert_page(&ours[9 + shift], &word[9], "Word page 10");
    assert_eq!(ours.len(), word.len() + shift);
}

/// A page break ending the last paragraph of a section that a continuous
/// section joins carries into the joined section's first paragraph (same
/// story), where a break before a NEW story is dropped as redundant.
#[test]
fn a_break_ending_a_section_carries_into_the_joined_section__feat__plugin_doc_word_pagination() {
    use docx_core::{Block, BreakKind, RunBreak};
    let mut doc = import_docx(&continuous_docx()).expect("import");
    // A1-04 is body block 3, the last of A1; A2 (continuous) joins.
    let Block::Paragraph(p) = &mut doc.body[3] else {
        unreachable!()
    };
    let run = p.runs.last_mut().unwrap();
    run.breaks.push(RunBreak {
        at: run.text.chars().count(),
        kind: BreakKind::Page,
    });
    let ir = docx_lower::lower(&doc);
    let LoweredBlock::Paragraph(a2) = &ir.story.blocks[4] else {
        unreachable!()
    };
    assert_eq!(label(&ir.story.blocks[4]), "A2-01");
    assert_eq!(
        resolved(&ir, a2.para_style_id.as_deref(), "paragraphStartParagraph"),
        Some(PropValue::Text("NextPage".into()))
    );
    assert!(
        !ir.diagnostics
            .iter()
            .any(|d| d.message.contains("break ends the section")),
        "nothing dropped"
    );
}
