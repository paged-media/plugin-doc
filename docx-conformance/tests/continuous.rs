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

#[path = "support/layout.rs"]
mod layout;

use docx_conformance::{continuous_docx, CONTINUOUS_CASES};
use docx_core::SectionKind;
use docx_import::import_docx;
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue};
use layout::{assert_page, label, labels, lay_out, length, resolved, word_pages};

fn word_map() -> serde_json::Value {
    serde_json::from_str(include_str!("../fixtures/continuous.word.json"))
        .expect("continuous.word.json")
}

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&continuous_docx()).expect("import"))
}

fn placements(options: docx_lower::LowerOptions) -> Vec<docx_lower::sections::SectionPlacement> {
    let doc = import_docx(&continuous_docx()).expect("import");
    docx_lower::sections::place_sections_with(&doc.sections, options, &mut Vec::new())
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
            // (b) the column count changes mid-page: one story in two
            //     columns, its one-column sections spanning them.
            ("B1", 1),
            ("B2", 1),
            ("B3", 1),
            ("B4", 1),
            // (d) nextColumn: joins.
            ("D1", 2),
            ("D2", 2),
            // (c1) other left/right margins: joins (as indents).
            ("C1", 3),
            ("C2", 3),
            ("C3", 3),
            // (c2) other top/bottom margins: joins (later pages differ).
            ("E1", 4),
            ("E2", 4),
            // (c3) another page size: Word itself starts a new page.
            ("F1", 5),
            ("F2", 6),
            ("G1", 7),
        ]
    );
    let placed = placements(docx_lower::LowerOptions::default());
    let frame = |k: usize| (placed[k].frame.count, placed[k].frame.gutter_pt);
    assert_eq!(
        frame(3),
        (2, 36.0),
        "B's story: Word's two columns, 0.5 in apart"
    );
    use docx_lower::sections::SectionColumns::{Frame, SpanAll};
    assert_eq!(
        placed[3..7].iter().map(|p| p.columns).collect::<Vec<_>>(),
        vec![SpanAll, Frame, SpanAll, Frame]
    );

    // Every difference from Word is diagnosed, nothing else: no warning,
    // the span/split lowering said once, the top/bottom margins once.
    let warnings: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.severity == "warning")
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warnings, Vec::<&str>::new());
    let columns: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.message.contains("mid-page"))
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(columns.len(), 1, "{columns:?}");
    assert!(columns[0].starts_with("sections 4–7 "), "{}", columns[0]);
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

/// An engine that refuses span/split columns (before protocol 64) gets
/// the page-break lowering: each column change opens a new page, with a
/// warning, and no paragraph carries a span/split property.
#[test]
fn without_span_columns_a_column_change_opens_a_page__feat__plugin_doc_word_pagination() {
    let options = docx_lower::LowerOptions {
        mid_page_columns: false,
    };
    let ir = docx_lower::lower_with(&import_docx(&continuous_docx()).expect("import"), options);
    let b: Vec<(&str, usize)> = stories(&ir)[3..7].to_vec();
    assert_eq!(b, vec![("B1", 1), ("B2", 2), ("B3", 3), ("B4", 4)]);
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
            w.starts_with(n) && w.contains("protocol 64") && w.contains("new page"),
            "{w}"
        );
    }
    assert!(!ir
        .styles
        .iter()
        .flat_map(|s| &s.props)
        .any(|p| p.path.contains("SpanColumn") || p.path.contains("SplitColumn")));
    // The page map is today's: B's four sections on four pages.
    let ours = lay_out(&ir, &placements(options));
    assert_eq!(ours.len(), word_pages(&word_map()).len() + 3);
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

/// The page map: every Word page is reproduced line for line (label, x,
/// top) except the one documented difference, which is pinned as it is.
#[test]
fn the_lowering_reproduces_words_page_map__feat__plugin_doc_word_pagination() {
    let ir = lowered();
    let ours = lay_out(&ir, &placements(docx_lower::LowerOptions::default()));
    let word = word_pages(&word_map());
    assert_eq!(word.len(), 10);
    assert_eq!(ours.len(), word.len(), "Word's page count");

    // (a) Word pages 1–2: the invisible boundary, A3 running on.
    assert_page(&ours[0], &word[0], "Word page 1");
    assert_page(&ours[1], &word[1], "Word page 2");
    // (b) Word page 3: B1, then B2 in two columns BALANCED 5 / 4 (a
    // continuous break follows), B3 below the deeper column, and B4 in two
    // columns NOT balanced (a nextPage section follows): all six lines in
    // column 1. The story's frame has two columns; B1 and B3 span them; the
    // engine balances the text above a span and fills the columns after the
    // last one in turn — line for line Word's.
    assert_page(&ours[2], &word[2], "Word page 3");
    // (d) nextColumn: D2 opens column 2 of the same page.
    assert_page(&ours[3], &word[3], "Word page 4");
    // (c1) other left/right margins, applied mid-page and undone.
    assert_page(&ours[4], &word[4], "Word page 5");
    // (c2) other top/bottom margins: the section continues on Word's page
    // under the old margins…
    assert_page(&ours[5], &word[5], "Word page 6");
    // …but Word's next page uses the NEW 1 in top margin; the native story's
    // grown page keeps the first section's 0.5 in (the diagnosed loss). The
    // same lines land on it, 36 pt higher.
    let (o, w) = (&ours[6], &word[6]);
    assert_eq!(labels(o), labels(w));
    for (a, b) in o.1.iter().zip(&w.1) {
        assert!((b.top - a.top - 36.0).abs() < 0.25, "{}", a.label);
    }
    // (c3) another page size: Word's own new page, reproduced.
    assert_page(&ours[7], &word[7], "Word page 8");
    assert_page(&ours[8], &word[8], "Word page 9");
    assert_page(&ours[9], &word[9], "Word page 10");
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
