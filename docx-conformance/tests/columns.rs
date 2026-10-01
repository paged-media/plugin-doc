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

//! thoughts ADR 029 — continuous sections that CHANGE THE COLUMNS mid-page,
//! as Word lays them out. `columns_docx()` is the document Word paginated;
//! `fixtures/columns.word.json` is its answer (`scripts/word-columns-probe.sh`).
//!
//! What Word does (the map, page by page):
//! - a section before a `continuous` break is balanced by line count
//!   (6 in 2 → 3 / 3, 7 in 3 → 3 / 3 / 1, 5 in 2 → 3 / 2), also where one
//!   column count changes straight into another (H, N, P) and where two
//!   sections in a row have the SAME columns: each is balanced on its own
//!   (Q: 3 / 2 then 3 / 2; R: 3 / 2, then the last section unbalanced);
//! - a page-starting multi-column section is balanced the same way when a
//!   continuous section follows (I1: 4 / 3);
//! - a section before a `nextPage` section, or at the document's end, is
//!   NOT balanced: its lines fill column 1 first (I3, M2, O2, R3);
//! - a section running past the page fills that page's columns to the
//!   bottom (17 / 17) and is balanced on its last page only if a continuous
//!   break follows (L2: 8 / 8; M2: 6 in column 1);
//! - the gap is `w:space` (J2: 18 pt); unequal columns keep their widths
//!   (K2: 180 pt and 90 pt, balanced 4 / 2).
//!
//! What the lowering reproduces, line for line under the engine's span /
//! split rule (`support/layout.rs`): H, I, J, L, M, N, O and P — a column
//! change straight into another count or gap is a new split block (core
//! `d4311c7`), balanced on its own, directly below the last. Pinned
//! differences, on Word's page with Word's lines: two sections in a row
//! with the same columns share one column flow (the engine has no block
//! boundary where nothing changes: Q is one block 5 / 5, R fills column 1);
//! unequal columns are equal (K).
#![allow(non_snake_case)]

#[path = "support/layout.rs"]
mod layout;

use docx_conformance::{columns_docx, COLUMNS_CASES};
use docx_import::import_docx;
use docx_lower::ir::LoweredDoc;
use docx_lower::sections::{SectionColumns, SectionPlacement};
use docx_lower::LowerOptions;
use layout::{assert_page, labels, lay_out, word_pages, Page};

fn word() -> Vec<Page> {
    word_pages(
        &serde_json::from_str(include_str!("../fixtures/columns.word.json"))
            .expect("columns.word.json"),
    )
}

fn lowered(options: LowerOptions) -> (LoweredDoc, Vec<SectionPlacement>) {
    let doc = import_docx(&columns_docx()).expect("import");
    let placed = docx_lower::sections::place_sections_with(&doc.sections, options, &mut Vec::new());
    (docx_lower::lower_with(&doc, options), placed)
}

#[test]
fn every_section_imports_its_columns__feat__plugin_doc_word_pagination() {
    let doc = import_docx(&columns_docx()).expect("import");
    assert_eq!(doc.sections.len(), COLUMNS_CASES.len());
    for (sec, case) in doc.sections.iter().zip(COLUMNS_CASES) {
        assert_eq!(sec.columns, case.columns, "{}", case.label);
        assert_eq!(sec.column_space, case.space, "{}", case.label);
        assert_eq!(sec.column_widths, case.widths.to_vec(), "{}", case.label);
    }
}

/// Which story each section pours into, and how it sits in its columns.
#[test]
fn column_changes_lower_to_span_and_split_columns__feat__plugin_doc_word_pagination() {
    use SectionColumns::{Frame, SpanAll};
    let split = |count, inside_pt| SectionColumns::Split { count, inside_pt };
    let (ir, placed) = lowered(LowerOptions::default());
    let got: Vec<(&str, usize, u32, SectionColumns)> = COLUMNS_CASES
        .iter()
        .zip(&placed)
        .map(|(c, p)| (c.label, p.story, p.frame.count, p.columns))
        .collect();
    assert_eq!(
        got,
        vec![
            // (h) one story: H2 and H3 are two split blocks (the count
            //     changes), each balanced on its own.
            ("H1", 0, 1, Frame),
            ("H2", 0, 1, split(2, 36.0)),
            ("H3", 0, 1, split(3, 36.0)),
            ("H4", 0, 1, Frame),
            // (i) the last section is left unbalanced: the frame keeps its
            //     columns (filled in turn after the last span), the
            //     one-column section spans them.
            ("I1", 1, 2, Frame),
            ("I2", 1, 2, SpanAll),
            ("I3", 1, 2, Frame),
            // (j) split, at Word's gap.
            ("J1", 2, 1, Frame),
            ("J2", 2, 1, split(2, 18.0)),
            ("J3", 2, 1, Frame),
            // (k) unequal columns, as equal ones at the first one's gap.
            ("K1", 3, 1, Frame),
            ("K2", 3, 1, split(2, 18.0)),
            ("K3", 3, 1, Frame),
            // (l) split: the block fills the page, then balances the rest.
            ("L1", 4, 1, Frame),
            ("L2", 4, 1, split(2, 36.0)),
            ("L3", 4, 1, Frame),
            // (m) unbalanced last section: the frame's columns.
            ("M1", 5, 2, SpanAll),
            ("M2", 5, 2, Frame),
            // (n) 3 → 2: two blocks.
            ("N1", 6, 1, Frame),
            ("N2", 6, 1, split(3, 36.0)),
            ("N3", 6, 1, split(2, 36.0)),
            ("N4", 6, 1, Frame),
            // (p) another gap: two blocks.
            ("P1", 7, 1, Frame),
            ("P2", 7, 1, split(2, 36.0)),
            ("P3", 7, 1, split(2, 18.0)),
            ("P4", 7, 1, Frame),
            // (q) the same columns twice: one block to the engine.
            ("Q1", 8, 1, Frame),
            ("Q2", 8, 1, split(2, 36.0)),
            ("Q3", 8, 1, split(2, 36.0)),
            ("Q4", 8, 1, Frame),
            // (r) the same, unbalanced at the end: the frame's columns.
            ("R1", 9, 2, SpanAll),
            ("R2", 9, 2, Frame),
            ("R3", 9, 2, Frame),
            // (o) the document ends unbalanced: the frame's columns.
            ("O1", 10, 2, SpanAll),
            ("O2", 10, 2, Frame),
        ]
    );
    assert_eq!(placed[4].frame.gutter_pt, 36.0, "I's 0.5 in gap");

    let warnings: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.severity == "warning")
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warnings.len(), 3, "{warnings:?}");
    assert!(
        warnings[0].starts_with("section 12 ") && warnings[0].contains("unequal"),
        "{}",
        warnings[0]
    );
    for (w, n) in warnings[1..].iter().zip(["section 29 ", "section 33 "]) {
        assert!(
            w.starts_with(n) && w.contains("same 2 columns") && w.contains("one column flow"),
            "{w}"
        );
    }
    assert!(
        !warnings.iter().any(|w| w.contains("new page")),
        "no column change opens a page: {warnings:?}"
    );
}

/// The page map, line by line (label, x within 0.2 pt, top within
/// 0.25 pt), with the pinned differences.
#[test]
fn the_lowering_reproduces_words_column_changes__feat__plugin_doc_word_pagination() {
    let (ir, placed) = lowered(LowerOptions::default());
    let ours = lay_out(&ir, &placed);
    let word = word();
    assert_eq!(word.len(), 13);
    // Every section on Word's page: no column change costs a page.
    assert_eq!(ours.len(), word.len());

    // (h) 1 → 2 → 3 → 1, (i), (j): Word's pages exactly.
    assert_page(&ours[0], &word[0], "Word page 1 (H)");
    assert_page(&ours[1], &word[1], "Word page 2 (I)");
    assert_page(&ours[2], &word[2], "Word page 3 (J)");

    // (k) unequal columns: same lines on the same page; K2's second column
    // sits at the equal-column x (189 pt, Word 234.17) and is balanced 3 / 3
    // (Word 4 / 2).
    let (o, w) = (&ours[3], &word[3]);
    let mut a = labels(o);
    let mut b = labels(w);
    a.sort();
    b.sort();
    assert_eq!(a, b, "K: the same lines");
    let col2: Vec<&str> =
        o.1.iter()
            .filter(|l| l.x > 150.0)
            .map(|l| l.label.as_str())
            .collect();
    assert_eq!(col2, ["K2-04", "K2-05", "K2-06"]);
    assert!(o
        .1
        .iter()
        .filter(|l| l.x > 150.0)
        .all(|l| (l.x - 189.0).abs() < 0.01));

    // (l) past the page: 17 / 17, then balanced 8 / 8 under L3's span.
    assert_page(&ours[4], &word[4], "Word page 5 (L)");
    assert_page(&ours[5], &word[5], "Word page 6 (L)");
    // (m) past the page before a nextPage section: not balanced.
    assert_page(&ours[6], &word[6], "Word page 7 (M)");
    assert_page(&ours[7], &word[7], "Word page 8 (M)");
    // (n) 3 → 2 and (p) another gap: two blocks, each balanced.
    assert_page(&ours[8], &word[8], "Word page 9 (N)");
    assert_page(&ours[9], &word[9], "Word page 10 (P)");

    // (q), (r) the same columns twice: Word's page and lines, other column
    // breaks. A line's row on the page and its column (0 or 1), and the
    // same from (section, first line number, rows, column).
    let at = |p: &Page| -> Vec<(String, i64, usize)> {
        let mut v: Vec<_> =
            p.1.iter()
                .map(|l| {
                    let row = ((l.top - 38.02) / 12.0).round() as i64;
                    (l.label.clone(), row, usize::from(l.x > 150.0))
                })
                .collect();
        v.sort();
        v
    };
    let renumbered = |spec: &[(&str, usize, std::ops::RangeInclusive<i64>, usize)]| {
        let mut v = Vec::new();
        for (prefix, first, range, col) in spec {
            for (n, row) in range.clone().enumerate() {
                v.push((format!("{prefix}-{:02}", first + n), row, *col));
            }
        }
        v.sort();
        v
    };
    // Word: Q2 3 / 2, Q3 3 / 2 below it, Q4 below that.
    let q_word = renumbered(&[
        ("Q1", 1, 0..=1, 0),
        ("Q2", 1, 2..=4, 0),
        ("Q2", 4, 2..=3, 1),
        ("Q3", 1, 5..=7, 0),
        ("Q3", 4, 5..=6, 1),
        ("Q4", 1, 8..=8, 0),
    ]);
    assert_eq!(at(&word[10]), q_word, "Word's Q");
    // Ours: one block of ten lines, 5 / 5, then Q4.
    let q_ours = renumbered(&[
        ("Q1", 1, 0..=1, 0),
        ("Q2", 1, 2..=6, 0),
        ("Q3", 1, 2..=6, 1),
        ("Q4", 1, 7..=7, 0),
    ]);
    assert_eq!(at(&ours[10]), q_ours, "our Q");
    // Word: R2 3 / 2, then R3 in column 1 below it.
    let r_word = renumbered(&[
        ("R1", 1, 0..=1, 0),
        ("R2", 1, 2..=4, 0),
        ("R2", 4, 2..=3, 1),
        ("R3", 1, 5..=9, 0),
    ]);
    assert_eq!(at(&word[11]), r_word, "Word's R");
    // Ours: R2 and R3 fill column 1 in turn.
    let r_ours = renumbered(&[
        ("R1", 1, 0..=1, 0),
        ("R2", 1, 2..=6, 0),
        ("R3", 1, 7..=11, 0),
    ]);
    assert_eq!(at(&ours[11]), r_ours, "our R");

    // (o) the document ends in two columns: not balanced.
    assert_page(&ours[12], &word[12], "Word page 13 (O)");
}

/// On an engine without span/split columns every change opens a page and
/// no paragraph asks for them.
#[test]
fn without_span_columns_each_change_opens_a_page__feat__plugin_doc_word_pagination() {
    let options = LowerOptions {
        mid_page_columns: false,
    };
    let (ir, placed) = lowered(options);
    assert!(placed.iter().all(|p| p.columns == SectionColumns::Frame));
    assert!(!ir
        .styles
        .iter()
        .flat_map(|s| &s.props)
        .any(|p| p.path.contains("SpanColumn") || p.path.contains("SplitColumn")));
    // Every section that changes the columns is a story of its own.
    let changes = COLUMNS_CASES
        .windows(2)
        .filter(|w| {
            w[1].kind == "continuous" && (w[0].columns, w[0].space) != (w[1].columns, w[1].space)
        })
        .count();
    let stories = placed.last().unwrap().story + 1;
    let page_starting = COLUMNS_CASES
        .iter()
        .filter(|c| c.kind == "nextPage")
        .count();
    assert_eq!(stories, page_starting + changes);
}
