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
//!   column count changes straight into another (H, N, P);
//! - a page-starting multi-column section is balanced the same way when a
//!   continuous section follows (I1: 4 / 3);
//! - a section before a `nextPage` section, or at the document's end, is
//!   NOT balanced: its lines fill column 1 first (I3, M2, O2);
//! - a section running past the page fills that page's columns to the
//!   bottom (17 / 17) and is balanced on its last page only if a continuous
//!   break follows (L2: 8 / 8; M2: 6 in column 1);
//! - the gap is `w:space` (J2: 18 pt); unequal columns keep their widths
//!   (K2: 180 pt and 90 pt, balanced 4 / 2).
//!
//! What the lowering reproduces, line for line under the engine's span /
//! split rule (`support/layout.rs`): I, J, L, M, O, and the first section
//! pair of H, N and P. Pinned differences: where one multi-column section
//! follows another (H3, N3, P3) the engine would merge them into one split
//! block, so the second opens a page; unequal columns are equal.
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
            // (h) H2 must end balanced (H3 follows continuous) and H3 cannot
            //     follow it in the same story (the engine would merge the two
            //     split blocks): H1 + H2, then H3 + H4 on a page of their own.
            ("H1", 0, 1, Frame),
            ("H2", 0, 1, split(2, 36.0)),
            ("H3", 1, 1, split(3, 36.0)),
            ("H4", 1, 1, Frame),
            // (i) the last section is left unbalanced: the frame keeps its
            //     columns (filled in turn after the last span), the
            //     one-column section spans them.
            ("I1", 2, 2, Frame),
            ("I2", 2, 2, SpanAll),
            ("I3", 2, 2, Frame),
            // (j) split, at Word's gap.
            ("J1", 3, 1, Frame),
            ("J2", 3, 1, split(2, 18.0)),
            ("J3", 3, 1, Frame),
            // (k) unequal columns, as equal ones at the first one's gap.
            ("K1", 4, 1, Frame),
            ("K2", 4, 1, split(2, 18.0)),
            ("K3", 4, 1, Frame),
            // (l) split: the block fills the page, then balances the rest.
            ("L1", 5, 1, Frame),
            ("L2", 5, 1, split(2, 36.0)),
            ("L3", 5, 1, Frame),
            // (m) unbalanced last section: the frame's columns.
            ("M1", 6, 2, SpanAll),
            ("M2", 6, 2, Frame),
            ("N1", 7, 1, Frame),
            ("N2", 7, 1, split(3, 36.0)),
            ("N3", 8, 1, split(2, 36.0)),
            ("N4", 8, 1, Frame),
            ("P1", 9, 1, Frame),
            ("P2", 9, 1, split(2, 36.0)),
            ("P3", 10, 1, split(2, 18.0)),
            ("P4", 10, 1, Frame),
            // (o) the document ends unbalanced: the frame's columns.
            ("O1", 11, 2, SpanAll),
            ("O2", 11, 2, Frame),
        ]
    );
    assert_eq!(placed[4].frame.gutter_pt, 36.0, "I's 0.5 in gap");

    let warnings: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.severity == "warning")
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warnings.len(), 4, "{warnings:?}");
    assert!(
        warnings[1].starts_with("section 12 ") && warnings[1].contains("unequal"),
        "{}",
        warnings[1]
    );
    for (w, n) in [warnings[0], warnings[2], warnings[3]].iter().zip([
        "section 3 ",
        "section 21 ",
        "section 25 ",
    ]) {
        assert!(
            w.starts_with(n) && w.contains("one split-column block") && w.contains("new page"),
            "{w}"
        );
    }
}

/// The page map, line by line (label, x within 0.2 pt, top within
/// 0.25 pt), with the pinned differences.
#[test]
fn the_lowering_reproduces_words_column_changes__feat__plugin_doc_word_pagination() {
    let (ir, placed) = lowered(LowerOptions::default());
    let ours = lay_out(&ir, &placed);
    let word = word();
    assert_eq!(word.len(), 11);
    // Three column changes into another column count each cost a page.
    assert_eq!(ours.len(), word.len() + 3);
    let lines_of = |p: &Page, prefix: &str| -> Page {
        (
            p.0.clone(),
            p.1.iter()
                .filter(|l| l.label.starts_with(prefix))
                .cloned()
                .collect(),
        )
    };
    let shifted = |p: &Page, rows: f64| -> Page {
        let mut p = p.clone();
        for l in &mut p.1 {
            l.top -= rows * 12.0;
        }
        p
    };

    // (h) Word page 1: H1, H2 balanced 3 / 3, H3 balanced 3 / 3 / 1, H4.
    // Ours: H1 and H2 exactly; H3 and H4 the same lines on the next page,
    // four rows (H1 + H2's three) higher.
    let w = &word[0];
    for prefix in ["H1", "H2"] {
        assert_page(&lines_of(&ours[0], prefix), &lines_of(w, prefix), prefix);
    }
    assert_eq!(labels(&ours[0]).len(), 2 + 6);
    let rest: Page = (
        w.0.clone(),
        w.1.iter()
            .filter(|l| l.label.starts_with("H3") || l.label.starts_with("H4"))
            .cloned()
            .collect(),
    );
    assert_page(&ours[1], &shifted(&rest, 5.0), "H3 + H4 on their own page");

    // (i) balanced before the span, not before the nextPage section.
    assert_page(&ours[2], &word[1], "Word page 2 (I)");
    // (j) Word's 0.25 in gap.
    assert_page(&ours[3], &word[2], "Word page 3 (J)");

    // (k) unequal columns: same lines on the same page; K2's second column
    // sits at the equal-column x (189 pt, Word 234.17) and is balanced 3 / 3
    // (Word 4 / 2).
    let (o, w) = (&ours[4], &word[3]);
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
    assert_page(&ours[5], &word[4], "Word page 5 (L)");
    assert_page(&ours[6], &word[5], "Word page 6 (L)");
    // (m) past the page before a nextPage section: not balanced.
    assert_page(&ours[7], &word[6], "Word page 7 (M)");
    assert_page(&ours[8], &word[7], "Word page 8 (M)");

    // (n) 3 → 2: N1 and N2 (3 / 3 / 1) exactly; N3 and N4 on the next
    // page, five rows higher.
    let w = &word[8];
    for prefix in ["N1", "N2"] {
        assert_page(&lines_of(&ours[9], prefix), &lines_of(w, prefix), prefix);
    }
    let rest: Page = (
        w.0.clone(),
        w.1.iter()
            .filter(|l| l.label.starts_with("N3") || l.label.starts_with("N4"))
            .cloned()
            .collect(),
    );
    assert_page(&ours[10], &shifted(&rest, 5.0), "N3 + N4 on their own page");

    // (p) two columns into two with another gap: P1 and P2 exactly; P3 and
    // P4 on the next page, four rows higher.
    let w = &word[9];
    for prefix in ["P1", "P2"] {
        assert_page(&lines_of(&ours[11], prefix), &lines_of(w, prefix), prefix);
    }
    let rest: Page = (
        w.0.clone(),
        w.1.iter()
            .filter(|l| l.label.starts_with("P3") || l.label.starts_with("P4"))
            .cloned()
            .collect(),
    );
    assert_page(&ours[12], &shifted(&rest, 4.0), "P3 + P4 on their own page");

    // (o) the document ends in two columns: not balanced.
    assert_page(&ours[13], &word[10], "Word page 11 (O)");
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
