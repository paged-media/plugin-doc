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

//! thoughts ADR 028/029 — Word's breaks lower to the engine's break-before
//! rule (`paragraphStartParagraph`, core protocol 64). `breaks_docx()` is the
//! document Word paginated; `fixtures/breaks.word.json` is its answer.
//!
//! What is pinned here:
//! - the import carries `w:pageBreakBefore` and `w:br w:type="page"|"column"`
//!   (no longer a `\n` in the text);
//! - the lowering puts each rule on the paragraph Word starts over with, and
//!   splits a paragraph only where a break sits INSIDE it;
//! - those rules, laid out by the engine's documented rule (ADR 028: start a
//!   new unit unless the paragraph already opens one; odd/even also check
//!   parity), reproduce Word's whole page map;
//! - a split paragraph still saves back as the one Word paragraph it was.

use std::collections::HashMap;

use docx_conformance::breaks_docx;
use docx_core::{Block, BreakKind, RunBreak, SectionKind};
use docx_import::import_docx;
use docx_lower::ir::{LoweredBlock, LoweredDoc, LoweredParagraph, PropValue};

const START: &str = "paragraphStartParagraph";

fn word_map() -> serde_json::Value {
    serde_json::from_str(include_str!("../fixtures/breaks.word.json")).expect("breaks.word.json")
}

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&breaks_docx()).expect("import"))
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

fn rule(ir: &LoweredDoc, style: Option<&str>) -> Option<String> {
    match resolved(ir, style, START) {
        Some(PropValue::Text(t)) if t != "Anywhere" => Some(t),
        _ => None,
    }
}

/// One native paragraph as the pour makes it: its text and its style.
struct Native {
    text: String,
    style: Option<String>,
}

/// The native paragraphs a lowered paragraph pours as (one per segment).
fn natives(p: &LoweredParagraph) -> Vec<Native> {
    let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
    let chars: Vec<char> = text.chars().collect();
    let mut cuts: Vec<(usize, Option<String>)> = vec![(0, p.para_style_id.clone())];
    cuts.extend(
        p.segments
            .iter()
            .map(|s| (s.at as usize, s.para_style_id.clone())),
    );
    (0..cuts.len())
        .map(|k| {
            let end = cuts.get(k + 1).map_or(chars.len(), |c| c.0);
            Native {
                text: chars[cuts[k].0..end].iter().collect(),
                style: cuts[k].1.clone(),
            }
        })
        .collect()
}

fn label(text: &str) -> &str {
    text.split_whitespace().next().unwrap_or("")
}

#[test]
fn breaks_and_page_break_before_import() {
    let doc = import_docx(&breaks_docx()).expect("import");
    let paras: Vec<&docx_core::Paragraph> = doc
        .body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    let text =
        |p: &docx_core::Paragraph| -> String { p.runs.iter().map(|r| r.text.as_str()).collect() };
    let find = |l: &str| *paras.iter().find(|p| text(p).starts_with(l)).unwrap();

    assert_eq!(find("A05").props.page_break_before, Some(true));
    assert_eq!(find("A15").props.page_break_before, Some(true));
    assert_eq!(find("A06").props.page_break_before, None);

    // The paragraph holding only a page break: no text, one break at 0.
    let empty: Vec<_> = paras.iter().filter(|p| text(p).is_empty()).collect();
    assert_eq!(
        empty.len(),
        2,
        "the page-break and the column-break paragraphs"
    );
    let only = |p: &docx_core::Paragraph| -> Vec<RunBreak> {
        p.runs.iter().flat_map(|r| r.breaks.clone()).collect()
    };
    assert_eq!(
        only(empty[0]),
        vec![RunBreak {
            at: 0,
            kind: BreakKind::Page
        }]
    );
    assert_eq!(
        only(empty[1]),
        vec![RunBreak {
            at: 0,
            kind: BreakKind::Column
        }]
    );

    // A break INSIDE a run sits at its char offset, out of the text.
    let a21 = find("A21a");
    assert_eq!(a21.runs.len(), 1);
    assert_eq!(a21.runs[0].text, "A21a beforeA21b after mid break");
    assert_eq!(
        a21.runs[0].breaks,
        vec![RunBreak {
            at: 11,
            kind: BreakKind::Page
        }]
    );
    let b08 = find("B08a");
    assert_eq!(
        b08.runs[0].breaks,
        vec![RunBreak {
            at: 8,
            kind: BreakKind::Column
        }]
    );
    // A break after the run's text (A19's own break run).
    assert_eq!(
        only(find("A19")),
        vec![RunBreak {
            at: 0,
            kind: BreakKind::Page
        }]
    );
    assert!(
        paras.iter().all(|p| !text(p).contains('\n')),
        "a page/column break is no longer a line break in the text"
    );

    let kinds: Vec<SectionKind> = doc.sections.iter().map(|s| s.kind).collect();
    assert_eq!(
        kinds,
        vec![
            SectionKind::NextPage,
            SectionKind::NextPage,
            SectionKind::OddPage,
            SectionKind::OddPage,
            SectionKind::EvenPage,
            SectionKind::EvenPage
        ]
    );
    assert_eq!(doc.sections[1].columns, 2);
}

/// Every native paragraph's break-before rule, by label. This is what the
/// lowering emits for every case of the fixture.
#[test]
fn every_break_lowers_to_the_rule_on_the_paragraph_that_starts_over() {
    let ir = lowered();
    let mut rules: Vec<(String, Option<String>)> = Vec::new();
    for block in &ir.story.blocks {
        let LoweredBlock::Paragraph(p) = block else {
            panic!("the fixture has no tables")
        };
        for n in natives(p) {
            rules.push((label(&n.text).to_string(), rule(&ir, n.style.as_deref())));
        }
    }
    let ruled: Vec<(&str, &str)> = rules
        .iter()
        .filter_map(|(l, r)| r.as_deref().map(|r| (l.as_str(), r)))
        .collect();
    assert_eq!(
        ruled,
        vec![
            ("A05", "NextPage"),   // w:pageBreakBefore mid-page
            ("A15", "NextPage"),   // w:pageBreakBefore, already opening a page
            ("A18", "NextPage"),   // after the paragraph holding only a page break
            ("A20", "NextPage"),   // after A19's trailing page break
            ("A21b", "NextPage"),  // the split part after a mid-paragraph break
            ("B04", "NextColumn"), // after the paragraph holding only a column break
            ("B08b", "NextColumn"),
            ("C01", "NextOddPage"),
            ("D01", "NextOddPage"),
            ("E01", "NextEvenPage"),
            ("F01", "NextEvenPage"),
        ],
        "all rules: {rules:?}"
    );

    // The break paragraphs themselves stay (empty, unruled): Word keeps
    // their paragraph mark where the break is.
    assert_eq!(rules.iter().filter(|(l, _)| l.is_empty()).count(), 2);
    // One Word paragraph = one block, split only where a break is inside it.
    assert_eq!(ir.story.blocks.len(), 46);
    let split: Vec<(u32, u32)> = ir
        .story
        .blocks
        .iter()
        .filter_map(|b| match b {
            LoweredBlock::Paragraph(p) if !p.segments.is_empty() => {
                Some((p.source_index, p.segments[0].at))
            }
            _ => None,
        })
        .collect();
    assert_eq!(split.len(), 2);
    assert_eq!(split[0].1, 11, "A21: after 'A21a before'");
    assert_eq!(split[1].1, 8, "B08: after 'B08a col'");
}

/// The part after a break inside a Word paragraph is not a new Word
/// paragraph: no space before, no first-line indent; the part before it no
/// space after.
#[test]
fn a_split_paragraphs_parts_keep_one_paragraphs_spacing() {
    let ir = lowered();
    let p = ir
        .story
        .paragraphs()
        .into_iter()
        .find(|p| !p.segments.is_empty())
        .unwrap();
    let seg = p.segments[0].para_style_id.as_deref();
    assert_eq!(
        resolved(&ir, seg, "paragraphSpaceBefore"),
        Some(PropValue::Length(0.0))
    );
    assert_eq!(
        resolved(&ir, seg, "paragraphFirstLineIndent"),
        Some(PropValue::Length(0.0))
    );
    assert_eq!(
        resolved(&ir, p.para_style_id.as_deref(), "paragraphSpaceAfter"),
        Some(PropValue::Length(0.0))
    );
    assert_eq!(rule(&ir, p.para_style_id.as_deref()), None);
    // The exact 12 pt grid holds for every part.
    assert_eq!(
        resolved(&ir, seg, "characterLeading"),
        Some(PropValue::Length(12.0))
    );
}

/// The lowered rules, laid out by the engine's rule (ADR 028, as measured
/// against InDesign: a rule opens a new column / page unless the paragraph
/// already opens one; odd/even pages also check parity, with a blank page
/// when the next page has the wrong one), on the skeleton's geometry (each
/// section its own story starting on its own new page, ten lines per column),
/// reproduce Word's page map line for line.
#[test]
fn the_rules_reproduce_words_page_map() {
    let ir = lowered();
    const LINES: usize = 10;
    // pages[page][column] = labels
    let mut pages: Vec<Vec<Vec<String>>> = Vec::new();
    for (k, sec) in ir.sections.iter().enumerate() {
        let cols = sec.columns as usize;
        let new_page = |pages: &mut Vec<Vec<Vec<String>>>| pages.push(vec![Vec::new(); cols]);
        new_page(&mut pages);
        let (mut col, mut line) = (0usize, 0usize);
        let end = ir
            .sections
            .get(k + 1)
            .map_or(ir.story.blocks.len(), |s| s.first_block);
        for block in &ir.story.blocks[sec.first_block..end] {
            let LoweredBlock::Paragraph(p) = block else {
                unreachable!()
            };
            for n in natives(p) {
                let advance_page = |pages: &mut Vec<Vec<Vec<String>>>| {
                    new_page(pages);
                };
                // Anything on the current page yet (lines fill in order).
                let used = (col, line) != (0, 0);
                match rule(&ir, n.style.as_deref()).as_deref() {
                    Some("NextColumn") if line > 0 => {
                        if col + 1 < cols {
                            (col, line) = (col + 1, 0);
                        } else {
                            advance_page(&mut pages);
                            (col, line) = (0, 0);
                        }
                    }
                    Some("NextPage") if used => {
                        advance_page(&mut pages);
                        (col, line) = (0, 0);
                    }
                    Some(r @ ("NextOddPage" | "NextEvenPage")) => {
                        if used {
                            advance_page(&mut pages);
                            (col, line) = (0, 0);
                        }
                        let want_odd = r == "NextOddPage";
                        while (pages.len() % 2 == 1) != want_odd {
                            advance_page(&mut pages);
                        }
                    }
                    _ => {}
                }
                if line == LINES {
                    if col + 1 < cols {
                        (col, line) = (col + 1, 0);
                    } else {
                        advance_page(&mut pages);
                        (col, line) = (0, 0);
                    }
                }
                let l = label(&n.text);
                if !l.is_empty() {
                    pages.last_mut().unwrap()[col].push(l.to_string());
                }
                line += 1;
            }
        }
    }

    let word = word_map();
    let word_pages = word["pages"].as_array().unwrap();
    assert_eq!(pages.len(), word_pages.len(), "page count: {pages:?}");
    for (i, (ours, theirs)) in pages.iter().zip(word_pages).enumerate() {
        let mut cols: Vec<Vec<String>> = ours.clone();
        while cols.len() > 1 && cols.last().is_some_and(Vec::is_empty) {
            cols.pop();
        }
        let expected: Vec<Vec<String>> = if theirs.get("blank").is_some() {
            vec![Vec::new()]
        } else if let Some(c) = theirs.get("columns") {
            serde_json::from_value(c.clone()).unwrap()
        } else {
            vec![serde_json::from_value(theirs["labels"].clone()).unwrap()]
        };
        assert_eq!(cols, expected, "page {}", i + 1);
    }
}

/// Save-back folds a split paragraph's native parts back into its one Word
/// paragraph: an identity read-back changes nothing, and an edit AFTER the
/// split still lands on its own `w:p`.
#[test]
fn a_split_paragraph_saves_back_as_one_word_paragraph() {
    use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};

    let original = breaks_docx();
    let session = docx_js::DocSession::load(&original).unwrap();
    let ir = session.lowered();
    // What the host reads back: one paragraph per native paragraph, a run
    // split in two where a break cut it.
    let read_back = || StoryContentIn {
        self_id: "docx_s0".into(),
        paragraphs: ir
            .story
            .blocks
            .iter()
            .flat_map(|b| {
                let LoweredBlock::Paragraph(p) = b else {
                    unreachable!()
                };
                let cuts: Vec<usize> = p.segments.iter().map(|s| s.at as usize).collect();
                let mut parts = vec![ParagraphContentIn {
                    paragraph_style: p.para_style_id.clone(),
                    runs: Vec::new(),
                }];
                let mut at = 0usize;
                for r in &p.runs {
                    let mut piece = String::new();
                    for ch in r.text.chars() {
                        if cuts.contains(&at) && at > 0 {
                            if !piece.is_empty() {
                                parts.last_mut().unwrap().runs.push(RunContentIn {
                                    text: std::mem::take(&mut piece),
                                    character_style: r.char_style_id.clone(),
                                });
                            }
                            parts.push(ParagraphContentIn {
                                paragraph_style: None,
                                runs: Vec::new(),
                            });
                        }
                        piece.push(ch);
                        at += 1;
                    }
                    parts.last_mut().unwrap().runs.push(RunContentIn {
                        text: piece,
                        character_style: r.char_style_id.clone(),
                    });
                }
                parts
            })
            .collect(),
    };
    let identity = read_back();
    assert_eq!(identity.paragraphs.len(), 46 + 2, "two splits pour 2 extra");
    let (saved, _) = session.save_edited_from_content(&identity).unwrap();
    let a = paged_ooxml::OpcPackage::read(&original).unwrap();
    let b = paged_ooxml::OpcPackage::read(&saved).unwrap();
    for name in a.file_names() {
        assert_eq!(a.part(name), b.part(name), "part {name} unchanged");
    }

    // Edit A22 — the paragraph right after the split one.
    let mut edited = read_back();
    let i = edited
        .paragraphs
        .iter()
        .position(|p| p.runs.first().is_some_and(|r| r.text.starts_with("A22")))
        .unwrap();
    edited.paragraphs[i].runs[0].text = "A22 edited".into();
    let (saved, _) = session.save_edited_from_content(&edited).unwrap();
    let re = import_docx(&saved).unwrap();
    let texts: Vec<String> = re
        .body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        })
        .collect();
    let at = texts
        .iter()
        .position(|t| t == "A22 edited")
        .expect("edit landed");
    assert_eq!(texts[at - 1], "A21a beforeA21b after mid break");
    assert_eq!(texts[at + 1], "A23 last of A");
    let a21 = match &re.body[at - 1] {
        Block::Paragraph(p) => p,
        _ => unreachable!(),
    };
    assert_eq!(
        a21.runs[0].breaks,
        vec![RunBreak {
            at: 11,
            kind: BreakKind::Page
        }],
        "the untouched split paragraph keeps its break"
    );
}

/// ADR 029 — a paragraph with no line spacing anywhere (direct, style chain,
/// docDefaults) is laid by Word at single spacing, the face's own line
/// (`fixtures/line-spacing.word.json` L1: Calibri 10 pt, 12.21 pt), not at
/// the engine's 120% of the point size. The lowering now says so.
#[test]
fn unset_line_spacing_lowers_to_words_single() {
    let doc = import_docx(&docx_conformance::memo_docx()).expect("import");
    assert!(doc.styles.doc_defaults.para.line_spacing.is_none());
    let ir = docx_lower::lower(&doc);
    let paras = ir.story.paragraphs();
    let leading = |p: &LoweredParagraph| match resolved(
        &ir,
        p.para_style_id.as_deref(),
        "characterLeading",
    ) {
        Some(PropValue::Length(v)) => v,
        other => panic!("no leading: {other:?}"),
    };
    // Plain body text: no size anywhere → Word's 10 pt default, theme face.
    let single = docx_lower::line_height::FALLBACK_EM * 10.0;
    assert!(
        (leading(paras[0]) - single).abs() < 1e-3,
        "{}",
        leading(paras[0])
    );
    // Heading1 (24 pt): single of 24 pt.
    let heading = docx_lower::line_height::FALLBACK_EM * 24.0;
    assert!(
        (leading(paras[1]) - heading).abs() < 1e-3,
        "{}",
        leading(paras[1])
    );
}
