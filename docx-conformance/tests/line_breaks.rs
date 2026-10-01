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

//! Word's plain line breaks and blank lines (core `ab383b1` / `65cf615`).
//! `line_breaks_docx()` is the document Word laid out;
//! `fixtures/line-breaks.word.json` is its answer.
//!
//! What is pinned here:
//! - the import turns `<w:br/>`, `<w:br w:type="textWrapping"/>` and
//!   `<w:cr/>` into U+2028 (a line break INSIDE the paragraph), never `\n`;
//! - each Word paragraph lowers to ONE paragraph, and a blank line carries a
//!   paragraph style with its pitch;
//! - those paragraphs, laid out by the engine's rules (a U+2028 starts a new
//!   line, a trailing one an empty line; a blank line is one line at its
//!   style's leading; blank lines at one offset take the last caret's style),
//!   put every line where Word does — except the one documented case, the
//!   blank lines with different pitches, which is diagnosed;
//! - save-back writes U+2028 back as `<w:br/>` where it was, and a zero-edit
//!   save stays byte-identical.

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use std::collections::HashMap;

use docx_conformance::{line_breaks_docx, LINE_BREAK_CASES};
use docx_core::{Block, LINE_BREAK};
use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};
use docx_import::import_docx;
use docx_js::DocSession;
use docx_lower::ir::{LoweredBlock, LoweredDoc, LoweredParagraph, PropValue};
use paged_ooxml::OpcPackage;

fn word_map() -> serde_json::Value {
    serde_json::from_str(include_str!("../fixtures/line-breaks.word.json"))
        .expect("line-breaks.word.json")
}

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&line_breaks_docx()).expect("import"))
}

fn paragraphs(ir: &LoweredDoc) -> Vec<&LoweredParagraph> {
    ir.story
        .blocks
        .iter()
        .map(|b| match b {
            LoweredBlock::Paragraph(p) => p,
            LoweredBlock::Table(_) => panic!("no tables in the fixture"),
        })
        .collect()
}

/// A paragraph style property, resolved through `basedOn` (nearest first).
fn resolved(ir: &LoweredDoc, style: Option<&str>, path: &str) -> Option<PropValue> {
    let by_id: HashMap<&str, _> = ir.styles.iter().map(|s| (s.id.as_str(), s)).collect();
    let mut next = style;
    while let Some(id) = next {
        let s = by_id.get(id)?;
        if let Some(p) = s.props.iter().rev().find(|p| p.path == path) {
            return Some(p.value.clone());
        }
        next = s.based_on.as_deref();
    }
    None
}

fn leading(ir: &LoweredDoc, style: Option<&str>) -> f32 {
    match resolved(ir, style, "characterLeading") {
        Some(PropValue::Length(v)) => v,
        other => panic!("no leading on {style:?}: {other:?}"),
    }
}

fn text(p: &LoweredParagraph) -> String {
    p.runs.iter().map(|r| r.text.as_str()).collect()
}

#[test]
fn plain_breaks_import_as_u2028() {
    let doc = import_docx(&line_breaks_docx()).expect("import");
    let texts: Vec<String> = doc
        .body
        .iter()
        .map(|b| match b {
            Block::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect(),
            Block::Table(_) => panic!("no tables"),
        })
        .collect();
    assert_eq!(texts.len(), LINE_BREAK_CASES.len());
    assert!(texts.iter().all(|t| !t.contains('\n')), "{texts:?}");
    let by = |label: &str| &texts[LINE_BREAK_CASES.iter().position(|c| c.0 == label).unwrap()];
    assert_eq!(by("L02"), "L02a before\u{2028}L02b after");
    assert_eq!(by("L03"), "L03a before\u{2028}\u{2028}L03c after two");
    assert_eq!(by("L04"), "L04a ends in br\u{2028}");
    assert_eq!(by("L05"), "L05a wrap\u{2028}L05b own run");
    assert_eq!(by("L06"), "L06a cr\u{2028}L06b after cr");
    assert_eq!(by("B01"), "");
    // No plain break became a page/column break.
    let breaks: usize = doc
        .body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.breaks.len()).sum::<usize>()),
            _ => None,
        })
        .sum();
    assert_eq!(breaks, 0);
}

#[test]
fn one_word_paragraph_is_one_native_paragraph_and_blank_lines_are_styled() {
    let ir = lowered();
    let paras = paragraphs(&ir);
    assert_eq!(paras.len(), LINE_BREAK_CASES.len());
    for (p, (label, pitch, runs)) in paras.iter().zip(LINE_BREAK_CASES) {
        assert!(p.segments.is_empty(), "{label}: a line break never splits");
        assert_eq!(runs.is_empty(), p.runs.is_empty(), "{label}");
        assert!(p.para_style_id.is_some(), "{label} has a paragraph style");
        assert_eq!(
            leading(&ir, p.para_style_id.as_deref()),
            *pitch as f32 / 20.0,
            "{label}'s pitch"
        );
    }
}

/// Lay the lowered paragraphs out on the engine's rules: page breaks from
/// the break-before rule, a line per U+2028-separated part (a trailing break
/// leaves an empty last line), a blank line one line at its leading, and
/// blank lines that share an offset all at the LAST one's leading (the caret
/// the pour sends last). Returns, per page, each visible line's label and
/// its top in 12 pt grid lines.
///
/// This is the layout the lowering ASKS for. The engine on core-p64
/// (`df63a5f`) does not yet give it for blank lines: an empty paragraph with
/// no runs is laid at auto leading (1.2 × its style's point size) instead of
/// its paragraph style's leading — the editor spec `doc-line-breaks.spec.ts`
/// AC-DOCLB-2 measures that and stays red until core honours the style.
fn engine_layout(ir: &LoweredDoc) -> Vec<Vec<(String, u32)>> {
    let paras = paragraphs(ir);
    let mut pitch: Vec<f32> = paras
        .iter()
        .map(|p| leading(ir, p.para_style_id.as_deref()))
        .collect();
    // Blank lines at one offset: the last caret wins.
    let mut k = 0;
    while k < paras.len() {
        let mut j = k;
        while j < paras.len() && paras[j].runs.is_empty() {
            j += 1;
        }
        if j > k {
            let last = pitch[j - 1];
            pitch[k..j].iter_mut().for_each(|v| *v = last);
            k = j;
        } else {
            k += 1;
        }
    }
    let mut pages: Vec<Vec<(String, u32)>> = vec![vec![]];
    let mut y = 0.0f32;
    for (p, lead) in paras.iter().zip(pitch) {
        let starts_page = matches!(
            resolved(ir, p.para_style_id.as_deref(), "paragraphStartParagraph"),
            Some(PropValue::Text(t)) if t == "NextPage"
        );
        if starts_page && y > 0.0 {
            pages.push(vec![]);
            y = 0.0;
        }
        let t = text(p);
        for line in t.split(LINE_BREAK) {
            if let Some(label) = line.split_whitespace().next() {
                pages
                    .last_mut()
                    .unwrap()
                    .push((label.to_string(), (y / 12.0).round() as u32));
            }
            y += lead;
        }
        assert!(y <= 240.0, "the fixture never overflows a page");
    }
    pages
}

fn word_layout() -> Vec<Vec<(String, u32)>> {
    word_map()["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|p| {
            p["lines"]
                .as_array()
                .unwrap()
                .iter()
                .map(|l| {
                    (
                        l["label"].as_str().unwrap().to_string(),
                        l["line"].as_u64().unwrap() as u32,
                    )
                })
                .collect()
        })
        .collect()
}

#[test]
fn the_engine_lays_every_line_where_word_does() {
    let ir = lowered();
    let engine = engine_layout(&ir);
    let word = word_layout();
    assert_eq!(engine.len(), word.len(), "page count");
    // Page 1: every line break and every same-style blank line, exactly.
    assert_eq!(engine[0], word[0]);
    // Page 2: B04 (24 pt) + B05 (12 pt) take 3 lines in Word, but blank
    // lines at one offset share the last caret's style: 2 lines here, so L10
    // sits one line higher. Diagnosed, below.
    assert_eq!(
        word[1],
        vec![("L09".to_string(), 0), ("L10".to_string(), 4)]
    );
    assert_eq!(
        engine[1],
        vec![("L09".to_string(), 0), ("L10".to_string(), 3)]
    );
}

#[test]
fn blank_lines_with_different_styles_are_diagnosed() {
    let ir = lowered();
    let index = |l: &str| LINE_BREAK_CASES.iter().position(|c| c.0 == l).unwrap();
    let warnings: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.message.starts_with("blank lines at body blocks"))
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warnings.len(), 1, "{warnings:?}");
    let w = warnings[0];
    assert!(w.contains(&format!("{} (", index("B04"))), "{w}");
    assert!(w.contains(&format!("{} (", index("B05"))), "{w}");
    // B02/B03 share a style: not reported.
    assert!(!w.contains(&format!("{} (", index("B02"))), "{w}");
    let b05 = paragraphs(&ir)[index("B05")].para_style_id.clone().unwrap();
    assert!(w.contains(&format!("all take {b05}")), "{w}");
}

/// The read-back the editor would hand save-back: the baseline, as poured.
fn identity_content(ir: &LoweredDoc) -> StoryContentIn {
    StoryContentIn {
        self_id: "Story/doc".into(),
        paragraphs: paragraphs(ir)
            .iter()
            .map(|p| ParagraphContentIn {
                paragraph_style: p.para_style_id.clone(),
                runs: p
                    .runs
                    .iter()
                    .map(|r| RunContentIn {
                        text: r.text.clone(),
                        character_style: r.char_style_id.clone(),
                    })
                    .collect(),
            })
            .collect(),
    }
}

#[test]
fn zero_edit_save_with_line_breaks_is_byte_identical() {
    let original = line_breaks_docx();
    let session = DocSession::load(&original).unwrap();
    let (saved, _) = session
        .save_edited_from_content(&identity_content(&lowered()))
        .unwrap();
    let (a, b) = (
        OpcPackage::read(&original).unwrap(),
        OpcPackage::read(&saved).unwrap(),
    );
    for name in a.file_names() {
        assert_eq!(a.part(name), b.part(name), "part {name}");
    }
}

#[test]
fn edited_line_breaks_save_back_as_w_br_in_place() {
    let original = line_breaks_docx();
    let session = DocSession::load(&original).unwrap();
    let ir = lowered();
    let mut content = identity_content(&ir);
    let index = |l: &str| LINE_BREAK_CASES.iter().position(|c| c.0 == l).unwrap();
    // Edit the second line of L02, add a break to L01, drop L06's `w:cr`.
    content.paragraphs[index("L02")].runs[0].text = "L02a before\u{2028}L02b EDITED".into();
    content.paragraphs[index("L01")].runs[0].text = "L01 one\u{2028}line".into();
    content.paragraphs[index("L06")].runs[0].text = "L06a cr L06b after cr".into();
    let (saved, skips) = session.save_edited_from_content(&content).unwrap();
    assert!(skips.is_empty(), "{skips:?}");
    let pkg = OpcPackage::read(&saved).unwrap();
    let xml = std::str::from_utf8(pkg.part("word/document.xml").unwrap()).unwrap();
    assert!(
        xml.contains(r#"<w:t xml:space="preserve">L02a before</w:t><w:br/><w:t xml:space="preserve">L02b EDITED</w:t></w:r>"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<w:t xml:space="preserve">L01 one</w:t><w:br/><w:t xml:space="preserve">line</w:t></w:r>"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<w:t xml:space="preserve">L06a cr L06b after cr</w:t></w:r>"#),
        "{xml}"
    );
    assert!(!xml.contains("<w:cr/>"), "{xml}");
    // Untouched runs keep their bytes: L03's two breaks, L05's own break run.
    assert!(xml.contains(r#"<w:t xml:space="preserve">L03a before</w:t><w:br/><w:br/><w:t xml:space="preserve">L03c after two</w:t>"#));
    assert!(xml.contains(r#"<w:br w:type="textWrapping"/>"#));
    // And it re-imports to the edited text.
    let re = import_docx(&saved).unwrap();
    let ir2 = docx_lower::lower(&re);
    let t = |l: &str| text(paragraphs(&ir2)[index(l)]);
    assert_eq!(t("L02"), "L02a before\u{2028}L02b EDITED");
    assert_eq!(t("L01"), "L01 one\u{2028}line");
    assert_eq!(t("L04"), "L04a ends in br\u{2028}");
}

#[test]
fn blank_lines_in_a_cell_are_styled_and_differences_diagnosed__feat__plugin_doc_read_path() {
    let ir = docx_lower::lower(&import_docx(&docx_conformance::cell_blank_lines_docx()).unwrap());
    let LoweredBlock::Table(t) = &ir.story.blocks[0] else {
        panic!("a table first");
    };
    // Every blank cell line carries its paragraph style, with its pitch.
    let pitches: Vec<Vec<f32>> = t
        .cells
        .iter()
        .map(|c| {
            c.paragraphs
                .iter()
                .map(|p| leading(&ir, p.para_style_id.as_deref()))
                .collect()
        })
        .collect();
    assert_eq!(
        pitches,
        vec![
            vec![12.0, 24.0, 12.0, 12.0],
            vec![12.0, 24.0, 12.0, 24.0],
            vec![12.0]
        ]
    );
    // Cell (0, 0)'s two blank lines share one offset and differ: diagnosed.
    // Cell (0, 1)'s blank lines sit at different offsets: not.
    let warnings: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.message.starts_with("blank lines in the table"))
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warnings.len(), 1, "{warnings:?}");
    assert!(
        warnings[0].contains("cell row 0 column 0, paragraphs 1 ("),
        "{}",
        warnings[0]
    );
    assert!(warnings[0].contains(", 2 ("), "{}", warnings[0]);
}
