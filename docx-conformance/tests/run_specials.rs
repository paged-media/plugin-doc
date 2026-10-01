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

//! Tabs, non-breaking hyphens and the run content around them, through
//! save-back (`run_specials_docx()`).
//!
//! What is pinned here:
//! - the import turns `<w:tab/>` into `\t` and `<w:noBreakHyphen/>` into
//!   U+2011 in the run's text, and drops `<w:sym>`, `<w:softHyphen/>` and
//!   `<w:ptab>` (they have no character in the engine's text);
//! - a zero-edit save stays byte-identical;
//! - an edited run writes each of those characters back the way the source
//!   run wrote it — `<w:tab/>` / `<w:noBreakHyphen/>` / `<w:br …/>` elements
//!   verbatim and in position, a literal tab as a literal tab — and a NEW
//!   one as Word's element (never a tab character folded into `<w:t>`);
//! - an edit to a run holding content the text cannot place is refused into
//!   the skip ledger, and that run keeps its bytes.

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::{run_specials_docx, RUN_SPECIAL_CASES};
use docx_core::Block;
use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};
use docx_import::import_docx;
use docx_js::DocSession;
use docx_lower::ir::{LoweredBlock, LoweredDoc, LoweredParagraph};
use paged_ooxml::OpcPackage;

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&run_specials_docx()).expect("import"))
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

fn index(label: &str) -> usize {
    RUN_SPECIAL_CASES.iter().position(|c| c.0 == label).unwrap()
}

fn text(p: &LoweredParagraph) -> String {
    p.runs.iter().map(|r| r.text.as_str()).collect()
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

fn document_xml(docx: &[u8]) -> String {
    let pkg = OpcPackage::read(docx).unwrap();
    String::from_utf8(pkg.part("word/document.xml").unwrap().to_vec()).unwrap()
}

/// The source bytes of the `<w:p>` holding `label`.
fn paragraph_xml<'a>(xml: &'a str, label: &str) -> &'a str {
    let at = xml
        .find(label)
        .unwrap_or_else(|| panic!("{label} in {xml}"));
    let start = xml[..at].rfind("<w:p>").unwrap();
    let end = at + xml[at..].find("</w:p>").unwrap() + "</w:p>".len();
    &xml[start..end]
}

#[test]
fn tabs_and_no_break_hyphens_import_as_characters__feat__plugin_doc_save_back() {
    let doc = import_docx(&run_specials_docx()).expect("import");
    let texts: Vec<String> = doc
        .body
        .iter()
        .map(|b| match b {
            Block::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect(),
            Block::Table(_) => panic!("no tables"),
        })
        .collect();
    let by = |l: &str| texts[index(l)].as_str();
    assert_eq!(by("T01"), "T01\tvalue");
    assert_eq!(by("T02"), "T02a\t\tT02b");
    assert_eq!(by("T03"), "\tT03 leading");
    assert_eq!(by("N01"), "N01 well\u{2011}known");
    assert_eq!(by("M01"), "M01a\tb\u{2028}M01c\u{2011}d");
    assert_eq!(by("X01"), "X01\tliteral");
    // Not carried as text.
    assert_eq!(by("S01"), "S01 beforeafter");
    assert_eq!(by("H01"), "H01 hyphenation");
    assert_eq!(by("P01"), "P01 leftright");
}

#[test]
fn zero_edit_save_with_tabs_and_symbols_is_byte_identical__feat__plugin_doc_save_back() {
    let original = run_specials_docx();
    let session = DocSession::load(&original).unwrap();
    let (saved, skips) = session
        .save_edited_from_content(&identity_content(&lowered()))
        .unwrap();
    assert!(skips.is_empty(), "{skips:?}");
    let (a, b) = (
        OpcPackage::read(&original).unwrap(),
        OpcPackage::read(&saved).unwrap(),
    );
    for name in a.file_names() {
        assert_eq!(a.part(name), b.part(name), "part {name}");
    }
}

/// The edits [`edited_save`] makes, by label: the new text of the
/// paragraph's one run.
const EDITS: &[(&str, &str)] = &[
    ("T01", "T01\tEDITED"),
    // A third word after a new tab.
    ("T02", "T02a\tT02b\tT02c"),
    ("T03", "\tT03 leading EDITED"),
    ("N01", "N01 well\u{2011}known EDITED"),
    ("M01", "M01a\tB\u{2028}M01c\u{2011}D"),
    ("X01", "X01\tLITERAL"),
    // Content the text cannot place: all three refused.
    ("S01", "S01 before EDITED after"),
    ("H01", "H01 hyphenation EDITED"),
    ("P01", "P01 left EDITED right"),
];

fn edited_save() -> (Vec<u8>, Vec<String>) {
    let session = DocSession::load(&run_specials_docx()).unwrap();
    let mut content = identity_content(&lowered());
    for (label, new) in EDITS {
        content.paragraphs[index(label)].runs[0].text = (*new).into();
    }
    session.save_edited_from_content(&content).unwrap()
}

#[test]
fn edited_runs_keep_their_tab_and_hyphen_elements_in_place__feat__plugin_doc_save_back() {
    let (saved, _) = edited_save();
    let xml = document_xml(&saved);
    let t = |s: &str| format!(r#"<w:t xml:space="preserve">{s}</w:t>"#);
    let has = |label: &str, want: &str| {
        let p = paragraph_xml(&xml, label);
        assert!(p.contains(want), "{label}: want {want}\n in {p}");
    };
    has("T01", &format!("{}<w:tab/>{}</w:r>", t("T01"), t("EDITED")));
    // The two original tabs, then a new one (Word's own element).
    has(
        "T02",
        &format!(
            "{}<w:tab/>{}<w:tab/>{}</w:r>",
            t("T02a"),
            t("T02b"),
            t("T02c")
        ),
    );
    // The leading tab stays before the text, after the run's rPr.
    has(
        "T03",
        &format!("</w:rPr><w:tab/>{}</w:r>", t("T03 leading EDITED")),
    );
    has(
        "N01",
        &format!(
            "{}<w:noBreakHyphen/>{}</w:r>",
            t("N01 well"),
            t("known EDITED")
        ),
    );
    has(
        "M01",
        &format!(
            "{}<w:tab/>{}<w:br w:clear=\"all\"/>{}<w:noBreakHyphen/>{}</w:r>",
            t("M01a"),
            t("B"),
            t("M01c"),
            t("D")
        ),
    );
    // A literal tab stays literal: the save does not change its form.
    has("X01", &t("X01\tLITERAL"));
    // Every edited paragraph is patched in place: its `<w:pPr>` (the tab
    // stops) keeps its bytes.
    let original = document_xml(&run_specials_docx());
    for (label, _) in EDITS {
        let ppr = |x: &str| {
            let p = paragraph_xml(x, label);
            p[..p.find("</w:pPr>").unwrap()].to_string()
        };
        assert_eq!(ppr(&xml), ppr(&original), "{label}'s pPr");
    }
    // No tab character was folded into any other `<w:t>`.
    for label in ["T01", "T02", "T03", "M01"] {
        let p = paragraph_xml(&xml, label);
        assert!(!p.contains('\t'), "{label}: a tab character in {p}");
    }
}

#[test]
fn edits_the_text_cannot_place_are_refused_and_keep_their_bytes__feat__plugin_doc_save_back() {
    let original = document_xml(&run_specials_docx());
    let (saved, skips) = edited_save();
    let xml = document_xml(&saved);
    assert_eq!(skips.len(), 3, "{skips:?}");
    for (label, element) in [("S01", "w:sym"), ("H01", "w:softHyphen"), ("P01", "w:ptab")] {
        let block = index(label);
        assert!(
            skips.iter().any(|s| s.starts_with(&format!(
                "run edit skipped: block {block} run 0: it holds a <{element}>"
            ))),
            "{label}: {skips:?}"
        );
        assert_eq!(
            paragraph_xml(&xml, label),
            paragraph_xml(&original, label),
            "{label} is byte-identical"
        );
    }
}

#[test]
fn an_edited_save_re_imports_to_the_edited_text__feat__plugin_doc_save_back() {
    let (saved, _) = edited_save();
    // `scripts/word-run-specials-probe.sh` has Word open this very file.
    if let Ok(path) = std::env::var("RUN_SPECIALS_EDITED_OUT") {
        std::fs::write(path, &saved).expect("write the edited save");
    }
    let ir = docx_lower::lower(&import_docx(&saved).unwrap());
    let paras = paragraphs(&ir);
    for (label, new) in EDITS {
        let got = text(paras[index(label)]);
        if ["S01", "H01", "P01"].contains(label) {
            assert_eq!(got, text(paragraphs(&lowered())[index(label)]), "{label}");
        } else {
            assert_eq!(got, *new, "{label}");
        }
    }
}

/// Word's answer for the edited save (`fixtures/run-specials.word.json`, via
/// `scripts/word-run-specials-probe.sh`): Word opened it without a repair
/// prompt, and every piece of text after a tab starts at the next tab stop
/// (72 pt / 144 pt from the 36 pt margin), new tabs included — the tabs are
/// tabs to Word, not text. (A literal tab inside `<w:t>`, X01, is one too.)
#[test]
fn word_puts_the_edited_text_after_each_tab_at_its_tab_stop__feat__plugin_doc_save_back() {
    let word: serde_json::Value =
        serde_json::from_str(include_str!("../fixtures/run-specials.word.json")).unwrap();
    let lines: Vec<Vec<(String, f64)>> = word["edited"]
        .as_array()
        .unwrap()
        .iter()
        .map(|l| {
            l.as_array()
                .unwrap()
                .iter()
                .map(|w| (w[0].as_str().unwrap().into(), w[1].as_f64().unwrap()))
                .collect()
        })
        .collect();
    let stops = [36.0, 108.0, 180.0];
    for label in ["T01", "T02", "T03", "M01", "X01"] {
        let new = EDITS.iter().find(|e| e.0 == label).unwrap().1;
        let first_line = new.split('\u{2028}').next().unwrap();
        let line = lines
            .iter()
            .find(|l| l.iter().any(|(w, _)| w.starts_with(label)))
            .unwrap_or_else(|| panic!("{label} in Word's map"));
        for (k, part) in first_line.split('\t').enumerate() {
            let Some(word) = part.split_whitespace().next() else {
                continue;
            };
            let (_, x) = line
                .iter()
                .find(|(w, _)| w == word)
                .unwrap_or_else(|| panic!("{label}: {word} in {line:?}"));
            assert!(
                (x - stops[k]).abs() < 0.5,
                "{label}: {word} at {x}, want {}",
                stops[k]
            );
        }
    }
}
