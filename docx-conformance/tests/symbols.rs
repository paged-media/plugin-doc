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

//! Symbols, optional hyphens, absolute-position tabs and empty table cells
//! (`symbols_docx()`), against Word's answer (`fixtures/symbols.word.json`,
//! `scripts/word-symbols-probe.sh`).

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::{symbols_docx, SYMBOL_CASES};
use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};
use docx_import::import_docx;
use docx_js::DocSession;
use docx_lower::ir::{LoweredBlock, LoweredDoc, LoweredParagraph, LoweredTabStop, PropValue};
use paged_ooxml::OpcPackage;

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&symbols_docx()).expect("import"))
}

fn index(label: &str) -> usize {
    SYMBOL_CASES.iter().position(|c| c.0 == label).unwrap()
}

fn paragraph(ir: &LoweredDoc, label: &str) -> LoweredParagraph {
    match &ir.story.blocks[index(label)] {
        LoweredBlock::Paragraph(p) => p.clone(),
        LoweredBlock::Table(_) => panic!("{label} is a paragraph"),
    }
}

fn text(p: &LoweredParagraph) -> String {
    p.runs.iter().map(|r| r.text.as_str()).collect()
}

/// The read-back the editor would hand save-back: the baseline, as poured.
fn identity_content(ir: &LoweredDoc) -> StoryContentIn {
    StoryContentIn {
        self_id: "Story/doc".into(),
        paragraphs: ir
            .story
            .blocks
            .iter()
            .filter_map(|b| match b {
                LoweredBlock::Paragraph(p) => Some(p),
                LoweredBlock::Table(_) => None,
            })
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

/// The edits [`edited_save`] makes, by label: the new text of the
/// paragraph's one run.
const EDITS: &[(&str, &str)] = &[
    ("S01", "S01 a\u{2022}b \u{03B1}\u{0394}\u{03A9} EDITED"),
    (
        "S02",
        "S02 \u{2713} \u{2718} \u{25AA} \u{27A2} \u{263A} EDITED",
    ),
    // No place for the Windows logo in the edited text: refused.
    ("S03", "S03 xEDITEDy"),
    (
        "H01",
        "H01 Donau\u{00AD}dampf\u{00AD}schiff\u{00AD}fahrts\u{00AD}gesell\u{00AD}schaft EDITED",
    ),
    ("P01", "P01 left\tright EDITED"),
    ("P04", "P04\tdots EDITED"),
];

fn edited_save() -> (Vec<u8>, Vec<String>) {
    let session = DocSession::load(&symbols_docx()).unwrap();
    let ir = lowered();
    let mut content = identity_content(&ir);
    for (label, new) in EDITS {
        content.paragraphs[index(label)].runs[0].text = (*new).into();
    }
    session.save_edited_from_content(&content).unwrap()
}

#[test]
fn an_edited_save_re_imports_to_the_edited_text__feat__plugin_doc_save_back() {
    let (saved, _) = edited_save();
    // `scripts/word-symbols-probe.sh` has Word open this very file.
    if let Ok(path) = std::env::var("SYMBOLS_EDITED_OUT") {
        std::fs::write(path, &saved).expect("write the edited save");
    }
    let ir = docx_lower::lower(&import_docx(&saved).unwrap());
    for (label, new) in EDITS {
        let want = if *label == "S03" {
            text(&paragraph(&lowered(), label))
        } else {
            (*new).to_string()
        };
        assert_eq!(text(&paragraph(&ir, label)), want, "{label}");
    }
}

/// One line of Word's PDF: its baseline y and its words `(text, x0, x1)`.
struct Line {
    y: f64,
    words: Vec<(String, f64, f64)>,
}

fn word_lines(which: &str) -> Vec<Line> {
    let v: serde_json::Value =
        serde_json::from_str(include_str!("../fixtures/symbols.word.json")).unwrap();
    v[which]
        .as_array()
        .unwrap()
        .iter()
        .map(|l| Line {
            y: l["y"].as_f64().unwrap(),
            words: l["words"]
                .as_array()
                .unwrap()
                .iter()
                .map(|w| {
                    (
                        w[0].as_str().unwrap().to_string(),
                        w[1].as_f64().unwrap(),
                        w[2].as_f64().unwrap(),
                    )
                })
                .collect(),
        })
        .collect()
}

/// The lines from the one starting with `label` up to the next label's.
fn lines_of<'a>(lines: &'a [Line], label: &str) -> &'a [Line] {
    let starts = |l: &Line, s: &str| l.words.first().is_some_and(|w| w.0.starts_with(s));
    let first = lines
        .iter()
        .position(|l| starts(l, label))
        .unwrap_or_else(|| panic!("{label} in Word's answer"));
    let end = lines[first + 1..]
        .iter()
        .position(|l| SYMBOL_CASES.iter().any(|c| starts(l, c.0)) || starts(l, "C01"))
        .map_or(lines.len(), |k| first + 1 + k);
    &lines[first..end]
}

fn is_private(s: &str) -> bool {
    s.chars().any(|c| ('\u{E000}'..='\u{F8FF}').contains(&c))
}

/// Word draws a symbol-font character with the font's glyph and names it,
/// in its PDF, by a Unicode character where it knows one (the bullet, the
/// check mark, the small square, the arrowhead, the smiley) and by the
/// private-use code elsewhere. Where Word names one, the import carries the
/// same character; the glyph Word cannot name either (the Windows logo, its
/// `x\u{F0FF}y`) has no character here and a warning says so.
#[test]
fn symbols_carry_the_character_word_names_them_by__feat__plugin_doc_read_path() {
    let ir = lowered();
    let word = word_lines("original");
    let mut compared = 0;
    for label in ["S01", "S02"] {
        let ours: Vec<String> = text(&paragraph(&ir, label))
            .split_whitespace()
            .map(str::to_string)
            .collect();
        let line = &lines_of(&word, label)[0];
        assert_eq!(ours.len(), line.words.len(), "{label}: {ours:?}");
        for (o, (w, _, _)) in ours.iter().zip(&line.words) {
            if !is_private(w) {
                assert_eq!(o, w, "{label}");
                compared += 1;
            }
        }
    }
    // S01 a•b, S02 ✓ ▪ ➢ ☺ (plus the labels).
    assert!(compared >= 7, "{compared}");
    assert_eq!(
        text(&paragraph(&ir, "S01")),
        "S01 a\u{2022}b \u{03B1}\u{0394}\u{03A9}"
    );

    // The Windows logo: Word draws it; it has no Unicode equivalent.
    assert_eq!(lines_of(&word, "S03")[0].words[1].0, "x\u{F0FF}y");
    assert_eq!(text(&paragraph(&ir, "S03")), "S03 xy");
    let warned: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.severity == "warning" && d.message.contains("Wingdings F0FF"))
        .map(|d| d.message.as_str())
        .collect();
    assert_eq!(warned.len(), 1, "{:?}", ir.diagnostics);
    assert!(
        warned[0].contains(&format!("body block {}", index("S03"))),
        "{warned:?}"
    );
    assert!(
        ir.diagnostics.iter().any(|d| d.severity == "info"
            && d.message.contains("4 from Symbol")
            && d.message.contains("5 from Wingdings")),
        "{:?}",
        ir.diagnostics
    );
}

/// Word breaks a long word at its optional hyphens and draws a hyphen
/// there (`Donaudampf-` / `schifffahrtsgesell-` / `schaft` in a 90 pt
/// measure); where the line does not break, nothing shows (`unbroken`). The
/// import carries each one as U+00AD at its place, which the engine's
/// composer takes as the only break its word gets, drawing a hyphen; and
/// the lowering leaves hyphenation on, which the composer needs to see one.
///
/// Without optional hyphens Word does not hyphenate at all (automatic
/// hyphenation is off by default): it breaks the word where the line is
/// full, with no hyphen (`Donaudampfschifffa` / `hrtsgesellschaft`). The
/// engine hyphenates by its dictionary instead — not modelled here.
#[test]
fn soft_hyphens_are_where_word_breaks_the_word__feat__plugin_doc_read_path() {
    let ir = lowered();
    let word = word_lines("original");
    let h01 = text(&paragraph(&ir, "H01"));
    let pieces: Vec<&str> = h01.trim_start_matches("H01 ").split('\u{00AD}').collect();
    assert_eq!(
        pieces,
        ["Donau", "dampf", "schiff", "fahrts", "gesell", "schaft"]
    );
    // Every line Word ended inside the word ends with a drawn hyphen, at a
    // soft hyphen of ours: its text is a run of whole pieces.
    let lines: Vec<String> = lines_of(&word, "H01")
        .iter()
        .flat_map(|l| l.words.iter().map(|w| w.0.clone()))
        .filter(|w| w != "H01")
        .collect();
    assert_eq!(lines, ["Donaudampf-", "schifffahrtsgesell-", "schaft"]);
    let mut k = 0;
    for line in &lines {
        let mut joined = String::new();
        while joined.len() < line.trim_end_matches('-').len() {
            joined.push_str(pieces[k]);
            k += 1;
        }
        assert_eq!(
            joined,
            line.trim_end_matches('-'),
            "Word broke at one of ours"
        );
    }
    assert_eq!(k, pieces.len());

    // H02, no optional hyphen: Word's break draws no hyphen; nothing is
    // invented here.
    let h02: Vec<String> = lines_of(&word, "H02")
        .iter()
        .flat_map(|l| l.words.iter().map(|w| w.0.clone()))
        .collect();
    assert_eq!(h02, ["H02", "Donaudampfschifffa", "hrtsgesellschaft"]);
    assert!(!text(&paragraph(&ir, "H02")).contains('\u{00AD}'));

    // H03: not at a break, so invisible to Word, and carried all the same.
    assert_eq!(lines_of(&word, "H03")[0].words[1].0, "unbroken");
    assert_eq!(text(&paragraph(&ir, "H03")), "H03 un\u{00AD}broken");

    // Nothing turns the composer's hyphenation off.
    for st in &ir.styles {
        assert!(
            !st.props.iter().any(|p| p.path == "paragraphHyphenation"),
            "{}",
            st.id
        );
    }
}

/// The tab stops a paragraph's style resolves to (nearest first).
fn tab_stops(ir: &LoweredDoc, p: &LoweredParagraph) -> Vec<LoweredTabStop> {
    let mut id = p.para_style_id.clone();
    while let Some(current) = id {
        let st = ir.styles.iter().find(|s| s.id == current).unwrap();
        if let Some(PropValue::TabStops(stops)) = st
            .props
            .iter()
            .rev()
            .find(|p| p.path == "paragraphTabStops")
            .map(|p| &p.value)
        {
            return stops.clone();
        }
        id = st.based_on.clone();
    }
    Vec::new()
}

/// Word puts the text after an absolute-position tab at the margins' or
/// the indents' right / centre, whatever the tab stops: `right` ends at the
/// right margin (324), `centre` is centred on 180, `indent` ends at the
/// right indent (252), `dots` ends at 324 after a dot leader. Each is the
/// paragraph's only tab, so it lowers to the one tab stop that puts the
/// text there (positions from the 36 pt margin).
#[test]
fn absolute_position_tabs_lower_to_where_word_puts_the_text__feat__plugin_doc_read_path() {
    let ir = lowered();
    let word = word_lines("original");
    for (label, align, leader) in [
        ("P01", "right", None),
        ("P02", "center", None),
        ("P03", "right", None),
        ("P04", "right", Some(".")),
    ] {
        let p = paragraph(&ir, label);
        assert_eq!(text(&p).matches('\t').count(), 1, "{label}");
        let stops = tab_stops(&ir, &p);
        assert_eq!(stops.len(), 1, "{label}: {stops:?}");
        let stop = &stops[0];
        assert_eq!(stop.alignment.as_deref(), Some(align), "{label}");
        assert_eq!(stop.leader.as_deref(), leader, "{label}");
        let line = &lines_of(&word, label)[0];
        let (_, x0, x1) = line.words.last().unwrap();
        let at = match align {
            "center" => (x0 + x1) / 2.0,
            _ => *x1,
        };
        assert!(
            (36.0 + f64::from(stop.position) - at).abs() < 0.6,
            "{label}: stop at {} + 36, Word's text at {at}",
            stop.position
        );
        if let Some(l) = leader {
            assert!(line
                .words
                .iter()
                .any(|w| w.0.chars().all(|c| c.to_string() == l) && w.0.len() > 10));
        }
    }
    assert!(
        !ir.diagnostics
            .iter()
            .any(|d| d.message.contains("absolute-position")),
        "{:?}",
        ir.diagnostics
    );
}

/// Word lays an empty cell paragraph on its own line height: the cell whose
/// one empty paragraph has an exact 48 pt line makes the row 48 pt tall
/// (`AFTER` sits 48 pt + its own 12 pt below the line before the table).
/// So that cell's paragraph keeps its style in the lowering — one empty
/// paragraph, styled — which the pour seeds with an empty insertText and
/// styles with a cell caret (`doc-host-model`, "blank lines in table
/// cells"): `insertTable` mints cells with no paragraph at all.
#[test]
fn an_empty_cell_paragraph_keeps_the_style_word_lays_it_with__feat__plugin_doc_read_path() {
    let word = word_lines("original");
    let y = |w: &str| {
        word.iter()
            .find(|l| l.words.iter().any(|x| x.0 == w))
            .unwrap()
            .y
    };
    let row = y("AFTER") - y("P04") - 12.0;
    assert!((row - 48.0).abs() < 1.5, "Word's row is {row} pt");

    let ir = lowered();
    let table = ir
        .story
        .blocks
        .iter()
        .find_map(|b| match b {
            LoweredBlock::Table(t) => Some(t),
            _ => None,
        })
        .unwrap();
    let cell = table.cells.iter().find(|c| c.col == 1).unwrap();
    assert_eq!(cell.paragraphs.len(), 1);
    assert!(cell.paragraphs[0].runs.is_empty());
    let style = cell.paragraphs[0].para_style_id.clone().expect("styled");
    let leading = ir
        .styles
        .iter()
        .find(|s| s.id == style)
        .unwrap()
        .props
        .iter()
        .find(|p| p.path == "characterLeading")
        .map(|p| p.value.clone());
    assert_eq!(leading, Some(PropValue::Length(48.0)));
}

#[test]
fn zero_edit_save_with_symbols_is_byte_identical__feat__plugin_doc_save_back() {
    let original = symbols_docx();
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

/// An edited save writes every symbol, optional hyphen and ptab back as the
/// element it was, in place; a symbol with no character in the text (S03)
/// refuses its run's edit into the skip ledger, and that run keeps its
/// bytes. Word opened the edited save without a repair prompt and laid it
/// as before: the glyphs, the breaks at the optional hyphens, the
/// right-aligned text after the ptabs (now `right EDITED`, `dots EDITED`
/// ending at the margin).
#[test]
fn an_edited_save_keeps_the_elements_and_word_lays_it_as_before__feat__plugin_doc_save_back() {
    let (saved, skips) = edited_save();
    let xml = document_xml(&saved);
    let original = document_xml(&symbols_docx());
    assert_eq!(skips.len(), 1, "{skips:?}");
    assert!(
        skips[0].starts_with(&format!(
            "run edit skipped: block {} run 0: it holds a <w:sym> with no character",
            index("S03")
        )),
        "{skips:?}"
    );
    assert!(skips[0].contains(r#"w:char="F0FF""#), "{skips:?}");
    assert_eq!(paragraph_xml(&xml, "S03"), paragraph_xml(&original, "S03"));
    for (label, elements) in [
        ("S01", 4usize),
        ("S02", 5),
        ("H01", 5),
        ("P01", 1),
        ("P04", 1),
    ] {
        let (new, old) = (paragraph_xml(&xml, label), paragraph_xml(&original, label));
        let count = |x: &str| {
            x.matches("<w:sym ").count()
                + x.matches("<w:softHyphen/>").count()
                + x.matches("<w:ptab ").count()
        };
        assert_eq!(count(new), elements, "{label}: {new}");
        assert_eq!(count(old), elements, "{label}");
        assert!(new.contains("EDITED</w:t>"), "{label}: {new}");
        // No character the import read from an element is in the text.
        assert!(
            !new.contains(['\u{2022}', '\u{00AD}', '\t', '\u{2713}', '\u{03B1}']),
            "{label}: {new}"
        );
        // The elements keep their bytes and order.
        let elems = |x: &str| -> Vec<String> {
            x.split('<')
                .filter(|e| {
                    e.starts_with("w:sym ")
                        || e.starts_with("w:softHyphen")
                        || e.starts_with("w:ptab ")
                })
                .map(str::to_string)
                .collect()
        };
        assert_eq!(elems(new), elems(old), "{label}");
    }

    let before = word_lines("original");
    let after = word_lines("edited");
    for label in ["S01", "S02", "H01", "P01", "P04"] {
        let words = |l: &[Line]| -> Vec<String> {
            // A leader's length follows the text after it; the words do not.
            l.iter()
                .flat_map(|x| x.words.iter().map(|w| w.0.clone()))
                .filter(|w| !w.chars().all(|c| c == '.'))
                .collect()
        };
        let (b, a) = (
            words(lines_of(&before, label)),
            words(lines_of(&after, label)),
        );
        assert_eq!(
            a.last().map(String::as_str),
            Some("EDITED"),
            "{label}: {a:?}"
        );
        assert_eq!(
            &a[..a.len() - 1],
            &b[..],
            "{label}: Word laid the rest as before"
        );
    }
    for label in ["P01", "P04"] {
        let line = &lines_of(&after, label)[0];
        let (_, _, x1) = line.words.last().unwrap();
        assert!((x1 - 324.0).abs() < 0.6, "{label}: EDITED ends at {x1}");
    }
}
