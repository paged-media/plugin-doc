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

//! ADR 029 — Word line spacing lowers to the leading WORD lays.
//! `line_spacing_docx()` is the document Word laid out
//! (`scripts/word-line-spacing-probe.sh`); its per-case line count and
//! baseline pitch are `fixtures/line-spacing.word.json`. Each case's lowered
//! `characterLeading` must equal Word's pitch within ±0.1 pt.

// `__feat__<id>` names link these tests to the Cockpit feature row.
#![allow(non_snake_case)]

use docx_conformance::{line_spacing_docx, LINE_SPACING_CASES};
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue};

/// One measured case from the Word fixture.
struct WordCase {
    case: String,
    rule: String,
    lines: u32,
    pitch_pt: f32,
}

/// The `cases` rows of `line-spacing.word.json` (one object per line; no
/// JSON dependency in this crate).
fn word_cases() -> Vec<WordCase> {
    let json = include_str!("../fixtures/line-spacing.word.json");
    let field = |row: &str, key: &str| -> String {
        let at = row.find(&format!("\"{key}\": ")).expect(key) + key.len() + 4;
        row[at..]
            .split([',', ' ', '}'])
            .next()
            .unwrap()
            .trim_matches('"')
            .to_string()
    };
    json.lines()
        .filter(|l| l.trim_start().starts_with("{ \"case\""))
        .map(|row| WordCase {
            case: field(row, "case"),
            rule: field(row, "rule"),
            lines: field(row, "lines").parse().unwrap(),
            pitch_pt: field(row, "pitch_pt").parse().unwrap(),
        })
        .collect()
}

/// The `characterLeading` a paragraph style resolves to through `basedOn`.
fn resolved_leading(doc: &LoweredDoc, style_id: &str) -> Option<f32> {
    let mut next = Some(style_id.to_string());
    for _ in 0..32 {
        let style = doc.styles.iter().find(|s| Some(&s.id) == next.as_ref())?;
        if let Some(PropValue::Length(pt)) = style
            .props
            .iter()
            .find(|p| p.path == "characterLeading")
            .map(|p| &p.value)
        {
            return Some(*pt);
        }
        next = style.based_on.clone();
    }
    None
}

/// The leading each case's paragraphs lower to, in case order.
fn lowered_leadings() -> Vec<f32> {
    let doc = docx_import::import_docx(&line_spacing_docx()).expect("import");
    let lowered = docx_lower::lower(&doc);
    let blocks = &lowered.story.blocks;
    LINE_SPACING_CASES
        .iter()
        .enumerate()
        .map(|(ci, case)| {
            // Every paragraph of a case lowers identically; check them all.
            let leadings: Vec<f32> = blocks[ci * 80..(ci + 1) * 80]
                .iter()
                .map(|b| match b {
                    LoweredBlock::Paragraph(p) => {
                        let id = p.para_style_id.as_deref().expect("styled paragraph");
                        resolved_leading(&lowered, id)
                            .unwrap_or_else(|| panic!("{}: no leading on {id}", case.label))
                    }
                    LoweredBlock::Table(_) => panic!("no tables in the fixture"),
                })
                .collect();
            assert!(
                leadings.windows(2).all(|w| w[0] == w[1]),
                "{}: one leading per case",
                case.label
            );
            leadings[0]
        })
        .collect()
}

#[test]
fn every_case_lowers_to_the_pitch_word_laid__feat__plugin_doc_read_path() {
    let word = word_cases();
    assert_eq!(
        word.len(),
        LINE_SPACING_CASES.len(),
        "one Word row per case"
    );
    let lowered = lowered_leadings();
    let mut misses = Vec::new();
    for ((case, w), got) in LINE_SPACING_CASES.iter().zip(&word).zip(&lowered) {
        assert_eq!(case.label, w.case, "fixture rows follow the case order");
        if (got - w.pitch_pt).abs() > 0.1 {
            misses.push(format!(
                "{} {} {}pt {} {}: lowered {got:.3} pt, Word laid {:.3} pt",
                case.label,
                case.font,
                case.half_pts as f32 / 2.0,
                case.rule,
                case.line,
                w.pitch_pt
            ));
        }
    }
    assert!(misses.is_empty(), "{misses:#?}");
}

/// A second, independent check on the same numbers: Word's LINE COUNT per
/// 648 pt page follows from the lowered pitch. Word fits `n` lines when
/// `(n − 1) × pitch + last ≤ 648`, where the last line needs only the face's
/// single line under `auto` (the extra multiple-spacing may run past the
/// bottom margin: double-spaced Calibri 10 pt fits 27 lines, not 26) and
/// the full pitch under `exact`/`atLeast`.
#[test]
fn the_lowered_pitch_reproduces_words_lines_per_page__feat__plugin_doc_read_path() {
    let word = word_cases();
    let lowered = lowered_leadings();
    let mut misses = Vec::new();
    for ((case, w), pitch) in LINE_SPACING_CASES.iter().zip(&word).zip(&lowered) {
        let last = if w.rule == "auto" {
            pitch * 240.0 / case.line as f32
        } else {
            *pitch
        };
        let lines = ((648.0 - last + 1e-3) / pitch).floor() as u32 + 1;
        if lines != w.lines {
            misses.push(format!(
                "{}: lowered pitch {pitch:.3} fits {lines} lines, Word laid {}",
                case.label, w.lines
            ));
        }
    }
    assert!(misses.is_empty(), "{misses:#?}");
}

/// A face Word has not been measured on is named, not silently guessed.
#[test]
fn an_unmeasured_face_is_named_in_a_diagnostic__feat__plugin_doc_read_path() {
    let doc = docx_import::import_docx(&line_spacing_docx()).expect("import");
    let lowered = docx_lower::lower(&doc);
    let diag = lowered
        .diagnostics
        .iter()
        .find(|d| d.message.contains("not been measured"))
        .expect("diagnostic for the unmeasured face");
    assert!(diag.message.contains("Inter"), "{}", diag.message);
}

/// Line spacing set on a STYLE (here docDefaults, as Word 2007+ writes it)
/// reaches every paragraph, at the paragraph's own resolved size.
#[test]
fn style_level_auto_spacing_follows_the_runs_size__feat__plugin_doc_read_path() {
    let styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Big"><w:name w:val="Big"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>
</w:styles>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>normal</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Big"/></w:pPr><w:r><w:t>big</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>direct 14</w:t></w:r></w:p>
</w:body></w:document>"#;
    let bytes = docx_conformance::zip_parts(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/_rels/document.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
        ),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", styles.as_bytes()),
    ]);
    let doc = docx_import::import_docx(&bytes).expect("import");
    let lowered = docx_lower::lower(&doc);
    let leading = |i: usize| match &lowered.story.blocks[i] {
        LoweredBlock::Paragraph(p) => {
            resolved_leading(&lowered, p.para_style_id.as_deref().unwrap()).expect("leading")
        }
        LoweredBlock::Table(_) => unreachable!(),
    };
    // Calibri single = 2500/2048 em; auto 259 = × 259/240.
    let expect = |size: f32| 2500.0 / 2048.0 * size * 259.0 / 240.0;
    for (i, size) in [(0, 11.0), (1, 20.0), (2, 14.0)] {
        assert!(
            (leading(i) - expect(size)).abs() < 0.01,
            "paragraph {i} at {size} pt: {} vs {}",
            leading(i),
            expect(size)
        );
    }
}
