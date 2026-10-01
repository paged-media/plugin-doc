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

//! Footnote numbering (`w:footnotePr`, thoughts ADR 034, RFI DOC-06) from
//! `footnote_numbering_docx()`, against Word's answer
//! (`fixtures/footnote-numbering.word.json`,
//! `scripts/word-footnotes-probe.sh`): each note's mark as Word prints it.

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::{footnote_numbering_docx, zip_parts, FOOTNOTE_NUMBERING_SECTIONS};
use docx_core::{format_note_number, DocxDocument, NoteRestart};
use docx_import::import_docx;
use paged_ooxml::OpcPackage;
use serde_json::Value;

fn word() -> Value {
    serde_json::from_str(include_str!("../fixtures/footnote-numbering.word.json"))
        .expect("footnote-numbering.word.json")
}

/// The fixture with `document.xml` rewritten, as the probe makes its copies.
fn variant(edit: impl Fn(String) -> String) -> Vec<u8> {
    let pkg = OpcPackage::read(&footnote_numbering_docx()).unwrap();
    let names: Vec<String> = pkg.file_names().map(str::to_string).collect();
    let parts: Vec<(String, Vec<u8>)> = names
        .iter()
        .map(|n| {
            let b = pkg.part(n).unwrap().to_vec();
            let b = if n == "word/document.xml" {
                edit(String::from_utf8(b).unwrap()).into_bytes()
            } else {
                b
            };
            (n.clone(), b)
        })
        .collect();
    let refs: Vec<(&str, &[u8])> = parts
        .iter()
        .map(|(n, b)| (n.as_str(), b.as_slice()))
        .collect();
    zip_parts(&refs)
}

/// Every note's mark as the model predicts it, in document order: the
/// section's own numbering, the note's ordinals in the document, section
/// and page (pages from FOOTNOTE_NUMBERING_SECTIONS).
fn predicted(doc: &DocxDocument) -> Vec<(String, String)> {
    let mut out = Vec::new();
    let mut in_document = 0;
    for (k, (label, _, pages, per_page)) in FOOTNOTE_NUMBERING_SECTIONS.iter().enumerate() {
        let numbering = doc.footnote_numbering(k);
        let mut in_section = 0;
        for pg in 1..=*pages {
            for n in 0..*per_page {
                let number = numbering.note_number(in_document, in_section, n);
                let mark = format_note_number(number, numbering.num_fmt.as_deref())
                    .expect("a format the fixture uses");
                out.push((format!("{label}-p{pg}-n{}", n + 1), mark));
                in_document += 1;
                in_section += 1;
            }
        }
    }
    out
}

fn assert_matches_word(doc: &DocxDocument, key: &str) {
    let word: Vec<(String, String)> = word()[key]
        .as_array()
        .unwrap()
        .iter()
        .map(|m| {
            (
                m["note"].as_str().unwrap().to_string(),
                m["mark"].as_str().unwrap().to_string(),
            )
        })
        .collect();
    assert_eq!(predicted(doc), word, "{key}");
}

#[test]
fn footnote_numbering_is_parsed_per_section_and_from_settings__feat__plugin_doc_read_path() {
    let doc = import_docx(&footnote_numbering_docx()).unwrap();
    assert_eq!(doc.footnote_props.num_fmt.as_deref(), Some("lowerRoman"));
    let s = &doc.sections;
    assert_eq!(s.len(), 4);
    assert_eq!(s[0].footnote_props.num_start, Some(3));
    assert_eq!(s[0].footnote_props.num_fmt, None);
    assert_eq!(s[1].footnote_props.num_fmt.as_deref(), Some("upperLetter"));
    assert_eq!(
        s[1].footnote_props.num_restart,
        Some(NoteRestart::EachSection)
    );
    assert!(s[2].footnote_props.is_empty());
    assert_eq!(s[3].footnote_props.num_restart, Some(NoteRestart::EachPage));
    // Word numbers by the section's own settings alone.
    assert_eq!(doc.footnote_numbering(0).num_fmt, None);
    assert!(doc.footnote_numbering(2).is_empty());
}

#[test]
fn every_footnote_mark_is_words__feat__plugin_doc_read_path() {
    assert_matches_word(&import_docx(&footnote_numbering_docx()).unwrap(), "marks");
    let start10 =
        variant(|d| d.replace(r#"<w:numStart w:val="3"/>"#, r#"<w:numStart w:val="10"/>"#));
    assert_matches_word(&import_docx(&start10).unwrap(), "start10");
    let none = variant(|d| {
        let mut out = String::new();
        let mut rest = d.as_str();
        while let Some(i) = rest.find("<w:footnotePr>") {
            out.push_str(&rest[..i]);
            let j = rest[i..].find("</w:footnotePr>").unwrap() + i + "</w:footnotePr>".len();
            rest = &rest[j..];
        }
        out.push_str(rest);
        out
    });
    assert_matches_word(&import_docx(&none).unwrap(), "no_section_props");
}

#[test]
fn the_footnote_diagnostic_says_what_is_carried_and_what_is_not__feat__plugin_doc_read_path() {
    let ir = docx_lower::lower(&import_docx(&footnote_numbering_docx()).unwrap());
    let n: Vec<_> = ir
        .sections
        .iter()
        .map(|s| s.footnote_numbering.clone())
        .collect();
    assert_eq!(n[0].as_ref().unwrap().num_start, Some(3));
    assert_eq!(
        n[1].as_ref().unwrap().num_restart.as_deref(),
        Some("eachSect")
    );
    assert_eq!(
        n[1].as_ref().unwrap().num_fmt.as_deref(),
        Some("upperLetter")
    );
    assert_eq!(n[2], None, "S3 says nothing: Word's defaults");
    assert_eq!(
        n[3].as_ref().unwrap().num_restart.as_deref(),
        Some("eachPage")
    );
    let msg = ir
        .diagnostics
        .iter()
        .map(|d| d.message.as_str())
        .find(|m| m.contains("footnote(s)"))
        .expect("a footnote diagnostic");
    assert!(msg.contains("10 footnote(s)"), "{msg}");
    assert!(msg.contains("differing"), "{msg}");
    assert!(msg.contains("NOT placed on the page"), "{msg}");
    assert!(msg.contains("no create door"), "{msg}");
    assert!(
        !msg.contains("no footnote construct"),
        "the stale claim is gone: {msg}"
    );
}

#[test]
fn note_numbers_format_as_word_writes_them__feat__plugin_doc_read_path() {
    let f = |n, fmt| format_note_number(n, Some(fmt)).unwrap();
    assert_eq!(f(4, "lowerRoman"), "iv");
    assert_eq!(f(1994, "upperRoman"), "MCMXCIV");
    assert_eq!(f(2, "upperLetter"), "B");
    assert_eq!(f(28, "lowerLetter"), "bb");
    assert_eq!(format_note_number(7, None).unwrap(), "7");
    assert_eq!(format_note_number(7, Some("chicago")), None);
}
