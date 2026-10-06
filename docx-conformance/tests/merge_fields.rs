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

//! Word mail-merge fields (`MERGEFIELD`): parsed in both field forms, carried
//! through the lowering for the object model (`mergeField`), and written back
//! by save-back as the field's RESULT only — the field code survives, so Word
//! can still merge the document.

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::merge_field_docx;
use docx_core::{Block, MergeField};
use docx_export::{EditSet, RunEdit};
use docx_import::{import_docx, parse_merge_field};
use docx_js::DocSession;

fn para(doc: &docx_core::DocxDocument, i: usize) -> &docx_core::Paragraph {
    match &doc.body[i] {
        Block::Paragraph(p) => p,
        _ => panic!("body[{i}] is not a paragraph"),
    }
}

#[test]
fn the_instruction_parser_reads_names_and_switches__feat__plugin_doc_object_model() {
    let f = parse_merge_field(
        r#" MERGEFIELD "First Name" \b "Mx " \f "!" \* Upper \* MERGEFORMAT \m "#,
        2,
    )
    .unwrap();
    assert_eq!(f.name, "First Name");
    assert_eq!(f.before.as_deref(), Some("Mx "));
    assert_eq!(f.after.as_deref(), Some("!"));
    assert_eq!(f.format, ["Upper", "MERGEFORMAT"]);
    assert!(f.mapped && !f.vertical);
    assert_eq!(f.occurrence, 2);
    assert_eq!(
        parse_merge_field("mergefield City", 0).unwrap().name,
        "City"
    );
    assert_eq!(parse_merge_field(" PAGE ", 0), None);
    assert_eq!(parse_merge_field(" MERGEFIELD ", 0), None);
    assert_eq!(parse_merge_field(" MERGEFIELD \\* MERGEFORMAT ", 0), None);
}

#[test]
fn both_field_forms_stamp_their_result_runs__feat__plugin_doc_object_model() {
    let doc = import_docx(&merge_field_docx()).unwrap();
    let p = para(&doc, 0);
    let names: Vec<(String, Option<(String, u32)>)> = p
        .runs
        .iter()
        .map(|r| {
            (
                r.text.clone(),
                r.merge_field
                    .as_ref()
                    .map(|m| (m.name.clone(), m.occurrence)),
            )
        })
        .collect();
    let mf = |n: &str, o: u32| Some((n.to_string(), o));
    assert_eq!(
        names,
        [
            ("Dear ".into(), None),
            ("«First".into(), mf("First Name", 0)),
            (" Name»".into(), mf("First Name", 0)),
            (", from ".into(), None),
            ("«City»".into(), mf("City", 1)),
            (" and ".into(), None),
            ("«City»".into(), mf("City", 2)),
            (", page ".into(), None),
            ("1".into(), None),
        ]
    );
    // The split instruction is reassembled before it is parsed.
    let first = p.runs[1].merge_field.as_ref().unwrap();
    assert_eq!(first.before.as_deref(), Some("Mx "));
    assert_eq!(first.format, ["MERGEFORMAT"]);
    assert_eq!(p.runs[4].merge_field.as_ref().unwrap().format, ["Upper"]);
    // The PAGE field keeps its frozen-field name and is no merge field.
    assert_eq!(p.runs[8].field.as_deref(), Some("PAGE"));
    let second = para(&doc, 1).runs[1].merge_field.clone().unwrap();
    assert_eq!(
        second,
        MergeField {
            name: "OrderNo".into(),
            instruction: r#" MERGEFIELD OrderNo \f "." \m "#.into(),
            occurrence: 0,
            after: Some(".".into()),
            mapped: true,
            ..Default::default()
        }
    );
}

#[test]
fn the_lowering_carries_the_field_on_its_runs__feat__plugin_doc_object_model() {
    let doc = import_docx(&merge_field_docx()).unwrap();
    let ir = docx_lower::lower(&doc);
    let runs = &ir.story.paragraphs()[0].runs;
    let fields: Vec<Option<&str>> = runs
        .iter()
        .map(|r| r.merge_field.as_ref().map(|m| m.name.as_str()))
        .collect();
    assert_eq!(
        fields,
        [
            None,
            Some("First Name"),
            Some("First Name"),
            None,
            Some("City"),
            None,
            Some("City"),
            None,
            None
        ]
    );
    let json = serde_json::to_string(&ir).unwrap();
    assert!(json.contains(r#""mergeField":{"name":"City""#), "{json}");
}

#[test]
fn save_back_writes_the_result_and_keeps_the_field_code__feat__plugin_doc_object_model() {
    let original = merge_field_docx();
    let session = DocSession::load(&original).unwrap();
    // A field's value is its first result run's text; the other result runs
    // are emptied (what the bundle's mergeField.value write sends).
    let edits = EditSet {
        runs: vec![
            RunEdit::text(0, 1, "Ada Lovelace"),
            RunEdit::text(0, 2, ""),
            RunEdit::text(0, 4, "London"),
            RunEdit::text(1, 1, "A-17"),
        ],
        ..Default::default()
    };
    let (saved, skips) = session.save_edited(&edits).unwrap();
    assert!(skips.is_empty(), "{skips:?}");
    let re = import_docx(&saved).unwrap();
    let p = para(&re, 0);
    let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
    assert_eq!(text, "Dear Ada Lovelace, from London and «City», page 1");
    let first = p.runs.iter().find(|r| r.text == "Ada Lovelace").unwrap();
    assert_eq!(
        first.merge_field.as_ref().unwrap().name,
        "First Name",
        "the complex field code survived"
    );
    let city = p.runs.iter().find(|r| r.text == "London").unwrap();
    assert_eq!(
        city.merge_field.as_ref().unwrap().format,
        ["Upper"],
        "the fldSimple instruction survived"
    );
    assert_eq!(para(&re, 1).runs[1].text, "A-17");
    assert_eq!(
        para(&re, 1).runs[1].merge_field.as_ref().unwrap().name,
        "OrderNo"
    );
    // Reloading the saved package lowers the new values (the bundle's cache).
    let again = DocSession::load(&saved).unwrap();
    let lowered = again.lowered_json();
    assert!(lowered.contains("Ada Lovelace") && lowered.contains("A-17"));
}

#[test]
fn object_edits_address_a_merge_field_by_occurrence__feat__plugin_doc_object_model() {
    use docx_js::core::{MergeFieldEdit, ObjectEdits, RunTextEdit};
    let session = DocSession::load(&merge_field_docx()).unwrap();
    let edits = ObjectEdits {
        runs: vec![RunTextEdit {
            block: 0,
            run: 0,
            text: "Hello ".into(),
        }],
        merge_fields: vec![
            MergeFieldEdit {
                block: 0,
                occurrence: 0,
                value: "Grace Hopper".into(),
            },
            MergeFieldEdit {
                block: 0,
                occurrence: 2,
                value: "Arlington".into(),
            },
        ],
    };
    // The set the session derives: the field's first result run takes the
    // value, the second is emptied.
    let set = session.object_edits(&edits).unwrap();
    let triples: Vec<(usize, usize, Option<&str>)> = set
        .runs
        .iter()
        .map(|r| (r.block, r.run, r.new_text.as_deref()))
        .collect();
    assert_eq!(
        triples,
        [
            (0, 0, Some("Hello ")),
            (0, 1, Some("Grace Hopper")),
            (0, 2, Some("")),
            (0, 6, Some("Arlington"))
        ]
    );
    let (saved, skips) = session.save_object_edits(&edits).unwrap();
    assert!(skips.is_empty(), "{skips:?}");
    let p = import_docx(&saved).unwrap();
    let text: String = para(&p, 0).runs.iter().map(|r| r.text.as_str()).collect();
    assert_eq!(
        text,
        "Hello Grace Hopper, from «City» and Arlington, page 1"
    );
    // Refusals: no such field, a run past the end.
    let none = ObjectEdits {
        merge_fields: vec![MergeFieldEdit {
            block: 1,
            occurrence: 3,
            value: "x".into(),
        }],
        ..Default::default()
    };
    assert!(session
        .object_edits(&none)
        .unwrap_err()
        .contains("no merge field #3"));
    let past = ObjectEdits {
        runs: vec![RunTextEdit {
            block: 1,
            run: 9,
            text: "x".into(),
        }],
        ..Default::default()
    };
    assert!(session.object_edits(&past).is_err());
}
