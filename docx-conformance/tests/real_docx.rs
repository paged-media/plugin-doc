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

//! What the real corpus documents of the ADR 029 acceptance exposed
//! (`docs/acceptance-real-docx.md`), one construct per paragraph of
//! `real_docx_docx()`, against Word's answer (`fixtures/real-docx.word.json`,
//! `scripts/word-real-docx-probe.sh`). The text area is x = 36..324.

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::real_docx_docx;
use docx_import::import_docx;
use docx_lower::ir::{LoweredDoc, LoweredParagraph, LoweredTabStop, PropValue};

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&real_docx_docx()).expect("import"))
}

/// Word's word boxes: `(text, x0, y0, x1, y1)`.
fn word() -> Vec<(String, f64, f64, f64, f64)> {
    let v: serde_json::Value =
        serde_json::from_str(include_str!("../fixtures/real-docx.word.json")).unwrap();
    v["words"]
        .as_array()
        .unwrap()
        .iter()
        .map(|w| {
            let f = |k: &str| w[k].as_f64().unwrap();
            (
                w["t"].as_str().unwrap().to_string(),
                f("x0"),
                f("y0"),
                f("x1"),
                f("y1"),
            )
        })
        .collect()
}

fn word_box(label: &str) -> (String, f64, f64, f64, f64) {
    word().into_iter().find(|w| w.0 == label).unwrap()
}

fn paragraph(ir: &LoweredDoc, label: &str) -> LoweredParagraph {
    ir.story
        .paragraphs()
        .into_iter()
        .find(|p| p.runs.first().is_some_and(|r| r.text.starts_with(label)))
        .unwrap_or_else(|| panic!("{label}"))
        .clone()
}

/// The value a paragraph style resolves `path` to, walking its `basedOn`
/// chain, and the chain itself.
fn resolve(ir: &LoweredDoc, style: &str, path: &str) -> (Option<PropValue>, Vec<String>) {
    let mut chain = Vec::new();
    let mut next = Some(style.to_string());
    let mut found = None;
    while let Some(id) = next {
        let st = ir
            .styles
            .iter()
            .find(|s| s.id == id)
            .unwrap_or_else(|| panic!("{id}"));
        chain.push(id.clone());
        if found.is_none() {
            found = st
                .props
                .iter()
                .find(|p| p.path == path)
                .map(|p| p.value.clone());
        }
        next = st.based_on.clone();
    }
    (found, chain)
}

fn length(v: Option<PropValue>) -> f32 {
    match v {
        Some(PropValue::Length(n)) => n,
        other => panic!("{other:?}"),
    }
}

/// T01: Word MERGES the paragraph's own stop (left, 1 in) with its style's
/// (left 0.5 in; right with a dot leader at 300 pt, past the 288 pt margin),
/// and sets the page number after the dots ON the entry's line, ending at
/// the 300 pt stop out in the margin. The native tab list replaces its
/// style's, so the paragraph carries the merged set; the stops carry IDML's
/// alignment names (the engine's composer reads no other, a Word name laid
/// every right tab as a left one); and the right stop moves to the margin,
/// where a native frame can set the number on Word's line.
#[test]
fn a_toc_entry_keeps_its_merged_stops_and_its_page_number_on_the_line__feat__plugin_doc_read_path()
{
    let (_, _, t01_y, _, _) = word_box("T01");
    let (_, an_x, _, _, _) = word_box("An");
    let (_, _, seven_y, seven_x1, _) = word_box("7");
    assert!(
        (seven_y - t01_y).abs() < 0.01,
        "Word keeps the page number on the line"
    );
    assert!(
        (seven_x1 - (36.0 + 300.0)).abs() < 0.6,
        "Word ends it at the stop: {seven_x1}"
    );
    assert!((an_x - (36.0 + 36.0)).abs() < 0.6, "the style's left stop");
    assert!(word()
        .iter()
        .any(|w| w.0.chars().all(|c| c == '.') && w.0.len() > 20));

    let ir = lowered();
    let p = paragraph(&ir, "T01");
    let (stops, _) = resolve(
        &ir,
        p.para_style_id.as_deref().unwrap(),
        "paragraphTabStops",
    );
    let Some(PropValue::TabStops(stops)) = stops else {
        panic!("{stops:?}")
    };
    let stop = |position: f32, alignment: &str, leader: Option<&str>| LoweredTabStop {
        position,
        alignment: Some(alignment.into()),
        alignment_character: None,
        leader: leader.map(Into::into),
    };
    assert_eq!(
        stops,
        [
            stop(36.0, "LeftAlign", None),
            stop(72.0, "LeftAlign", None),
            stop(288.0, "RightAlign", Some(".")),
        ]
    );
    assert!(ir.diagnostics.iter().any(|d| d
        .message
        .contains("past the right margin moved to the margin")));
}

/// B01: a paragraph style based on a CHARACTER style. Word ignores the
/// cross-type `basedOn`: the paragraph lays out on the document defaults,
/// which set no size, so at Word's 10 pt (its glyphs 10/12 of N01's), and it
/// is bulleted through its style's own `w:numPr`, at the level's indents
/// (bullet at 1 in − 0.25 in, text at 1 in). N01 names no style: Word lays
/// it in the default paragraph style, Normal (12 pt), not on the bare
/// defaults.
#[test]
fn style_bases_lists_and_the_default_style_resolve_as_word_lays_them__feat__plugin_doc_read_path() {
    let (_, bullet_x, _, _, _) = word_box("\u{2022}");
    let (_, b01_x, b01_y0, _, b01_y1) = word_box("B01");
    let (_, _, n01_y0, _, n01_y1) = word_box("N01");
    assert!((bullet_x - 90.0).abs() < 0.6, "{bullet_x}");
    assert!((b01_x - 108.0).abs() < 0.6, "{b01_x}");
    let ratio = (b01_y1 - b01_y0) / (n01_y1 - n01_y0);
    assert!((ratio - 10.0 / 12.0).abs() < 0.02, "{ratio}");

    let ir = lowered();
    let b01 = paragraph(&ir, "B01");
    let style = b01.para_style_id.as_deref().unwrap();
    let (size, chain) = resolve(&ir, style, "characterFontSize");
    assert!(!chain.iter().any(|s| s.contains("FootRef")), "{chain:?}");
    assert!(
        chain.iter().any(|s| s.ends_with("docx-Default")),
        "{chain:?}"
    );
    assert_eq!(length(size), 10.0);
    let (list, _) = resolve(&ir, style, "paragraphListType");
    assert_eq!(list, Some(PropValue::Text("BulletList".into())));
    let left = length(resolve(&ir, style, "paragraphLeftIndent").0);
    let first = length(resolve(&ir, style, "paragraphFirstLineIndent").0);
    assert!((36.0 + left - b01_x as f32).abs() < 0.6, "{left}");
    assert!(
        (36.0 + left + first - bullet_x as f32).abs() < 0.6,
        "{first}"
    );

    let n01 = paragraph(&ir, "N01");
    let (size, chain) = resolve(
        &ir,
        n01.para_style_id.as_deref().unwrap(),
        "characterFontSize",
    );
    assert!(
        chain.iter().any(|s| s.ends_with("docx-Normal")),
        "{chain:?}"
    );
    assert_eq!(length(size), 12.0);
}

/// V01 / F01: legacy VML pictures take their room in Word's flow — the
/// inline one 100 pt (its line's text sits on its foot: V01 ends 100 pt +
/// Normal's 12 pt after N01), the floating one with a top-and-bottom wrap a
/// 50 pt band of its own (F01's text resumes below it). Both lower as
/// pictures of their VML size.
#[test]
fn vml_pictures_lower_at_the_size_word_gives_them__feat__plugin_doc_read_path() {
    let (_, _, _, _, n01_y1) = word_box("N01");
    let (_, _, v01_y0, _, v01_y1) = word_box("V01");
    let (_, _, a01_y0, _, _) = word_box("A01");
    let (_, _, f01_y0, _, _) = word_box("F01");
    assert!(
        (v01_y1 - n01_y1 - (12.0 + 100.0)).abs() < 3.0,
        "{v01_y1} {n01_y1}"
    );
    let pitch = a01_y0 - v01_y0;
    assert!(
        (f01_y0 - a01_y0 - (pitch + 50.0)).abs() < 1.0,
        "{f01_y0} {a01_y0}"
    );

    let ir = lowered();
    for (label, w, h) in [("V01", 100.0, 100.0), ("F01", 100.0, 50.0)] {
        let p = paragraph(&ir, label);
        assert_eq!(p.images.len(), 1, "{label}");
        assert_eq!(
            (p.images[0].width_pt, p.images[0].height_pt),
            (w, h),
            "{label}"
        );
        assert!(
            p.images[0].uri.starts_with("data:image/png;base64,"),
            "{label}"
        );
    }
    assert!(
        !ir.diagnostics.iter().any(|d| d.message.contains("VML")),
        "{:?}",
        ir.diagnostics
    );
}

/// Word hyphenates automatically only when the settings say so (they do
/// not here), so the document defaults turn the engine's hyphenation off.
#[test]
fn hyphenation_is_off_unless_the_document_turns_it_on__feat__plugin_doc_read_path() {
    let ir = lowered();
    let n01 = paragraph(&ir, "N01");
    let (h, _) = resolve(
        &ir,
        n01.para_style_id.as_deref().unwrap(),
        "paragraphHyphenation",
    );
    assert_eq!(h, Some(PropValue::Bool(false)));
}
