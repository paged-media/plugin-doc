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

//! Floating drawings (`wp:anchor`, thoughts ADR 035, RFI DOC-07) from
//! `floats_docx()`: their position and wrap are read and carried in the
//! lowered IR, they stay placed inline, and every one is named in a
//! diagnostic (ADR 035 work item 0) rather than silently becoming an
//! inline picture.

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::{floats_docx, image_docx, FLOAT_CASES};
use docx_core::{Block, FloatWrap, Paragraph};
use docx_import::import_docx;
use docx_lower::ir::{LoweredDoc, LoweredFloatPosition};

fn para(doc: &docx_core::DocxDocument, i: usize) -> &Paragraph {
    match &doc.body[i] {
        Block::Paragraph(p) => p,
        Block::Table(_) => panic!("block {i} is a table"),
    }
}

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&floats_docx()).unwrap())
}

#[test]
fn a_floats_position_and_wrap_are_read__feat__plugin_doc_read_path() {
    let doc = import_docx(&floats_docx()).unwrap();
    let float = |i: usize| {
        para(&doc, i).runs[0].images[0]
            .float
            .clone()
            .expect("a float")
    };

    let f1 = float(0);
    assert_eq!(f1.wrap, FloatWrap::Square);
    assert_eq!(f1.wrap_text.as_deref(), Some("bothSides"));
    let h = f1.horizontal.unwrap();
    assert_eq!(
        (h.relative_from.as_str(), h.offset),
        ("column", Some(152400))
    );
    let v = f1.vertical.unwrap();
    assert_eq!((v.relative_from.as_str(), v.offset), ("paragraph", Some(0)));
    assert_eq!((f1.dist_left, f1.dist_right), (114300, 114300));
    assert!(f1.allow_overlap && f1.layout_in_cell && !f1.behind_doc && !f1.locked);
    assert_eq!(f1.relative_height, 251659264);

    let f2 = float(1);
    assert_eq!(f2.wrap, FloatWrap::TopAndBottom);
    let h = f2.horizontal.unwrap();
    assert_eq!(
        (h.relative_from.as_str(), h.align.as_deref()),
        ("margin", Some("center"))
    );
    assert_eq!(f2.vertical.unwrap().offset, Some(914400));
    assert_eq!((f2.dist_top, f2.dist_bottom), (25400, 50800));
    assert!(f2.locked && !f2.allow_overlap);

    let f3 = float(2);
    assert_eq!(f3.wrap, FloatWrap::None);
    assert!(f3.behind_doc);
    assert_eq!(f3.vertical.unwrap().align.as_deref(), Some("top"));

    let f4 = float(3);
    assert_eq!(f4.wrap, FloatWrap::Tight);
    assert_eq!(f4.wrap_text.as_deref(), Some("right"));
    assert_eq!(f4.simple_pos, Some((1270000, 2540000)));
    assert_eq!(
        (f4.dist_left, f4.dist_right),
        (38100, 38100),
        "the wrap's own distances"
    );

    // An inline picture has no float.
    let inline = import_docx(&image_docx()).unwrap();
    assert!(para(&inline, 1).runs[0].images[0].float.is_none());
}

#[test]
fn every_drawing_in_a_run_is_kept_and_a_non_picture_is_counted__feat__plugin_doc_read_path() {
    let doc = import_docx(&floats_docx()).unwrap();
    let n = FLOAT_CASES.len();
    let two = &para(&doc, n).runs[0];
    assert_eq!(two.images.len(), 2, "the inline picture AND the float");
    assert!(two.images[0].float.is_none() && two.images[1].float.is_some());
    let chart = &para(&doc, n + 2).runs[0];
    assert!(chart.images.is_empty());
    assert_eq!(chart.other_drawings, 1);
}

#[test]
fn the_lowered_ir_carries_the_float_in_points_and_places_it_inline__feat__plugin_doc_read_path() {
    let ir = lowered();
    let p = ir.story.paragraphs();
    let img = &p[0].images[0];
    assert_eq!((img.width_pt, img.height_pt), (72.0, 54.0));
    let f = img.float.as_ref().expect("the float is carried");
    assert_eq!(f.wrap, "wrapSquare");
    assert_eq!(
        f.horizontal,
        Some(LoweredFloatPosition {
            relative_from: "column".into(),
            offset_pt: Some(12.0),
            align: None,
            percent: None,
        })
    );
    assert_eq!((f.dist_left_pt, f.dist_right_pt), (9.0, 9.0));
    assert_eq!(
        p[3].images[0].float.as_ref().unwrap().simple_pos_pt,
        Some((100.0, 200.0))
    );
    // Both drawings of the two-drawing run reach the paragraph.
    assert_eq!(p[FLOAT_CASES.len()].images.len(), 2);
    // The JSON the bundle reads: an inline picture carries no `float` key.
    let json = serde_json::to_value(&ir).unwrap();
    let images = &json["story"]["blocks"][FLOAT_CASES.len()]["images"];
    assert!(images[0].get("float").is_none(), "{images}");
    assert_eq!(images[1]["float"]["wrapText"], "bothSides");
    assert_eq!(images[1]["float"]["horizontal"]["offsetPt"], 12.0);
}

#[test]
fn every_floating_picture_is_named_in_a_diagnostic__feat__plugin_doc_read_path() {
    let ir = lowered();
    let floats: Vec<&str> = ir
        .diagnostics
        .iter()
        .filter(|d| d.message.contains("floating picture"))
        .map(|d| d.message.as_str())
        .collect();
    // F1..F4 plus the float in the two-drawing run; the cell's float is
    // reported with the cell pictures.
    assert_eq!(floats.len(), FLOAT_CASES.len() + 1, "{floats:#?}");
    for m in &floats {
        assert!(m.contains("placed INLINE"), "{m}");
        assert!(m.contains("not carried onto the page yet"), "{m}");
    }
    assert!(floats[0].contains("body block 0"), "{}", floats[0]);
    assert!(floats[0].contains("wrapSquare bothSides"), "{}", floats[0]);
    assert!(
        floats[0].contains("horizontally column +12.0 pt"),
        "{}",
        floats[0]
    );
    assert!(floats[1].contains("wrapTopAndBottom"), "{}", floats[1]);
    assert!(
        floats[1].contains("horizontally margin center"),
        "{}",
        floats[1]
    );
    assert!(floats[2].contains("behind the text"), "{}", floats[2]);
    assert!(floats[3].contains("at (100.0, 200.0) pt"), "{}", floats[3]);
    assert!(floats.iter().all(|m| ir
        .diagnostics
        .iter()
        .any(|d| d.message == *m && d.severity == "warning")));

    let msgs: Vec<&str> = ir.diagnostics.iter().map(|d| d.message.as_str()).collect();
    let cell = msgs
        .iter()
        .find(|m| m.contains("table cells"))
        .expect("the cell-picture diagnostic");
    assert!(
        cell.contains("1 picture(s) in table cells (1 of them floating) are NOT shown"),
        "{cell}"
    );
    let chart = msgs
        .iter()
        .find(|m| m.contains("not pictures"))
        .expect("the non-picture diagnostic");
    assert!(
        chart.contains(&format!("body block {}", FLOAT_CASES.len() + 2)),
        "{chart}"
    );

    // An inline picture is not diagnosed as a float.
    let inline = docx_lower::lower(&import_docx(&image_docx()).unwrap());
    assert!(!inline
        .diagnostics
        .iter()
        .any(|d| d.message.contains("floating")));
}
