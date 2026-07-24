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

//! End-to-end Tier-0 conformance: a real (in-memory) `.docx` through
//! import → lower, plus the byte-identical zero-edit round-trip that proves the
//! preservation invariant.

use docx_conformance::{memo_docx, one_paragraph_docx, zip_parts};
use docx_core::{Block, StyleKind};
use docx_import::import_docx;
use docx_js::DocSession;
use docx_lower::ir::{PropValue, StyleCollection};
use docx_lower::lower;
use paged_ooxml::OpcPackage;

#[test]
fn opc_zero_edit_round_trip_is_byte_identical_per_part() {
    let bytes = memo_docx();
    let pkg = OpcPackage::read(&bytes).unwrap();
    let rewritten = pkg.write().unwrap();
    let pkg2 = OpcPackage::read(&rewritten).unwrap();

    let names: Vec<&str> = pkg.file_names().collect();
    let names2: Vec<&str> = pkg2.file_names().collect();
    assert_eq!(names, names2, "part set + order preserved");

    for name in names {
        assert_eq!(
            pkg.part(name),
            pkg2.part(name),
            "part {name} must round-trip byte-identical (decompressed)"
        );
    }
    // The unknown part specifically survives (preservation invariant).
    assert_eq!(
        pkg2.part("customXml/unknown.txt").unwrap(),
        b"paged preserves unknown parts"
    );
}

#[test]
fn imports_body_styles_and_section() {
    let doc = import_docx(&memo_docx()).unwrap();

    // 3 paragraphs (the sectPr is not a block).
    let paragraphs: Vec<_> = doc
        .body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    assert_eq!(paragraphs.len(), 3);

    // The heading paragraph references the Heading1 style.
    assert_eq!(paragraphs[1].style_id.as_deref(), Some("Heading1"));
    assert_eq!(paragraphs[1].runs[0].text, "A Centered Heading");

    // The mixed paragraph has a bold+red middle run.
    let mixed = &paragraphs[2];
    assert_eq!(mixed.runs.len(), 3);
    assert_eq!(mixed.runs[1].props.bold, Some(true));
    assert_eq!(mixed.runs[1].props.color.as_deref(), Some("FF0000"));

    // The style catalog resolved Normal + Heading1.
    let heading = doc
        .styles
        .styles
        .iter()
        .find(|s| s.style_id == "Heading1")
        .unwrap();
    assert_eq!(heading.kind, StyleKind::Paragraph);
    assert_eq!(heading.based_on.as_deref(), Some("Normal"));

    // A4 page geometry from sectPr (11906 x 16838 twips).
    assert_eq!(doc.sections.len(), 1);
    assert_eq!(doc.sections[0].page_width, 11906);
    assert_eq!(doc.sections[0].page_height, 16838);
}

#[test]
fn lowers_to_native_ir_with_synthesized_style_and_swatch() {
    let doc = import_docx(&memo_docx()).unwrap();
    let ir = lower(&doc);

    // The red color minted exactly one swatch.
    assert_eq!(ir.swatches.len(), 1);
    assert_eq!(ir.swatches[0].value, vec![255.0, 0.0, 0.0]);
    assert_eq!(ir.swatches[0].space, "RGB");

    // Heading1 exists and is created after its Normal parent.
    let ids: Vec<&str> = ir.styles.iter().map(|s| s.id.as_str()).collect();
    let n = ids.iter().position(|s| s.ends_with("docx-Normal")).unwrap();
    let h = ids
        .iter()
        .position(|s| s.ends_with("docx-Heading1"))
        .unwrap();
    assert!(n < h);

    // The heading style carries centered justification.
    let heading = ir
        .styles
        .iter()
        .find(|s| s.id.ends_with("docx-Heading1"))
        .unwrap();
    assert_eq!(heading.collection, StyleCollection::Paragraph);
    assert!(heading
        .props
        .iter()
        .any(|p| p.path == "paragraphJustification"
            && p.value == PropValue::Text("CenterAlign".into())));

    // The heading paragraph applies the Heading1 style.
    let heading_para = &ir.story.paragraphs[1];
    assert!(heading_para
        .para_style_id
        .as_deref()
        .unwrap()
        .ends_with("docx-Heading1"));

    // The bold-red run got a synthesized character style referencing the swatch.
    let mixed = &ir.story.paragraphs[2];
    let synth_id = mixed.runs[1].char_style_id.as_deref().unwrap();
    let synth = ir.styles.iter().find(|s| s.id == synth_id).unwrap();
    assert!(synth
        .props
        .iter()
        .any(|p| p.path == "characterFontStyle" && p.value == PropValue::Text("Bold".into())));
    assert!(synth.props.iter().any(|p| p.path == "characterFillColor"
        && p.value == PropValue::ColorRef(ir.swatches[0].id.clone())));
}

#[test]
fn doc_session_loads_and_emits_json() {
    let session = DocSession::load(&memo_docx()).unwrap();
    assert_eq!(session.block_count(), 3);
    let json = session.lowered_json();
    assert!(json.contains("\"swatches\""));
    assert!(json.contains("characterFillColor"));
    // Zero-edit save-back is verbatim.
    assert_eq!(session.save_verbatim(), memo_docx());
}

#[test]
fn smallest_document_without_styles_part() {
    let doc = import_docx(&one_paragraph_docx()).unwrap();
    assert_eq!(doc.body.len(), 1);
    assert!(doc.styles.styles.is_empty());
    let ir = lower(&doc);
    assert_eq!(ir.story.paragraphs[0].runs[0].text, "Hello, world.");
}

#[test]
fn malformed_container_errors_without_panicking() {
    assert!(import_docx(b"not a zip at all").is_err());
    // A zip that is not an OPC package (no document part) also errors cleanly.
    let junk = zip_parts(&[("random.bin", b"hello")]);
    assert!(import_docx(&junk).is_err());
}
