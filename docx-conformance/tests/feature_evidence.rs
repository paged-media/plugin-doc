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

//! Capability-evidence tests for paged.doc. Each test name carries a
//! `__feat__<feature id>` suffix (`.`/`-` → `_`) so `paged-state` attributes the
//! run to the registry row it proves — the same convention as core's
//! `scripting_feature_evidence.rs`. These are deliberately THIN: each asserts the
//! headline promise of one registry feature; the detailed behaviour lives in
//! `conformance.rs` / `save_back.rs`.

// The `__feat__` suffix is the state-registry attribution convention, not Rust
// style (same as core's scripting_feature_evidence.rs).
#![allow(non_snake_case)]

use docx_conformance::{memo_docx, table_docx};
use docx_core::Block;
use docx_export::{EditSet, RunEdit};
use docx_import::import_docx;
use docx_js::DocSession;
use paged_ooxml::OpcPackage;

/// The preservation invariant: a zero-edit round-trip is per-part byte-identical,
/// INCLUDING a part paged does not model.
#[test]
fn opc_zero_edit_round_trip_is_byte_identical__feat__plugin_doc_opc_foundation() {
    let bytes = memo_docx();
    let pkg = OpcPackage::read(&bytes).unwrap();
    let again = OpcPackage::read(&pkg.write().unwrap()).unwrap();
    for name in pkg.file_names() {
        assert_eq!(pkg.part(name), again.part(name), "part {name}");
    }
    assert_eq!(
        again.part("customXml/unknown.txt").unwrap(),
        b"paged preserves unknown parts",
        "an unmodelled part survives"
    );
}

/// The read path: a real `.docx` lowers to native stories + a style catalog.
#[test]
fn docx_lowers_to_native_stories_and_styles__feat__plugin_doc_read_path() {
    let doc = import_docx(&memo_docx()).unwrap();
    let ir = docx_lower::lower(&doc);
    let paras = ir.story.paragraphs();
    assert_eq!(paras.len(), 3, "three body paragraphs");
    assert_eq!(paras[1].runs[0].text, "A Centered Heading");
    // Direct Word formatting became a SYNTHESIZED named style (applyStyle is the
    // only range-styling mutation), referencing a minted swatch.
    let synth = paras[2].runs[1].char_style_id.as_deref().unwrap();
    assert!(ir.styles.iter().any(|s| s.id == synth));
    assert!(!ir.swatches.is_empty(), "the red run minted a swatch");
}

/// Embedded placement: the lowering carries what the bundle pours through host
/// mutations (a style catalog first, then the story blocks).
#[test]
fn lowering_carries_the_host_pour__feat__plugin_doc_embedded_placement() {
    let session = DocSession::load(&memo_docx()).unwrap();
    let ir = session.lowered();
    assert!(!ir.story.blocks.is_empty(), "a story to pour");
    assert!(!ir.styles.is_empty(), "a style catalog to create first");
    assert_eq!(session.block_count(), 3);
}

/// Edited save-back: a targeted patch changes the target and leaves every other
/// part byte-identical.
#[test]
fn edited_save_back_patches_only_the_target__feat__plugin_doc_save_back() {
    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    let edits = EditSet {
        runs: vec![RunEdit::text(0, 0, "Evidence edit.")],
        ..Default::default()
    };
    let (saved, _skips) = session.save_edited(&edits).unwrap();

    let re = import_docx(&saved).unwrap();
    let Block::Paragraph(p0) = &re.body[0] else {
        panic!("expected a paragraph")
    };
    assert_eq!(p0.runs[0].text, "Evidence edit.");

    let before = OpcPackage::read(&original).unwrap();
    let after = OpcPackage::read(&saved).unwrap();
    for name in before.file_names() {
        if name == "word/document.xml" {
            continue;
        }
        assert_eq!(before.part(name), after.part(name), "part {name} untouched");
    }
}

/// Save-back reaches TABLE CELL content too (its own tbl/tr/tc/p locator path).
#[test]
fn save_back_reaches_table_cells__feat__plugin_doc_save_back() {
    let session = DocSession::load(&table_docx()).unwrap();
    let edits = EditSet {
        cells: vec![docx_export::CellRunEdit::text(1, 0, 0, 0, "cell evidence")],
        ..Default::default()
    };
    let (saved, _skips) = session.save_edited(&edits).unwrap();
    let ir = docx_lower::lower(&import_docx(&saved).unwrap());
    let table = ir
        .story
        .blocks
        .iter()
        .find_map(|b| match b {
            docx_lower::ir::LoweredBlock::Table(t) => Some(t),
            _ => None,
        })
        .expect("the table survives");
    assert_eq!(table.cells[0].paragraphs[0].runs[0].text, "cell evidence");
}
