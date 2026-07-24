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

use docx_conformance::{
    hyperlink_docx, image_docx, list_docx, memo_docx, one_paragraph_docx, table_docx, tier1_docx,
    zip_parts,
};
use docx_core::{Block, ListKind, StyleKind, VMerge};
use docx_import::import_docx;
use docx_js::DocSession;
use docx_lower::ir::{LoweredBlock, PropValue, StyleCollection};
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
    let heading_para = &ir.story.paragraphs()[1];
    assert!(heading_para
        .para_style_id
        .as_deref()
        .unwrap()
        .ends_with("docx-Heading1"));

    // The bold-red run got a synthesized character style referencing the swatch.
    let mixed = &ir.story.paragraphs()[2];
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
    assert_eq!(ir.story.paragraphs()[0].runs[0].text, "Hello, world.");
}

#[test]
fn tier1_doc_defaults_tabs_keeps_and_underline() {
    let doc = import_docx(&tier1_docx()).unwrap();

    // docDefaults parsed the base font/size.
    assert_eq!(doc.styles.doc_defaults.run.font.as_deref(), Some("Calibri"));
    assert_eq!(doc.styles.doc_defaults.run.size_half_pts, Some(22));
    assert_eq!(doc.styles.doc_defaults.para.space_after, Some(160));

    let para = match &doc.body[0] {
        Block::Paragraph(p) => p,
        _ => panic!("expected a paragraph"),
    };
    // Two real tab stops; the "clear" stop is dropped.
    assert_eq!(para.props.tabs.len(), 2);
    assert_eq!(para.props.tabs[0].position, 720);
    assert_eq!(para.props.tabs[1].alignment.as_deref(), Some("right"));
    assert_eq!(para.props.keep_next, Some(true));
    // Underline on for the single, OFF for the explicit none.
    assert_eq!(para.runs[1].props.underline, Some(true));
    assert_eq!(para.runs[2].props.underline, Some(false));

    let ir = lower(&doc);

    // docDefaults became a base paragraph style every un-based style inherits.
    let default = ir
        .styles
        .iter()
        .find(|s| s.id.ends_with("docx-Default"))
        .expect("docDefaults base style");
    assert!(default
        .props
        .iter()
        .any(|p| p.path == "characterFontFamily" && p.value == PropValue::Text("Calibri".into())));
    let normal = ir
        .styles
        .iter()
        .find(|s| s.id.ends_with("docx-Normal"))
        .unwrap();
    assert_eq!(normal.based_on.as_deref(), Some(default.id.as_str()));

    // The paragraph synthesized a style carrying keepWithNext + the tab stops.
    let pstyle_id = ir.story.paragraphs()[0].para_style_id.as_deref().unwrap();
    let pstyle = ir.styles.iter().find(|s| s.id == pstyle_id).unwrap();
    assert!(pstyle
        .props
        .iter()
        .any(|p| p.path == "paragraphKeepWithNext"));
    let tabs = pstyle
        .props
        .iter()
        .find(|p| p.path == "paragraphTabStops")
        .expect("tab stops lowered");
    match &tabs.value {
        PropValue::TabStops(stops) => {
            assert_eq!(stops.len(), 2);
            assert_eq!(stops[0].position, 36.0); // 720 twips -> 36 pt
            assert_eq!(stops[1].alignment.as_deref(), Some("right"));
        }
        _ => panic!("expected tabStops value"),
    }
    // The synthesized paragraph style inherits the docDefaults base.
    assert_eq!(pstyle.based_on.as_deref(), Some(default.id.as_str()));
}

#[test]
fn lists_resolve_through_numbering_and_lower_to_native_markers() {
    let doc = import_docx(&list_docx()).unwrap();

    let paras: Vec<_> = doc
        .body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    assert_eq!(paras.len(), 4);

    // Bullet list: numId 1 -> abstractNum 0 -> numFmt=bullet, Wingdings F0B7 glyph
    // normalized to U+2022.
    let bullet = paras[0].list.as_ref().unwrap();
    assert_eq!(bullet.kind, ListKind::Bullet);
    assert_eq!(bullet.bullet_char.as_deref(), Some("\u{2022}"));
    // Numbered list: numId 2 -> abstractNum 1 -> decimal.
    let numbered = paras[2].list.as_ref().unwrap();
    assert_eq!(numbered.kind, ListKind::Numbered);
    assert_eq!(numbered.number_format.as_deref(), Some("1, 2, 3, 4..."));

    let ir = lower(&doc);

    // The bullet paragraph applies a synthesized style setting the native list
    // type + bullet glyph (the field the renderer gates markers on).
    let bstyle_id = ir.story.paragraphs()[0].para_style_id.as_deref().unwrap();
    let bstyle = ir.styles.iter().find(|s| s.id == bstyle_id).unwrap();
    assert!(bstyle
        .props
        .iter()
        .any(|p| p.path == "paragraphListType" && p.value == PropValue::Text("BulletList".into())));
    assert!(bstyle
        .props
        .iter()
        .any(|p| p.path == "paragraphBulletCharacter"
            && p.value == PropValue::Text("\u{2022}".into())));

    // The numbered paragraph applies NumberedList + the IDML numbering-format
    // sample; the engine auto-counts NumberedList paragraphs.
    let nstyle_id = ir.story.paragraphs()[2].para_style_id.as_deref().unwrap();
    let nstyle = ir.styles.iter().find(|s| s.id == nstyle_id).unwrap();
    assert!(nstyle.props.iter().any(
        |p| p.path == "paragraphListType" && p.value == PropValue::Text("NumberedList".into())
    ));
    assert!(nstyle
        .props
        .iter()
        .any(|p| p.path == "paragraphNumberingFormat"));

    // The two identical bullet paragraphs share ONE synthesized list style.
    assert_eq!(
        ir.story.paragraphs()[0].para_style_id,
        ir.story.paragraphs()[1].para_style_id
    );
}

#[test]
fn tables_parse_grid_spans_merges_and_lower_to_native_cells() {
    let doc = import_docx(&table_docx()).unwrap();

    // Body order: paragraph, table, paragraph.
    assert!(matches!(doc.body[0], Block::Paragraph(_)));
    let table = match &doc.body[1] {
        Block::Table(t) => t,
        _ => panic!("expected a table at body[1]"),
    };
    assert!(matches!(doc.body[2], Block::Paragraph(_)));

    assert_eq!(table.column_widths, vec![2000, 3000]);
    assert_eq!(table.rows.len(), 3);
    // Row 0: one cell spanning both columns.
    assert_eq!(table.rows[0].cells[0].grid_span, 2);
    // Row 1 col 0 starts a vertical merge; row 2 col 0 continues it.
    assert_eq!(table.rows[1].cells[0].v_merge, VMerge::Restart);
    assert_eq!(table.rows[2].cells[0].v_merge, VMerge::Continue);

    let ir = lower(&doc);
    let lt = match &ir.story.blocks[1] {
        LoweredBlock::Table(t) => t,
        _ => panic!("expected a lowered table at block[1]"),
    };
    assert_eq!(lt.rows, 3);
    assert_eq!(lt.cols, 2);
    assert_eq!(lt.column_widths_pt, vec![100.0, 150.0]); // 2000/3000 twips -> pt

    // The spanning header cell: (0,0) with colSpan 2.
    let header = lt.cells.iter().find(|c| c.row == 0 && c.col == 0).unwrap();
    assert_eq!(header.col_span, 2);
    // The merged cell: (1,0) absorbs the continue below -> rowSpan 2, and the
    // continue cell itself is NOT emitted.
    let merged = lt.cells.iter().find(|c| c.row == 1 && c.col == 0).unwrap();
    assert_eq!(merged.row_span, 2);
    assert!(!lt.cells.iter().any(|c| c.row == 2 && c.col == 0));
    // Emitted cells: header(0,0) + (1,0) + (1,1) + (2,1) = 4.
    assert_eq!(lt.cells.len(), 4);
}

#[test]
fn inline_image_resolves_media_and_lowers_to_anchored_frame() {
    let doc = import_docx(&image_docx()).unwrap();

    // The drawing run carries the resolved image (bytes + intrinsic EMU size).
    let img_para = match &doc.body[1] {
        Block::Paragraph(p) => p,
        _ => panic!("expected the image paragraph"),
    };
    let img = img_para
        .runs
        .iter()
        .find_map(|r| r.image.as_ref())
        .expect("the drawing run resolved an image");
    assert!(img.bytes.starts_with(b"\x89PNG"));
    assert_eq!(img.mime, "image/png");
    assert_eq!((img.width_emu, img.height_emu), (914400, 685800));

    let ir = lower(&doc);
    // Lowered: the image paragraph carries an anchored-frame placement sized in
    // points (914400 EMU = 72 pt, 685800 EMU = 54 pt) with a data: URI.
    let lp = ir.story.paragraphs()[1];
    assert_eq!(lp.images.len(), 1);
    assert_eq!(lp.images[0].width_pt, 72.0);
    assert_eq!(lp.images[0].height_pt, 54.0);
    assert!(lp.images[0].uri.starts_with("data:image/png;base64,"));
}

#[test]
fn hyperlink_resolves_target_and_styles_run_blue_underline() {
    let doc = import_docx(&hyperlink_docx()).unwrap();
    let para = match &doc.body[0] {
        Block::Paragraph(p) => p,
        _ => panic!("expected a paragraph"),
    };
    // 3 runs: "Visit ", the hyperlink "Paged Media", " today." — the middle run
    // carries the resolved external URL.
    assert_eq!(para.runs.len(), 3);
    assert_eq!(para.runs[0].hyperlink, None);
    assert_eq!(
        para.runs[1].hyperlink.as_deref(),
        Some("https://paged.media/")
    );
    assert_eq!(para.runs[2].hyperlink, None);

    let ir = lower(&doc);
    // The hyperlink run got a synthesized character style with blue fill +
    // underline; a swatch was minted for the link blue.
    let link_run = &ir.story.paragraphs()[0].runs[1];
    let style_id = link_run.char_style_id.as_deref().unwrap();
    let style = ir.styles.iter().find(|s| s.id == style_id).unwrap();
    assert!(style
        .props
        .iter()
        .any(|p| p.path == "characterUnderline" && p.value == PropValue::Bool(true)));
    assert!(style.props.iter().any(|p| p.path == "characterFillColor"));
    assert!(ir.swatches.iter().any(|s| s.value == vec![0.0, 0.0, 255.0]));
    // The run also carries the external URL so the host-model emits a native
    // clickable `insertHyperlink` over its range (not just the blue-underline look).
    assert_eq!(
        link_run.hyperlink_url.as_deref(),
        Some("https://paged.media/")
    );
    // The other runs are not links.
    assert_eq!(ir.story.paragraphs()[0].runs[0].hyperlink_url, None);
    // The diagnostic now reports how many became native clickable links.
    assert!(ir
        .diagnostics
        .iter()
        .any(|d| d.message.contains("native clickable link")));
}

#[test]
fn malformed_container_errors_without_panicking() {
    assert!(import_docx(b"not a zip at all").is_err());
    // A zip that is not an OPC package (no document part) also errors cleanly.
    let junk = zip_parts(&[("random.bin", b"hello")]);
    assert!(import_docx(&junk).is_err());
}
