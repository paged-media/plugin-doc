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

//! The Paged Annual's Word chapter fixture — `fixtures/annual-report.docx`.
//!
//! Unlike the in-memory builders, this fixture is a COMMITTED file: the
//! editor showcase places the byte-identical copy that lives at
//! `editor/apps/canvas/tests/showcase/assets/annual-report.docx`, generated
//! by the byte-stable `gen-annual-report-docx.py` beside it. This test is the
//! contract between the two repos: every tier the generator claims to
//! exercise must actually lower — styles + docDefaults, direct formatting,
//! both 2-level list formats through `numbering.xml`, a table with `gridSpan`
//! AND `vMerge`, one embedded (real, decodable) PNG, both external-hyperlink
//! forms, footnotes, tab stops, keepNext — with ZERO error-severity
//! diagnostics. If the generator drifts, this is the test that says so.

use std::path::PathBuf;

use docx_core::{Block, ListKind, Paragraph, VMerge};
use docx_import::import_docx;
use docx_lower::ir::{LoweredBlock, StyleCollection};
use docx_lower::lower;

fn fixture() -> Vec<u8> {
    let p = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("fixtures")
        .join("annual-report.docx");
    std::fs::read(&p).unwrap_or_else(|e| panic!("read fixture {}: {e}", p.display()))
}

fn paragraphs(doc: &docx_core::DocxDocument) -> Vec<&Paragraph> {
    doc.body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect()
}

#[test]
fn annual_report_fixture_lowers_every_tier() {
    let doc = import_docx(&fixture()).expect("annual-report.docx imports");
    let paras = paragraphs(&doc);

    // ── styles: docDefaults + the four named styles ─────────────────────────
    for id in ["Normal", "Heading1", "Heading2", "Caption"] {
        assert!(
            doc.styles.styles.iter().any(|s| s.style_id == id),
            "style {id} parsed"
        );
    }
    // keepNext rides the heading style (Tier-1a).
    let h1 = doc
        .styles
        .styles
        .iter()
        .find(|s| s.style_id == "Heading1")
        .unwrap();
    assert_eq!(h1.para.keep_next, Some(true));

    // ── lists: bullet + decimal, both at TWO levels ─────────────────────────
    let markers: Vec<_> = paras.iter().filter_map(|p| p.list.as_ref()).collect();
    let has = |kind: ListKind, level: u8| markers.iter().any(|m| m.kind == kind && m.level == level);
    assert!(has(ListKind::Bullet, 0), "level-0 bullet");
    assert!(has(ListKind::Bullet, 1), "level-1 bullet");
    assert!(has(ListKind::Numbered, 0), "level-0 decimal");
    assert!(has(ListKind::Numbered, 1), "level-1 decimal");
    // The Wingdings F0B7 glyph normalizes to U+2022; decimal maps to the IDML
    // numbering sample the engine's format_number reads.
    assert!(markers
        .iter()
        .any(|m| m.bullet_char.as_deref() == Some("\u{2022}")));
    assert!(markers
        .iter()
        .any(|m| m.number_format.as_deref() == Some("1, 2, 3, 4...")));

    // ── the table: a gridSpan title cell AND a vMerge'd region cell ─────────
    let table = doc
        .body
        .iter()
        .find_map(|b| match b {
            Block::Table(t) => Some(t),
            _ => None,
        })
        .expect("the circulation table");
    assert_eq!(table.column_widths.len(), 3);
    assert_eq!(table.rows[0].cells[0].grid_span, 3, "the spanning title cell");
    assert_eq!(table.rows[2].cells[0].v_merge, VMerge::Restart, "Alpine restarts");
    assert_eq!(table.rows[3].cells[0].v_merge, VMerge::Continue, "…and continues");

    // ── the embedded image: ONE real PNG ────────────────────────────────────
    let images: Vec<_> = paras
        .iter()
        .flat_map(|p| p.runs.iter())
        .filter_map(|r| r.image.as_ref())
        .collect();
    assert_eq!(images.len(), 1, "exactly one embedded image");
    assert!(images[0].bytes.starts_with(b"\x89PNG\r\n\x1a\n"));
    assert_eq!(images[0].mime, "image/png");
    // 1524000 x 1016000 EMU = the 120 x 80 pt plate mark.
    assert_eq!((images[0].width_emu, images[0].height_emu), (1524000, 1016000));

    // ── hyperlinks: BOTH Word forms resolve to their external URLs ──────────
    let urls: Vec<&str> = paras
        .iter()
        .flat_map(|p| p.runs.iter())
        .filter_map(|r| r.hyperlink.as_deref())
        .collect();
    assert!(
        urls.contains(&"https://paged.media/annual"),
        "the w:hyperlink form: {urls:?}"
    );
    assert!(
        urls.contains(&"https://docs.paged.media/"),
        "the fldSimple HYPERLINK form: {urls:?}"
    );

    // ── footnotes: the two real notes (separator pseudo-notes skipped) ──────
    assert_eq!(doc.notes.len(), 2, "two real footnotes");
    assert!(doc.notes.iter().all(|n| !n.endnote));
    let refs: Vec<i64> = paras
        .iter()
        .flat_map(|p| p.runs.iter())
        .filter_map(|r| r.note_ref)
        .collect();
    assert_eq!(refs, [2, 3], "both in-text references");

    // ── tab stops: the ledger lines' right tab at 4320 twips ────────────────
    // (The fixture authors a dot leader too, but leader glyphs are a labelled
    // later-tier refinement — the importer carries position + alignment only.)
    assert!(
        paras
            .iter()
            .any(|p| p.props.tabs.iter().any(|t| t.position == 4320
                && t.alignment.as_deref() == Some("right"))),
        "the right-aligned ledger tab stop"
    );

    // ── lowering: synthesized styles, native list props, zero errors ────────
    let ir = lower(&doc);
    for suffix in ["docx-Normal", "docx-Heading1", "docx-Heading2", "docx-Caption"] {
        assert!(
            ir.styles.iter().any(|s| s.id.ends_with(suffix)),
            "synthesized {suffix}"
        );
    }
    // Direct bold/italic/coloured runs synthesize character styles, and the
    // house-colour run minted an RGB swatch (28, 63, 148).
    assert!(ir
        .styles
        .iter()
        .any(|s| s.collection == StyleCollection::Character));
    assert!(ir
        .swatches
        .iter()
        .any(|s| s.value == vec![28.0, 63.0, 148.0]));

    // The table lowers with the span + merge intact.
    let lt = ir
        .story
        .blocks
        .iter()
        .find_map(|b| match b {
            LoweredBlock::Table(t) => Some(t),
            _ => None,
        })
        .expect("the lowered table");
    assert!(lt.cells.iter().any(|c| c.col_span == 3), "colSpan 3 title");
    assert!(lt.cells.iter().any(|c| c.row_span == 2), "rowSpan 2 region");

    // The image paragraph carries its anchored-frame placement as a data URI.
    assert!(ir
        .story
        .paragraphs()
        .iter()
        .any(|p| p.images.iter().any(|i| i.uri.starts_with("data:image/png;base64,")
            && i.width_pt == 120.0
            && i.height_pt == 80.0)));

    // Both link runs carry the URL for the native insertHyperlink lane.
    let lowered_urls: Vec<&str> = ir
        .story
        .paragraphs()
        .iter()
        .flat_map(|p| p.runs.iter())
        .filter_map(|r| r.hyperlink_url.as_deref())
        .collect();
    assert!(lowered_urls.contains(&"https://paged.media/annual"));
    assert!(lowered_urls.contains(&"https://docs.paged.media/"));

    // The footnotes surface their honest diagnostic (never inlined) — and NO
    // diagnostic anywhere is error-severity.
    assert!(ir
        .diagnostics
        .iter()
        .any(|d| d.message.contains("footnote")));
    let errors: Vec<_> = ir
        .diagnostics
        .iter()
        .filter(|d| d.severity == "error")
        .collect();
    assert!(errors.is_empty(), "zero error diagnostics: {errors:?}");
}
