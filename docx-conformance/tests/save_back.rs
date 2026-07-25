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

//! M2 edited save-back vertical slice: a `.docx` imports, two run edits (a text
//! change and a property change) are applied, and the saved package is asserted
//! to (a) carry exactly those edits, (b) leave every other part AND every
//! untouched subtree of `word/document.xml` byte-identical, and (c) round-trip
//! through re-import. The `EditSet` is hand-authored here — the LIVE editor
//! wiring is the deferred DOC-03 seam.

use docx_conformance::memo_docx;
use docx_core::{Block, RunProps};
use docx_export::{EditSet, RunEdit};
use docx_import::import_docx;
use docx_js::DocSession;
use paged_ooxml::OpcPackage;

fn body_para(doc: &docx_core::DocxDocument, i: usize) -> &docx_core::Paragraph {
    match &doc.body[i] {
        Block::Paragraph(p) => p,
        _ => panic!("body[{i}] is not a paragraph"),
    }
}

#[test]
fn edited_save_back_patches_targets_and_preserves_everything_else() {
    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();

    // Edit 1: replace p0's only run text. Edit 2: toggle bold OFF the "bold red"
    // run (p2, run 1) — its synthesized bold+red style projects to a direct
    // `<w:rPr>` carrying only the color.
    let edits = EditSet {
        structural: vec![],
        cells: vec![],
        runs: vec![
            RunEdit::text(0, 0, "Edited body text."),
            RunEdit::props(
                2,
                1,
                RunProps {
                    color: Some("FF0000".into()),
                    ..Default::default()
                },
            ),
        ],
    };
    let saved = session.save_edited(&edits).unwrap();

    let saved_pkg = OpcPackage::read(&saved).unwrap();
    let doc_xml =
        std::str::from_utf8(saved_pkg.part("word/document.xml").expect("document.xml")).unwrap();

    // (a) the two targets changed.
    assert!(doc_xml.contains(">Edited body text.<"), "p0 text replaced");
    assert!(!doc_xml.contains(">Plain body text.<"), "old p0 text gone");
    assert!(
        doc_xml.contains(r#"<w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>bold red</w:t>"#),
        "bold dropped, color kept, text intact:\n{doc_xml}"
    );
    assert!(
        !doc_xml.contains("<w:b/>"),
        "no bold toggle remains anywhere"
    );

    // (b) preservation — every OTHER part decompressed-identical.
    let orig_pkg = OpcPackage::read(&original).unwrap();
    for name in orig_pkg.file_names() {
        if name == "word/document.xml" {
            continue;
        }
        assert_eq!(
            orig_pkg.part(name),
            saved_pkg.part(name),
            "part {name} must be byte-identical"
        );
    }
    // ...and within document.xml, every untouched subtree survives verbatim.
    assert!(doc_xml.contains(r#"<w:pStyle w:val="Heading1"/>"#));
    assert!(doc_xml.contains(">A Centered Heading<"));
    assert!(doc_xml.contains("<w:sectPr>"));
    assert!(
        doc_xml.contains(">Mix of normal and <"),
        "p2 run 0 untouched"
    );
    assert!(doc_xml.contains("> text.<"), "p2 run 2 untouched");

    // (c) round-trip: re-import reflects exactly the edits and nothing else.
    let re = import_docx(&saved).unwrap();
    assert_eq!(body_para(&re, 0).runs[0].text, "Edited body text.");
    assert_eq!(body_para(&re, 1).runs[0].text, "A Centered Heading");
    let p2 = body_para(&re, 2);
    assert_eq!(p2.runs[0].text, "Mix of normal and ");
    assert_eq!(p2.runs[1].text, "bold red");
    assert_eq!(p2.runs[1].props.bold, None, "bold toggled off");
    assert_eq!(
        p2.runs[1].props.color.as_deref(),
        Some("FF0000"),
        "color kept"
    );
    assert_eq!(p2.runs[2].text, " text.");
}

#[test]
fn zero_edit_save_back_still_byte_identical() {
    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    // An empty edit set patches nothing; verbatim carry-through holds.
    let saved = session.save_edited(&EditSet::default()).unwrap();
    let orig_pkg = OpcPackage::read(&original).unwrap();
    let saved_pkg = OpcPackage::read(&saved).unwrap();
    for name in orig_pkg.file_names() {
        assert_eq!(orig_pkg.part(name), saved_pkg.part(name), "part {name}");
    }
}

#[test]
fn bindings_run_counts_match_the_lowered_story() {
    // `build_bindings` replays lowering's `!text.is_empty()` filter; if it drifts,
    // `(block, run)` coordinates would silently mis-resolve. Assert alignment.
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&memo_docx()).unwrap();
    let bindings = docx_export::build_bindings(&model);
    let lowered = docx_lower::lower(&model);
    assert_eq!(bindings.blocks.len(), lowered.story.blocks.len());
    for (i, block) in lowered.story.blocks.iter().enumerate() {
        if let docx_lower::ir::LoweredBlock::Paragraph(p) = block {
            match &bindings.blocks[i] {
                docx_export::BlockBinding::Paragraph { runs, .. } => {
                    assert_eq!(runs.len(), p.runs.len(), "run-count drift at block {i}");
                }
                _ => panic!("block {i} should bind as a paragraph"),
            }
        }
    }
}

#[test]
fn diff_of_identical_lowerings_produces_no_edits() {
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&memo_docx()).unwrap();
    let bindings = docx_export::build_bindings(&model);
    let base = docx_lower::lower(&model);
    let edits = docx_export::diff(&base, &base, &bindings);
    assert!(edits.runs.is_empty(), "identical lowerings ⇒ no edits");
}

#[test]
fn diff_drives_save_back_end_to_end() {
    use docx_lower::ir::{LoweredBlock, LoweredStyle, PropValue, StyleCollection, StyleProp};

    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&original).unwrap();
    let bindings = docx_export::build_bindings(&model);
    let base = docx_lower::lower(&model);

    // Simulate an edit: change p0's text, and repoint the "bold red" run (block 2,
    // run 1) at a NEW color-only synthesized style (i.e. the user removed bold).
    let red_swatch = base
        .swatches
        .iter()
        .find(|s| s.value == vec![255.0, 0.0, 0.0])
        .expect("red swatch")
        .id
        .clone();
    let mut edited = base.clone();
    if let LoweredBlock::Paragraph(p) = &mut edited.story.blocks[0] {
        p.runs[0].text = "Edited via diff.".into();
    }
    edited.styles.push(LoweredStyle {
        id: "CharacterStyle/docx-auto-cNEW".into(),
        name: "color only".into(),
        collection: StyleCollection::Character,
        based_on: None,
        props: vec![StyleProp {
            path: "characterFillColor".into(),
            value: PropValue::ColorRef(red_swatch),
        }],
    });
    if let LoweredBlock::Paragraph(p) = &mut edited.story.blocks[2] {
        p.runs[1].char_style_id = Some("CharacterStyle/docx-auto-cNEW".into());
    }

    // Diff → EditSet → save-back → assert both edits landed.
    let edits = docx_export::diff(&base, &edited, &bindings);
    assert_eq!(edits.runs.len(), 2, "one text + one property edit");
    let saved = session.save_edited(&edits).unwrap();

    let re = import_docx(&saved).unwrap();
    assert_eq!(body_para(&re, 0).runs[0].text, "Edited via diff.");
    let p2 = body_para(&re, 2);
    assert_eq!(p2.runs[1].text, "bold red", "text untouched");
    assert_eq!(p2.runs[1].props.bold, None, "bold removed by the diff");
    assert_eq!(
        p2.runs[1].props.color.as_deref(),
        Some("FF0000"),
        "color kept"
    );
}

#[test]
fn doc03_read_overlay_diff_save_round_trips() {
    // The LIVE DOC-03 path, with the host read mocked: build a `StoryContent`
    // mirroring the baseline (as `host.document.storyContent` would return),
    // apply two edits to it, and drive `save_edited_from_content`.
    use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};
    use docx_lower::ir::LoweredBlock;

    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&original).unwrap();
    let base = docx_lower::lower(&model);

    // The read-back reflects the plugin's own story: each run's characterStyle IS
    // the plugin's char_style_id token, so it mirrors the baseline exactly...
    let mut content = StoryContentIn {
        self_id: "Story/doc".into(),
        paragraphs: base
            .story
            .blocks
            .iter()
            .filter_map(|b| match b {
                LoweredBlock::Paragraph(p) => Some(ParagraphContentIn {
                    paragraph_style: p.para_style_id.clone(),
                    runs: p
                        .runs
                        .iter()
                        .map(|r| RunContentIn {
                            text: r.text.clone(),
                            character_style: r.char_style_id.clone(),
                        })
                        .collect(),
                }),
                _ => None,
            })
            .collect(),
    };
    // ...then the "user" edits: change p0's text, and clear the bold-red run's
    // style (bold + color removed).
    content.paragraphs[0].runs[0].text = "Edited via content.".into();
    content.paragraphs[2].runs[1].character_style = None;

    let saved = session.save_edited_from_content(&content).unwrap();

    let re = import_docx(&saved).unwrap();
    assert_eq!(body_para(&re, 0).runs[0].text, "Edited via content.");
    let p2 = body_para(&re, 2);
    assert_eq!(p2.runs[1].text, "bold red", "text untouched");
    assert_eq!(
        p2.runs[1].props.bold, None,
        "bold cleared via the read overlay"
    );
    assert_eq!(p2.runs[1].props.color, None, "color cleared");
    // Untouched runs stay put.
    assert_eq!(p2.runs[0].text, "Mix of normal and ");
    assert_eq!(body_para(&re, 1).runs[0].text, "A Centered Heading");
}

#[test]
fn doc03_identity_content_is_a_no_op() {
    // A read-back identical to the baseline ⇒ no edits ⇒ verbatim save.
    use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};
    use docx_lower::ir::LoweredBlock;

    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&original).unwrap();
    let base = docx_lower::lower(&model);
    let content = StoryContentIn {
        self_id: "Story/doc".into(),
        paragraphs: base
            .story
            .blocks
            .iter()
            .filter_map(|b| match b {
                LoweredBlock::Paragraph(p) => Some(ParagraphContentIn {
                    paragraph_style: p.para_style_id.clone(),
                    runs: p
                        .runs
                        .iter()
                        .map(|r| RunContentIn {
                            text: r.text.clone(),
                            character_style: r.char_style_id.clone(),
                        })
                        .collect(),
                }),
                _ => None,
            })
            .collect(),
    };
    let saved = session.save_edited_from_content(&content).unwrap();
    let orig_pkg = OpcPackage::read(&original).unwrap();
    let saved_pkg = OpcPackage::read(&saved).unwrap();
    for name in orig_pkg.file_names() {
        assert_eq!(orig_pkg.part(name), saved_pkg.part(name), "part {name}");
    }
}

#[test]
fn structural_edits_insert_and_delete_runs_and_paragraphs() {
    // Increment 2, driven through the DOC-03 overlay: delete a run, add a run,
    // and the untouched content stays byte-identical.
    use docx_export::{ParagraphContentIn, RunContentIn, StoryContentIn};
    use docx_lower::ir::LoweredBlock;

    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&original).unwrap();
    let base = docx_lower::lower(&model);

    let mut content = StoryContentIn {
        self_id: "Story/doc".into(),
        paragraphs: base
            .story
            .blocks
            .iter()
            .filter_map(|b| match b {
                LoweredBlock::Paragraph(p) => Some(ParagraphContentIn {
                    paragraph_style: p.para_style_id.clone(),
                    runs: p
                        .runs
                        .iter()
                        .map(|r| RunContentIn {
                            text: r.text.clone(),
                            character_style: r.char_style_id.clone(),
                        })
                        .collect(),
                }),
                _ => None,
            })
            .collect(),
    };
    // p2 is "Mix of normal and " + "bold red" + " text." — drop the middle run
    // and append a new one at the end.
    content.paragraphs[2].runs.remove(1);
    content.paragraphs[2].runs.push(RunContentIn {
        text: " (appended)".into(),
        character_style: None,
    });

    let saved = session.save_edited_from_content(&content).unwrap();
    let re = import_docx(&saved).unwrap();
    let p2 = body_para(&re, 2);
    let texts: Vec<&str> = p2.runs.iter().map(|r| r.text.as_str()).collect();
    assert_eq!(
        texts,
        ["Mix of normal and ", " text.", " (appended)"],
        "middle run deleted, new run appended"
    );
    // Other paragraphs untouched; other parts byte-identical.
    assert_eq!(body_para(&re, 0).runs[0].text, "Plain body text.");
    assert_eq!(body_para(&re, 1).runs[0].text, "A Centered Heading");
    let orig_pkg = OpcPackage::read(&original).unwrap();
    let saved_pkg = OpcPackage::read(&saved).unwrap();
    for name in orig_pkg.file_names() {
        if name == "word/document.xml" {
            continue;
        }
        assert_eq!(orig_pkg.part(name), saved_pkg.part(name), "part {name}");
    }
}

#[test]
fn structural_paragraph_delete_and_append() {
    use docx_export::{EditSet, StructuralEdit};

    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    let edits = EditSet {
        runs: vec![],
        cells: vec![],
        structural: vec![
            StructuralEdit::DeleteParagraph { block: 1 }, // the heading
            StructuralEdit::InsertParagraph {
                block: 2,
                text: "A brand new paragraph.".into(),
                props: Default::default(),
                para_style: None,
                rstyle: None,
            },
        ],
    };
    let saved = session.save_edited(&edits).unwrap();
    let re = import_docx(&saved).unwrap();
    let texts: Vec<String> = re
        .body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        })
        .collect();
    assert_eq!(
        texts,
        [
            "Plain body text.",
            "Mix of normal and bold red text.",
            "A brand new paragraph.",
        ],
        "heading deleted, new paragraph appended after the last one"
    );
}

#[test]
fn table_cell_text_round_trips() {
    use docx_conformance::table_docx;
    use docx_export::CellRunEdit;

    let original = table_docx();
    let session = DocSession::load(&original).unwrap();
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&original).unwrap();
    let base = docx_lower::lower(&model);

    // Find the table block + its first cell's first run.
    let (block, table) = base
        .story
        .blocks
        .iter()
        .enumerate()
        .find_map(|(i, b)| match b {
            docx_lower::ir::LoweredBlock::Table(t) => Some((i, t)),
            _ => None,
        })
        .expect("a table block");
    let before = table.cells[0].paragraphs[0].runs[0].text.clone();
    assert!(!before.is_empty(), "the first cell has text");

    let edits = docx_export::EditSet {
        cells: vec![CellRunEdit::text(block, 0, 0, 0, "PATCHED CELL")],
        ..Default::default()
    };
    let saved = session.save_edited(&edits).unwrap();

    // The cell text changed...
    let re_model = docx_import::import_docx(&saved).unwrap();
    let re = docx_lower::lower(&re_model);
    let re_table = re
        .story
        .blocks
        .iter()
        .find_map(|b| match b {
            docx_lower::ir::LoweredBlock::Table(t) => Some(t),
            _ => None,
        })
        .expect("table survives");
    assert_eq!(re_table.cells[0].paragraphs[0].runs[0].text, "PATCHED CELL");
    // ...the OTHER cells are untouched, and the grid survives.
    assert_eq!(re_table.cells.len(), table.cells.len());
    assert_eq!(re_table.column_widths_pt, table.column_widths_pt);
    for (i, (a, b)) in table.cells.iter().zip(&re_table.cells).enumerate().skip(1) {
        let at: String = a
            .paragraphs
            .iter()
            .flat_map(|p| p.runs.iter())
            .map(|r| r.text.as_str())
            .collect();
        let bt: String = b
            .paragraphs
            .iter()
            .flat_map(|p| p.runs.iter())
            .map(|r| r.text.as_str())
            .collect();
        assert_eq!(at, bt, "cell {i} untouched");
        assert_eq!(a.row_span, b.row_span, "cell {i} rowSpan kept");
        assert_eq!(a.col_span, b.col_span, "cell {i} colSpan kept");
    }
    // Every other part stays byte-identical.
    let orig_pkg = OpcPackage::read(&original).unwrap();
    let saved_pkg = OpcPackage::read(&saved).unwrap();
    for name in orig_pkg.file_names() {
        if name == "word/document.xml" {
            continue;
        }
        assert_eq!(orig_pkg.part(name), saved_pkg.part(name), "part {name}");
    }
}

#[test]
fn table_cell_diff_drives_save_back() {
    use docx_conformance::table_docx;

    let original = table_docx();
    let session = DocSession::load(&original).unwrap();
    let (model, _pkg, _main) = docx_import::import_docx_with_package(&original).unwrap();
    let bindings = docx_export::build_bindings(&model);
    let base = docx_lower::lower(&model);

    // Edit a cell in a cloned lowering; the differ must produce a cell edit.
    let mut edited = base.clone();
    if let docx_lower::ir::LoweredBlock::Table(t) = &mut edited.story.blocks[1] {
        t.cells[1].paragraphs[0].runs[0].text = "via diff".into();
    }
    let edits = docx_export::diff(&base, &edited, &bindings);
    assert_eq!(edits.cells.len(), 1, "one cell edit: {:?}", edits.cells);

    let saved = session.save_edited(&edits).unwrap();
    let re = docx_lower::lower(&docx_import::import_docx(&saved).unwrap());
    if let docx_lower::ir::LoweredBlock::Table(t) = &re.story.blocks[1] {
        assert_eq!(t.cells[1].paragraphs[0].runs[0].text, "via diff");
    } else {
        panic!("expected a table at block 1");
    }
}

#[test]
fn non_patchable_target_is_skipped_not_errored() {
    let original = memo_docx();
    let session = DocSession::load(&original).unwrap();
    // Out-of-range block/run resolves to nothing → skipped, document unchanged.
    let edits = EditSet {
        structural: vec![],
        cells: vec![],
        runs: vec![RunEdit::text(99, 0, "nowhere")],
    };
    let saved = session.save_edited(&edits).unwrap();
    let saved_pkg = OpcPackage::read(&saved).unwrap();
    let doc_xml = std::str::from_utf8(saved_pkg.part("word/document.xml").unwrap()).unwrap();
    assert!(doc_xml.contains(">Plain body text.<"), "unchanged");
    assert!(!doc_xml.contains("nowhere"));
}
