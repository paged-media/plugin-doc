//! thoughts ADR 029 — every Word section survives import, with its own
//! geometry and the block it starts at. A paragraph whose `pPr` carries a
//! `sectPr` ENDS a section; the importer used to read only the body-level
//! `sectPr`, so a two-section document came in as one section with the
//! LAST section's geometry. `pagination_docx()` is the document Word
//! paginated (fixtures/pagination.word.json).

use docx_conformance::pagination_docx;
use docx_core::SectionKind;
use docx_import::import_docx;

#[test]
fn both_sections_import_with_their_geometry_and_extent() {
    let doc = import_docx(&pagination_docx()).expect("import");
    assert_eq!(doc.sections.len(), 2, "{:#?}", doc.sections);

    let s1 = &doc.sections[0];
    assert_eq!((s1.page_width, s1.page_height), (12240, 15840), "US Letter");
    assert_eq!(s1.margin_top, 1440);
    assert_eq!(s1.first_block, 0);

    let s2 = &doc.sections[1];
    assert_eq!(
        (s2.page_width, s2.page_height),
        (11906, 8391),
        "A5 landscape"
    );
    assert_eq!(s2.margin_left, 720);
    assert_eq!(
        s2.first_block, 120,
        "section 2 starts after S1's 120 paragraphs"
    );
    assert_eq!(s2.kind, SectionKind::NextPage);
    assert_eq!(doc.body.len(), 160);
}

#[test]
fn a_single_section_document_still_has_one_section() {
    let doc = import_docx(&docx_conformance::memo_docx()).expect("import");
    assert_eq!(doc.sections.len(), 1);
    assert_eq!(doc.sections[0].first_block, 0);
}

/// Word's exact line spacing reaches the lowering as a leading. Without it the
/// engine used auto leading, and the standalone-open document ran a page
/// longer than Word's (thoughts ADR 029; the editor's doc-standalone-open
/// spec).
#[test]
fn exact_line_spacing_lowers_to_a_leading() {
    let doc = import_docx(&pagination_docx()).expect("import");
    let ls = doc
        .body
        .iter()
        .find_map(|b| match b {
            docx_core::Block::Paragraph(p) => p.props.line_spacing,
            _ => None,
        })
        .expect("the fixture's paragraphs carry w:spacing/@w:line");
    assert_eq!((ls.value, ls.rule), (240, docx_core::LineRule::Exact));

    let lowered = docx_lower::lower(&doc);
    // The first paragraph's own (synthesized) style carries it. (Word's
    // named styles carry a leading too since every unset line spacing lowers
    // to Word's single.)
    let first = lowered.story.paragraphs()[0]
        .para_style_id
        .clone()
        .expect("a styled paragraph");
    let leading = lowered
        .styles
        .iter()
        .find(|s| s.id == first)
        .and_then(|s| s.props.iter().find(|p| p.path == "characterLeading"))
        .map(|p| p.value.clone());
    assert_eq!(leading, Some(docx_lower::ir::PropValue::Length(12.0)));

    let firsts: Vec<usize> = lowered.sections.iter().map(|s| s.first_block).collect();
    assert_eq!(firsts, vec![0, 120], "every section, with its first block");
}
