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
    assert_eq!((s2.page_width, s2.page_height), (11906, 8391), "A5 landscape");
    assert_eq!(s2.margin_left, 720);
    assert_eq!(s2.first_block, 120, "section 2 starts after S1's 120 paragraphs");
    assert_eq!(s2.kind, SectionKind::NextPage);
    assert_eq!(doc.body.len(), 160);
}

#[test]
fn a_single_section_document_still_has_one_section() {
    let doc = import_docx(&docx_conformance::memo_docx()).expect("import");
    assert_eq!(doc.sections.len(), 1);
    assert_eq!(doc.sections[0].first_block, 0);
}
