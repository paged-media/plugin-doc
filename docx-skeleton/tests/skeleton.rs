// ADR 029 — the section skeleton of docx_conformance::pagination_docx(), the
// document Word paginated (docx-conformance/fixtures/pagination.word.json).

use std::io::Read;

use docx_skeleton::{section_story_id, skeleton, skeleton_document};

fn pagination() -> docx_core::DocxDocument {
    docx_import::import_docx(&docx_conformance::pagination_docx()).expect("import")
}

#[test]
fn one_page_frame_and_growing_story_per_section() {
    let doc = skeleton_document(&pagination());
    assert_eq!(doc.spreads.len(), 2);
    assert_eq!(doc.stories.len(), 2);

    // Section 1: Letter, 1 in margins; section 2: A5 landscape, 0.5 in.
    let p1 = &doc.spreads[0].spread.pages[0].bounds;
    assert_eq!((p1.right, p1.bottom), (612.0, 792.0));
    let f1 = &doc.spreads[0].spread.text_frames[0].bounds;
    assert_eq!(
        (f1.top, f1.left, f1.bottom, f1.right),
        (72.0, 72.0, 720.0, 540.0)
    );
    let p2 = &doc.spreads[1].spread.pages[0].bounds;
    assert!((p2.right - 595.3).abs() < 0.01 && (p2.bottom - 419.55).abs() < 0.01);
    let f2 = &doc.spreads[1].spread.text_frames[0].bounds;
    assert_eq!((f2.top, f2.left), (36.0, 36.0));

    for (k, s) in doc.stories.iter().enumerate() {
        assert_eq!(s.self_id, section_story_id(k));
        let grow = s.story.grow.as_ref().expect("each section story grows");
        assert!(
            grow.copy_frame_options,
            "Word mode: generated pages keep LeadingOffset"
        );
    }
    // Each section frame threads its own story.
    assert_eq!(doc.frame_chain(&section_story_id(0)).len(), 1);
    assert_eq!(doc.frame_chain(&section_story_id(1)).len(), 1);
    assert_eq!(doc.growing_stories().len(), 2);
}

#[test]
fn the_package_round_trips_through_the_native_codec() {
    let sk = skeleton(&pagination(), "pagination.docx").expect("skeleton");
    assert_eq!(
        sk.section_stories,
        vec![section_story_id(0), section_story_id(1)]
    );
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&sk.paged)).expect("zip");
    assert_eq!(
        zip.by_index(0).unwrap().name(),
        "mimetype",
        "OCF: mimetype first"
    );
    let mut pgm = Vec::new();
    zip.by_name(paged_store::DOCUMENT_PGM_PATH)
        .expect("native model part")
        .read_to_end(&mut pgm)
        .unwrap();
    let back = paged_store::from_bytes(&pgm).expect("decode");
    assert_eq!(
        back.growing_stories().len(),
        2,
        "grow rules survive the codec"
    );
}

/// ADR 029 — continuous sections that Word keeps on the page they continue
/// share that page's story (docx_conformance::continuous_docx(), measured in
/// docx-conformance/fixtures/continuous.word.json): 17 Word sections, 11
/// stories, and the lowering assigns every section the story it pours into.
#[test]
fn continuous_sections_share_the_story_of_the_page_they_continue() {
    let docx = docx_import::import_docx(&docx_conformance::continuous_docx()).expect("import");
    assert_eq!(docx.sections.len(), 17);
    let doc = skeleton_document(&docx);
    assert_eq!(doc.spreads.len(), 11, "a page per story, not per section");
    assert_eq!(doc.stories.len(), 11);
    assert_eq!(doc.growing_stories().len(), 11);
    for (k, s) in doc.stories.iter().enumerate() {
        assert_eq!(s.self_id, section_story_id(k));
        assert_eq!(doc.frame_chain(&section_story_id(k)).len(), 1);
    }

    // Each story's page and frame are its FIRST section's: story 0 is A1's
    // (A2, A3 join it), story 2 is B2's two columns, story 6 is C1's 0.5 in
    // margins (C2's other margins become indents), story 9 is F2's 6 in page.
    let frame = |k: usize| &doc.spreads[k].spread.text_frames[0];
    let page = |k: usize| &doc.spreads[k].spread.pages[0].bounds;
    assert_eq!((page(0).right, page(0).bottom), (360.0, 312.0));
    let f0 = &frame(0).bounds;
    assert_eq!(
        (f0.top, f0.left, f0.bottom, f0.right),
        (36.0, 36.0, 276.0, 324.0)
    );
    assert_eq!(frame(0).column_count, None);
    assert_eq!(frame(2).column_count, Some(2));
    assert_eq!(frame(5).column_count, Some(2), "D1 + nextColumn D2");
    let f6 = &frame(6).bounds;
    assert_eq!((f6.left, f6.right), (36.0, 324.0));
    assert_eq!(page(9).right, 432.0);
    assert_eq!(page(10).right, 360.0);

    // The pour's view agrees: one block group per skeleton story.
    let ir = docx_lower::lower(&docx);
    let stories: Vec<usize> = ir.sections.iter().map(|s| s.story).collect();
    assert_eq!(
        stories,
        vec![0, 0, 0, 1, 2, 3, 4, 5, 5, 6, 6, 6, 7, 7, 8, 9, 10]
    );
    let sk = skeleton(&docx, "continuous.docx").expect("skeleton");
    assert_eq!(sk.section_stories, docx_skeleton::story_ids(&docx));
    assert_eq!(sk.section_stories.len(), stories.last().unwrap() + 1);
}
