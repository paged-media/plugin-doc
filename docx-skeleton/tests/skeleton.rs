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
