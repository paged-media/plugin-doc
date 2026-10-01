// paged.doc — the native page skeleton for DOCX standalone open (ADR 029).
//
// A Word document is lowered as ONE STORY PER SECTION: each section gets a
// page of its size, its margin box as the text frame, and a story whose grow
// rule lets the engine add pages while it oversets. Growth appends after a
// story's last frame, so section 1's pages land before section 2's, where
// Word puts them. The frame's first baseline is `LeadingOffset` and generated
// pages copy it (`copy_frame_options`): Word fits a line only when its whole
// line box fits, which is what that expresses. Measured: core's
// docx_pagination_pipeline.rs matches Word's PDF of
// docx_conformance::pagination_docx() page for page.
//
// The skeleton carries NO text: the bundle's existing pour (styles, lists,
// tables, images, links) fills story k with section k's paragraphs after
// `host.nativeDocument.open`. The grow rules travel inside the native model,
// so opening needs no protocol op.
//
// Known simplification: every section starts on a new page. Word's
// `continuous` sections continue on the same page; that needs two stories
// sharing a page and is not modelled yet (the kind is kept on the import).

use std::collections::HashMap;

use docx_core::DocxDocument;
use paged_model::{
    Bounds, DesignMap, FirstBaselineOffset, FlowGrowRule, MarginPreference, Page, Paragraph,
    SpreadRef, Story, StoryRef, TextFrame,
};
use paged_scene::{Document, ParsedSpread, ParsedStory};

/// Word's default space between columns (`w:cols/@w:space`, 720 twips).
const WORD_COLUMN_GAP_PT: f32 = 36.0;

/// The skeleton: `.paged` bytes for `host.nativeDocument.open`, and the
/// story id of each section, in section order (the pour's targets).
#[derive(Debug, Clone)]
pub struct Skeleton {
    pub paged: Vec<u8>,
    pub section_stories: Vec<String>,
}

/// Story id of section `k` (0-based).
pub fn section_story_id(k: usize) -> String {
    format!("docx_s{k}")
}

fn twips(v: i32) -> f32 {
    v as f32 / 20.0
}

/// The native skeleton document for `doc`'s sections.
pub fn skeleton_document(doc: &DocxDocument) -> Document {
    let mut spreads = Vec::new();
    let mut stories = Vec::new();
    for (k, sec) in doc.sections.iter().enumerate() {
        let (w, h) = (twips(sec.page_width), twips(sec.page_height));
        let margins = MarginPreference {
            top: twips(sec.margin_top),
            bottom: twips(sec.margin_bottom),
            left: twips(sec.margin_left),
            right: twips(sec.margin_right),
            column_count: sec.columns.max(1),
            column_gutter: WORD_COLUMN_GAP_PT,
        };
        let page_id = format!("docx_p{k}");
        let spread_id = format!("docx_sp{k}");
        let story_id = section_story_id(k);

        let page = Page::new(
            page_id.clone(),
            Bounds {
                top: 0.0,
                left: 0.0,
                bottom: h,
                right: w,
            },
        );
        let mut frame = TextFrame::new(
            format!("docx_f{k}"),
            Some(story_id.clone()),
            Bounds {
                top: margins.top,
                left: margins.left,
                bottom: h - margins.bottom,
                right: w - margins.right,
            },
        );
        // Word's line-box fit, on this page and every page it grows.
        frame.first_baseline_offset = Some(FirstBaselineOffset::LeadingOffset);
        frame.inset_spacing = Some([0.0; 4]);
        if margins.column_count > 1 {
            frame.column_count = Some(margins.column_count);
            frame.column_gutter = Some(margins.column_gutter);
        }

        let mut page_margins = HashMap::new();
        page_margins.insert(page_id, margins);
        spreads.push(ParsedSpread {
            src: format!("Spreads/Spread_{spread_id}.xml"),
            spread: paged_model::Spread {
                self_id: Some(spread_id),
                pages: vec![page],
                text_frames: vec![frame],
                page_margins,
                ..Default::default()
            },
        });
        stories.push(ParsedStory {
            src: format!("Stories/Story_{story_id}.xml"),
            self_id: story_id,
            story: Story {
                paragraphs: vec![Paragraph::default()],
                grow: Some(FlowGrowRule {
                    copy_frame_options: true,
                    ..Default::default()
                }),
                ..Default::default()
            },
        });
    }
    let designmap = DesignMap {
        spreads: spreads
            .iter()
            .map(|s| SpreadRef { src: s.src.clone() })
            .collect(),
        stories: stories
            .iter()
            .map(|s| StoryRef { src: s.src.clone() })
            .collect(),
        ..Default::default()
    };
    let mut document = Document {
        designmap,
        spreads,
        stories,
        ..Default::default()
    };
    document.rebuild_indexes();
    document
}

/// The skeleton packaged as `.paged`, named `name`.
pub fn skeleton(doc: &DocxDocument, name: &str) -> Result<Skeleton, String> {
    let document = skeleton_document(doc);
    let (w, h) = doc
        .sections
        .first()
        .map(|s| (twips(s.page_width), twips(s.page_height)))
        .unwrap_or((612.0, 792.0));
    let paged = paged_store::package::wrap_document(&document, name, w, h)
        .map_err(|e| format!("package the skeleton: {e}"))?;
    Ok(Skeleton {
        paged,
        section_stories: (0..doc.sections.len()).map(section_story_id).collect(),
    })
}
