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

// ADR 029 — the section skeleton of docx_conformance::pagination_docx(), the
// document Word paginated (docx-conformance/fixtures/pagination.word.json),
// asserted on the IDML package this crate writes, read back with quick-xml.

use std::collections::HashMap;
use std::io::Read;

use docx_skeleton::{section_frame_id, section_story_id, skeleton, IDML_MIMETYPE};
use quick_xml::events::Event;

fn pagination() -> docx_core::DocxDocument {
    docx_import::import_docx(&docx_conformance::pagination_docx()).expect("import")
}

/// One XML element: name, attributes, and the text directly inside it.
#[derive(Debug, Clone)]
struct El {
    name: String,
    attrs: HashMap<String, String>,
    text: String,
}

fn elements(xml: &str) -> Vec<El> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut out: Vec<El> = Vec::new();
    let mut open: Vec<usize> = Vec::new();
    loop {
        match reader.read_event().expect("well-formed XML") {
            Event::Start(e) => {
                out.push(el(&e));
                open.push(out.len() - 1);
            }
            Event::Empty(e) => out.push(el(&e)),
            Event::End(_) => {
                open.pop();
            }
            Event::Text(t) => {
                if let Some(&i) = open.last() {
                    out[i].text.push_str(&t.decode().unwrap());
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    out
}

fn el(e: &quick_xml::events::BytesStart) -> El {
    El {
        name: String::from_utf8(e.name().as_ref().to_vec()).unwrap(),
        attrs: e
            .attributes()
            .map(|a| {
                let a = a.unwrap();
                (
                    String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap()
                        .into_owned(),
                )
            })
            .collect(),
        text: String::new(),
    }
}

/// The package: entries in order, and their contents.
struct Package {
    names: Vec<String>,
    first_stored: bool,
    parts: HashMap<String, String>,
}

fn open(bytes: &[u8]) -> Package {
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("zip");
    let first_stored = zip.by_index(0).unwrap().compression() == zip::CompressionMethod::Stored;
    let mut names = Vec::new();
    let mut parts = HashMap::new();
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).unwrap();
        let mut s = String::new();
        f.read_to_string(&mut s).unwrap();
        names.push(f.name().to_string());
        parts.insert(f.name().to_string(), s);
    }
    Package {
        names,
        first_stored,
        parts,
    }
}

impl Package {
    fn els(&self, path: &str) -> Vec<El> {
        elements(self.parts.get(path).unwrap_or_else(|| panic!("{path}")))
    }
}

fn find<'a>(els: &'a [El], name: &str) -> &'a El {
    els.iter()
        .find(|e| e.name == name)
        .unwrap_or_else(|| panic!("<{name}>"))
}

fn nums(s: &str) -> Vec<f32> {
    s.split_whitespace().map(|v| v.parse().unwrap()).collect()
}

/// What story `k`'s spread says: page bounds, margins, frame.
struct SpreadView {
    page: Vec<f32>,
    margin: HashMap<String, String>,
    frame: HashMap<String, String>,
    /// Frame `[top, left, bottom, right]` from its path anchors.
    frame_box: [f32; 4],
    frame_pref: HashMap<String, String>,
    inset: String,
}

fn spread(pkg: &Package, k: usize) -> SpreadView {
    let els = pkg.els(&format!("Spreads/Spread_docx_sp{k}.xml"));
    let page = find(&els, "Page");
    assert_eq!(page.attrs["Self"], format!("docx_p{k}"));
    assert_eq!(page.attrs["ItemTransform"], "1 0 0 1 0 0");
    let frame = find(&els, "TextFrame");
    assert_eq!(frame.attrs["ItemTransform"], "1 0 0 1 0 0");
    let pts: Vec<Vec<f32>> = els
        .iter()
        .filter(|e| e.name == "PathPointType")
        .map(|e| nums(&e.attrs["Anchor"]))
        .collect();
    assert_eq!(pts.len(), 4, "a rectangle");
    let xs = pts.iter().map(|p| p[0]);
    let ys = pts.iter().map(|p| p[1]);
    let frame_box = [
        ys.clone().fold(f32::MAX, f32::min),
        xs.clone().fold(f32::MAX, f32::min),
        ys.fold(f32::MIN, f32::max),
        xs.fold(f32::MIN, f32::max),
    ];
    SpreadView {
        page: nums(&page.attrs["GeometricBounds"]),
        margin: find(&els, "MarginPreference").attrs.clone(),
        frame: frame.attrs.clone(),
        frame_box,
        frame_pref: find(&els, "TextFramePreference").attrs.clone(),
        inset: find(&els, "InsetSpacing").text.clone(),
    }
}

#[test]
fn the_package_is_an_idml_package() {
    let sk = skeleton(&pagination(), "pagination.docx").expect("skeleton");
    let pkg = open(&sk.idml);
    assert_eq!(pkg.names[0], "mimetype", "OCF: mimetype first");
    assert!(pkg.first_stored, "OCF: mimetype stored");
    assert_eq!(pkg.parts["mimetype"], IDML_MIMETYPE);
    let container = pkg.els("META-INF/container.xml");
    assert_eq!(
        find(&container, "rootfile").attrs["full-path"],
        "designmap.xml"
    );

    // The designmap names every part this package holds, spreads and
    // stories in story order.
    let dm = pkg.els("designmap.xml");
    let doc = find(&dm, "Document");
    assert_eq!(doc.attrs["StoryList"], "docx_s0 docx_s1");
    assert_eq!(doc.attrs["Name"], "pagination.docx");
    let srcs: Vec<&str> = dm
        .iter()
        .filter(|e| e.name.starts_with("idPkg:"))
        .map(|e| e.attrs["src"].as_str())
        .collect();
    for src in &srcs {
        assert!(pkg.parts.contains_key(*src), "designmap names {src}");
    }
    let spreads: Vec<&&str> = srcs.iter().filter(|s| s.starts_with("Spreads/")).collect();
    assert_eq!(
        spreads,
        [
            &"Spreads/Spread_docx_sp0.xml",
            &"Spreads/Spread_docx_sp1.xml"
        ]
    );
    let stories: Vec<&&str> = srcs.iter().filter(|s| s.starts_with("Stories/")).collect();
    assert_eq!(
        stories,
        [&"Stories/Story_docx_s0.xml", &"Stories/Story_docx_s1.xml"]
    );

    // Single-sided, sized like section 1.
    let prefs = pkg.els("Resources/Preferences.xml");
    let dp = find(&prefs, "DocumentPreference");
    assert_eq!(dp.attrs["FacingPages"], "false");
    assert_eq!(
        dp.attrs["PagesPerDocument"], "1",
        "InDesign adds PagesPerDocument - 1 stray default pages"
    );
    assert_eq!(
        (
            dp.attrs["PageWidth"].as_str(),
            dp.attrs["PageHeight"].as_str()
        ),
        ("612", "792")
    );
    // Every style and swatch the parts reference is defined.
    let styles = &pkg.parts["Resources/Styles.xml"];
    for s in [
        "ParagraphStyle/$ID/[No paragraph style]",
        "CharacterStyle/$ID/[No character style]",
        "ObjectStyle/$ID/[None]",
    ] {
        assert!(styles.contains(&format!("Self=\"{s}\"")), "{s}");
    }
    assert!(pkg.parts["Resources/Graphic.xml"].contains("Self=\"Swatch/None\""));
}

#[test]
fn one_page_frame_and_empty_story_per_section() {
    let sk = skeleton(&pagination(), "pagination.docx").expect("skeleton");
    assert_eq!(
        sk.section_stories,
        vec![section_story_id(0), section_story_id(1)]
    );
    let pkg = open(&sk.idml);

    // Section 1: Letter, 1 in margins; section 2: A5 landscape, 0.5 in.
    let s1 = spread(&pkg, 0);
    assert_eq!(s1.page, [0.0, 0.0, 792.0, 612.0]);
    assert_eq!(s1.frame_box, [72.0, 72.0, 720.0, 540.0]);
    for side in ["Top", "Bottom", "Left", "Right"] {
        assert_eq!(s1.margin[side], "72");
    }
    assert_eq!(s1.margin["ColumnCount"], "1");
    let s2 = spread(&pkg, 1);
    assert!((s2.page[2] - 419.55).abs() < 0.01 && (s2.page[3] - 595.3).abs() < 0.01);
    assert_eq!((s2.frame_box[0], s2.frame_box[1]), (36.0, 36.0));
    assert!((s2.frame_box[3] - (595.3 - 36.0)).abs() < 0.01);

    for (k, s) in [s1, s2].iter().enumerate() {
        // Each section frame threads its own story, alone.
        assert_eq!(s.frame["Self"], section_frame_id(k));
        assert_eq!(s.frame["ParentStory"], section_story_id(k));
        assert_eq!(s.frame["PreviousTextFrame"], "n");
        assert_eq!(s.frame["NextTextFrame"], "n");
        // Word's line-box fit.
        assert_eq!(s.frame_pref["FirstBaselineOffset"], "LeadingOffset");
        assert_eq!(s.inset, "0");
        assert!(!s.frame_pref.contains_key("TextColumnCount"));

        let story = pkg.els(&format!("Stories/Story_{}.xml", section_story_id(k)));
        assert_eq!(find(&story, "Story").attrs["Self"], section_story_id(k));
        let paras = story
            .iter()
            .filter(|e| e.name == "ParagraphStyleRange")
            .count();
        assert_eq!(paras, 1, "one empty paragraph");
        assert!(!story.iter().any(|e| e.name == "Content"), "no text");
    }
}

#[test]
fn the_package_is_deterministic() {
    let a = skeleton(&pagination(), "pagination.docx").unwrap().idml;
    let b = skeleton(&pagination(), "pagination.docx").unwrap().idml;
    assert_eq!(a, b);
}

/// ADR 029 — continuous sections that Word keeps on the page they continue
/// share that page's story (docx_conformance::continuous_docx(), measured in
/// docx-conformance/fixtures/continuous.word.json): 17 Word sections, 11
/// stories, and the lowering assigns every section the story it pours into.
#[test]
fn continuous_sections_share_the_story_of_the_page_they_continue() {
    let docx = docx_import::import_docx(&docx_conformance::continuous_docx()).expect("import");
    assert_eq!(docx.sections.len(), 17);
    let sk = skeleton(&docx, "continuous.docx").expect("skeleton");
    let pkg = open(&sk.idml);
    let count =
        |pkg: &Package, prefix: &str| pkg.names.iter().filter(|n| n.starts_with(prefix)).count();
    assert_eq!(
        count(&pkg, "Spreads/"),
        8,
        "a page per story, not per section"
    );
    assert_eq!(count(&pkg, "Stories/"), 8);
    for k in 0..8 {
        let s = spread(&pkg, k);
        assert_eq!(s.frame["ParentStory"], section_story_id(k));
    }

    // Each story's page and frame are its FIRST section's: story 0 is A1's
    // (A2, A3 join it); story 1 is B1's page with the two columns the
    // lowering chose for B (B1 and B3 span them, ADR 029); story 3 is C1's
    // 0.5 in margins (C2's other margins become indents); story 6 is F2's
    // 6 in page.
    let s0 = spread(&pkg, 0);
    assert_eq!((s0.page[3], s0.page[2]), (360.0, 312.0));
    assert_eq!(s0.frame_box, [36.0, 36.0, 276.0, 324.0]);
    assert!(!s0.frame_pref.contains_key("TextColumnCount"));
    let s1 = spread(&pkg, 1);
    assert_eq!(s1.frame_pref["TextColumnCount"], "2");
    assert_eq!(s1.frame_pref["TextColumnGutter"], "36");
    assert_eq!(s1.margin["ColumnCount"], "2");
    assert_eq!(
        spread(&pkg, 2).frame_pref["TextColumnCount"],
        "2",
        "D1 + nextColumn D2"
    );
    let s3 = spread(&pkg, 3);
    assert_eq!((s3.frame_box[1], s3.frame_box[3]), (36.0, 324.0));
    assert_eq!(spread(&pkg, 6).page[3], 432.0);
    assert_eq!(spread(&pkg, 7).page[3], 360.0);

    // The pour's view agrees: one block group per skeleton story.
    let ir = docx_lower::lower(&docx);
    let story_of: Vec<usize> = ir.sections.iter().map(|s| s.story).collect();
    assert_eq!(
        story_of,
        vec![0, 0, 0, 1, 1, 1, 1, 2, 2, 3, 3, 3, 4, 4, 5, 6, 7]
    );
    assert_eq!(sk.section_stories, docx_skeleton::story_ids(&docx));
    assert_eq!(sk.section_stories.len(), story_of.last().unwrap() + 1);

    // An engine without span/split columns: B's four sections are four
    // pages, as before, each frame with its own section's columns.
    let old = docx_lower::sections::LowerOptions {
        mid_page_columns: false,
    };
    let sk = docx_skeleton::skeleton_with(&docx, "continuous.docx", old).expect("skeleton");
    let pkg = open(&sk.idml);
    assert_eq!(count(&pkg, "Spreads/"), 11);
    assert!(!spread(&pkg, 1).frame_pref.contains_key("TextColumnCount"));
    assert_eq!(spread(&pkg, 2).frame_pref["TextColumnCount"], "2");
    assert_eq!(
        sk.section_stories,
        docx_skeleton::story_ids_with(&docx, old)
    );
}
