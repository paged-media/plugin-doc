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

//! paged.doc — the page skeleton for DOCX standalone open (ADR 029), written
//! as a minimal **IDML package**.
//!
//! IDML is a public interchange format, like DOCX: this crate writes it with
//! `zip` and plain strings and has NO core dependency (the isolation
//! contract). The bundle opens the bytes through the SDK door
//! `host.nativeDocument.open`, whose importer reads IDML.
//!
//! A Word document is lowered as ONE STORY PER SECTION that starts a page:
//! each such section gets a page of its size with its margins
//! (`<MarginPreference>`), its margin box as the text frame, and a story with
//! one empty paragraph. The frame's first baseline is `LeadingOffset` with
//! zero insets: Word fits a line only when its whole line box fits, which is
//! what that expresses.
//!
//! IDML cannot carry the GROW RULE (the story adding pages while it oversets,
//! after its last frame, so section 1's pages land before section 2's, where
//! Word puts them). The bundle sets it on the wire after opening
//! (`setFlowGrowRule` with `copyFrameOptions`, protocol 64), then pours each
//! section's paragraphs into its story.
//!
//! A `continuous` / `nextColumn` section that Word continues on the same page
//! JOINS the story before it (no page of its own) when the native model can
//! say what Word does there; `docx_lower::sections` decides, measured against
//! Word (`docx-conformance/fixtures/continuous.word.json`), and the lowering
//! assigns each section the same story index, so the pour and save-back stay
//! aligned with this skeleton.

use std::io::{Cursor, Write};

use docx_core::{DocxDocument, Section};

/// Word's default space between columns (`w:cols/@w:space`, 720 twips).
const WORD_COLUMN_GAP_PT: f32 = 36.0;

/// The IDML package mimetype (stored, first entry: OCF).
pub const IDML_MIMETYPE: &str = "application/vnd.adobe.indesign-idml-package";

const PKG_NS: &str = "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging";
const XML_DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const NO_PARA_STYLE: &str = "ParagraphStyle/$ID/[No paragraph style]";
const NO_CHAR_STYLE: &str = "CharacterStyle/$ID/[No character style]";

/// The skeleton: IDML bytes for `host.nativeDocument.open`, and the story
/// ids, in story order (the pour's targets; sections that join a story share
/// its id, see [`story_sections`]).
#[derive(Debug, Clone)]
pub struct Skeleton {
    pub idml: Vec<u8>,
    pub section_stories: Vec<String>,
}

/// Id of native story `k` (0-based).
pub fn section_story_id(k: usize) -> String {
    format!("docx_s{k}")
}

/// Id of the text frame on story `k`'s authored page (the bundle binds the
/// Word source to story 0's frame).
pub fn section_frame_id(k: usize) -> String {
    format!("docx_f{k}")
}

/// Id of story `k`'s authored page.
pub fn section_page_id(k: usize) -> String {
    format!("docx_p{k}")
}

/// Id of the spread holding story `k`'s authored page.
pub fn section_spread_id(k: usize) -> String {
    format!("docx_sp{k}")
}

fn twips(v: i32) -> f32 {
    v as f32 / 20.0
}

/// A page's margins and column grid, in points (IDML `<MarginPreference>`).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Margins {
    pub top: f32,
    pub bottom: f32,
    pub left: f32,
    pub right: f32,
    pub column_count: u32,
    pub column_gutter: f32,
}

/// One skeleton story's authored page: its size, margins, and the frame on
/// its margin box. Points, page coordinates (origin top-left).
#[derive(Debug, Clone, PartialEq)]
pub struct SkeletonPage {
    pub story_id: String,
    pub page_id: String,
    pub spread_id: String,
    pub frame_id: String,
    pub width: f32,
    pub height: f32,
    pub margins: Margins,
}

impl SkeletonPage {
    /// The frame's bounds, `[top, left, bottom, right]`: the margin box.
    pub fn frame_bounds(&self) -> [f32; 4] {
        let m = &self.margins;
        [m.top, m.left, self.height - m.bottom, self.width - m.right]
    }
}

/// The first section of every native story, in story order: a section that
/// joins the story before it (`docx_lower::sections`) gets no page of its own.
pub fn story_sections(doc: &DocxDocument) -> Vec<&Section> {
    let placements = docx_lower::sections::place_sections(&doc.sections, &mut Vec::new());
    let mut out = Vec::new();
    for (sec, p) in doc.sections.iter().zip(&placements) {
        if p.story == out.len() {
            out.push(sec);
        }
    }
    out
}

/// The story ids of the skeleton, in story order (the pour's targets).
pub fn story_ids(doc: &DocxDocument) -> Vec<String> {
    (0..story_sections(doc).len())
        .map(section_story_id)
        .collect()
}

/// The skeleton's pages, one per native story, in story order.
pub fn skeleton_pages(doc: &DocxDocument) -> Vec<SkeletonPage> {
    story_sections(doc)
        .into_iter()
        .enumerate()
        .map(|(k, sec)| SkeletonPage {
            story_id: section_story_id(k),
            page_id: section_page_id(k),
            spread_id: section_spread_id(k),
            frame_id: section_frame_id(k),
            width: twips(sec.page_width),
            height: twips(sec.page_height),
            margins: Margins {
                top: twips(sec.margin_top),
                bottom: twips(sec.margin_bottom),
                left: twips(sec.margin_left),
                right: twips(sec.margin_right),
                column_count: sec.columns.max(1),
                column_gutter: WORD_COLUMN_GAP_PT,
            },
        })
        .collect()
}

/// The skeleton packaged as IDML, named `name`.
pub fn skeleton(doc: &DocxDocument, name: &str) -> Result<Skeleton, String> {
    let pages = skeleton_pages(doc);
    let idml = write_idml(&pages, name).map_err(|e| format!("package the skeleton: {e}"))?;
    Ok(Skeleton {
        idml,
        section_stories: pages.into_iter().map(|p| p.story_id).collect(),
    })
}

// ------------------------------------------------------------------ IDML

/// A number as IDML writes it: shortest form, no trailing `.0`.
fn num(v: f32) -> String {
    format!("{v}")
}

fn esc(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            c if (c as u32) < 0x20 && !matches!(c, '\t' | '\n' | '\r') => {}
            c => out.push(c),
        }
    }
    out
}

/// `<idPkg:{root} …>{body}</idPkg:{root}>` with the declaration.
fn pkg(root: &str, body: &str) -> String {
    format!(
        r#"{XML_DECL}<idPkg:{root} xmlns:idPkg="{PKG_NS}" DOMVersion="20.0">{body}</idPkg:{root}>"#
    )
}

fn spread_path(p: &SkeletonPage) -> String {
    format!("Spreads/Spread_{}.xml", p.spread_id)
}

fn story_path(p: &SkeletonPage) -> String {
    format!("Stories/Story_{}.xml", p.story_id)
}

fn container_xml() -> String {
    format!(
        r#"{XML_DECL}<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="designmap.xml" media-type="text/xml"/></rootfiles></container>"#
    )
}

fn designmap_xml(pages: &[SkeletonPage], name: &str) -> String {
    let story_list: Vec<&str> = pages.iter().map(|p| p.story_id.as_str()).collect();
    let mut parts = String::from(
        r#"<idPkg:Graphic src="Resources/Graphic.xml"/><idPkg:Fonts src="Resources/Fonts.xml"/><idPkg:Styles src="Resources/Styles.xml"/><idPkg:Preferences src="Resources/Preferences.xml"/><idPkg:Tags src="XML/Tags.xml"/>"#,
    );
    for p in pages {
        parts.push_str(&format!(r#"<idPkg:Spread src="{}"/>"#, spread_path(p)));
    }
    for p in pages {
        parts.push_str(&format!(r#"<idPkg:Story src="{}"/>"#, story_path(p)));
    }
    parts.push_str(r#"<idPkg:BackingStory src="XML/BackingStory.xml"/>"#);
    format!(
        r#"{XML_DECL}<?aid style="50" type="document" readerVersion="6.0" featureSet="257" product="20.0(32)"?><Document xmlns:idPkg="{PKG_NS}" DOMVersion="20.0" Self="d" StoryList="{}" Name="{}">{parts}</Document>"#,
        story_list.join(" "),
        esc(name),
    )
}

fn graphic_xml() -> String {
    pkg(
        "Graphic",
        r#"<Color Self="Color/Black" Model="Process" Space="CMYK" ColorValue="0 0 0 100" ColorOverride="Specialblack" Name="Black" ColorEditable="false" ColorRemovable="false" Visible="true"/><Color Self="Color/Paper" Model="Process" Space="CMYK" ColorValue="0 0 0 0" ColorOverride="Specialpaper" Name="Paper" ColorEditable="true" ColorRemovable="false" Visible="true"/><Swatch Self="Swatch/None" Name="None" ColorEditable="false" ColorRemovable="false" Visible="true"/>"#,
    )
}

fn styles_xml() -> String {
    pkg(
        "Styles",
        &format!(
            r#"<RootCharacterStyleGroup Self="u_rcsg"><CharacterStyle Self="{NO_CHAR_STYLE}" Name="$ID/[No character style]"/></RootCharacterStyleGroup><RootParagraphStyleGroup Self="u_rpsg"><ParagraphStyle Self="{NO_PARA_STYLE}" Name="$ID/[No paragraph style]"/></RootParagraphStyleGroup><RootObjectStyleGroup Self="u_rosg"><ObjectStyle Self="ObjectStyle/$ID/[None]" Name="$ID/[None]"/></RootObjectStyleGroup>"#
        ),
    )
}

/// Word documents are single-sided: no facing pages, the first story's
/// page as the document's default size. `PagesPerDocument` stays 1:
/// InDesign creates that many default pages BEFORE the package's spreads
/// and keeps all but one (measured on InDesign 20: `PagesPerDocument="2"`
/// opened the two-section skeleton with a stray default page first).
fn preferences_xml(pages: &[SkeletonPage]) -> String {
    let (w, h) = pages
        .first()
        .map(|p| (p.width, p.height))
        .unwrap_or((612.0, 792.0));
    pkg(
        "Preferences",
        &format!(
            r#"<DocumentPreference PageWidth="{}" PageHeight="{}" FacingPages="false" PagesPerDocument="1"/>"#,
            num(w),
            num(h),
        ),
    )
}

fn spread_xml(p: &SkeletonPage) -> String {
    let m = &p.margins;
    let [top, left, bottom, right] = p.frame_bounds();
    let anchors: String = [(left, top), (left, bottom), (right, bottom), (right, top)]
        .iter()
        .map(|(x, y)| {
            let a = format!("{} {}", num(*x), num(*y));
            format!(r#"<PathPointType Anchor="{a}" LeftDirection="{a}" RightDirection="{a}"/>"#)
        })
        .collect();
    let columns = if m.column_count > 1 {
        format!(
            r#" TextColumnCount="{}" TextColumnGutter="{}""#,
            m.column_count,
            num(m.column_gutter)
        )
    } else {
        String::new()
    };
    pkg(
        "Spread",
        &format!(
            concat!(
                r#"<Spread Self="{spread}" PageCount="1" BindingLocation="0" ShowMasterItems="true" AllowPageShuffle="false" ItemTransform="1 0 0 1 0 0">"#,
                r#"<Page Self="{page}" AppliedMaster="n" ItemTransform="1 0 0 1 0 0" GeometricBounds="0 0 {h} {w}" MasterPageTransform="1 0 0 1 0 0">"#,
                r#"<MarginPreference ColumnCount="{cc}" ColumnGutter="{cg}" Top="{mt}" Bottom="{mb}" Left="{ml}" Right="{mr}"/></Page>"#,
                r#"<TextFrame Self="{frame}" ParentStory="{story}" PreviousTextFrame="n" NextTextFrame="n" ContentType="TextType" AppliedObjectStyle="ObjectStyle/$ID/[None]" Visible="true" ItemTransform="1 0 0 1 0 0" FillColor="Swatch/None" StrokeColor="Swatch/None" StrokeWeight="0">"#,
                r#"<Properties><PathGeometry><GeometryPathType PathOpen="false"><PathPointArray>{anchors}</PathPointArray></GeometryPathType></PathGeometry></Properties>"#,
                r#"<TextFramePreference FirstBaselineOffset="LeadingOffset"{columns}><Properties><InsetSpacing type="unit">0</InsetSpacing></Properties></TextFramePreference>"#,
                r#"</TextFrame></Spread>"#,
            ),
            spread = p.spread_id,
            page = p.page_id,
            w = num(p.width),
            h = num(p.height),
            cc = m.column_count,
            cg = num(m.column_gutter),
            mt = num(m.top),
            mb = num(m.bottom),
            ml = num(m.left),
            mr = num(m.right),
            frame = p.frame_id,
            story = p.story_id,
            anchors = anchors,
            columns = columns,
        ),
    )
}

/// One empty paragraph: the pour fills it.
fn story_xml(p: &SkeletonPage) -> String {
    pkg(
        "Story",
        &format!(
            r#"<Story Self="{}"><ParagraphStyleRange AppliedParagraphStyle="{NO_PARA_STYLE}"><CharacterStyleRange AppliedCharacterStyle="{NO_CHAR_STYLE}"/></ParagraphStyleRange></Story>"#,
            p.story_id
        ),
    )
}

/// The IDML package for `pages`: `mimetype` (stored, first),
/// `META-INF/container.xml`, the designmap, the resources the importer
/// reads, a spread per page and a story per page.
pub fn write_idml(pages: &[SkeletonPage], name: &str) -> Result<Vec<u8>, String> {
    let mut entries: Vec<(String, String)> = vec![
        ("designmap.xml".into(), designmap_xml(pages, name)),
        ("META-INF/container.xml".into(), container_xml()),
        ("Resources/Graphic.xml".into(), graphic_xml()),
        ("Resources/Fonts.xml".into(), pkg("Fonts", "")),
        ("Resources/Styles.xml".into(), styles_xml()),
        ("Resources/Preferences.xml".into(), preferences_xml(pages)),
        (
            "XML/BackingStory.xml".into(),
            pkg(
                "BackingStory",
                r#"<XmlStory Self="BackingStory" AppliedXMLTag="XMLTag/$ID/Root"/>"#,
            ),
        ),
        (
            "XML/Tags.xml".into(),
            pkg("Tags", r#"<XMLTag Self="XMLTag/$ID/Root" Name="Root"/>"#),
        ),
    ];
    for p in pages {
        entries.push((spread_path(p), spread_xml(p)));
    }
    for p in pages {
        entries.push((story_path(p), story_xml(p)));
    }

    let mut out = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut out);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    let err = |e: &dyn std::fmt::Display| e.to_string();
    zip.start_file("mimetype", stored).map_err(|e| err(&e))?;
    zip.write_all(IDML_MIMETYPE.as_bytes())
        .map_err(|e| err(&e))?;
    for (path, xml) in entries {
        zip.start_file(path, deflated).map_err(|e| err(&e))?;
        zip.write_all(xml.as_bytes()).map_err(|e| err(&e))?;
    }
    zip.finish().map_err(|e| err(&e))?;
    Ok(out.into_inner())
}
