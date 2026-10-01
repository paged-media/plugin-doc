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

//! # docx-import — Tier-0 `.docx` → `docx-core`
//!
//! Reads the OPC package (`paged-ooxml`), resolves the main document + styles
//! parts through the `_rels` graph, parses them with the `ooxmlsdk` typed DOM,
//! and maps the enum-vector trees into the clean `docx-core` model. A **scanner,
//! not a validator**: a missing styles part yields an empty catalog; a body it
//! cannot fully understand yields the runs it can; only an unreadable container or
//! an unparseable main document part is a hard error.

use docx_core::{
    Block, BreakKind, CellPath, DocxDocument, Float, FloatPosition, FloatWrap, HeaderFooter,
    HeaderFooterKind, HeaderFooterRef, HeaderFooterSet, Image, Justification, LineRule,
    LineSpacing, ListKind, ListMarker, Note, NoteProps, NoteRestart, ParaProps, Paragraph,
    PositionalTab, PtabAlignment, PtabBase, Run, RunBreak, RunProps, RunSource, RunSymbol, Section,
    SectionKind, Style, StyleCatalog, StyleKind, TabStop, VertAlign, LINE_BREAK, SOFT_HYPHEN,
};
use paged_ooxml::ooxmlsdk::schemas::schemas_microsoft_com_office_word as w10;
use paged_ooxml::ooxmlsdk::schemas::schemas_microsoft_com_vml as vml;
use paged_ooxml::ooxmlsdk::schemas::schemas_openxmlformats_org_drawingml_2006_main as aml;
use paged_ooxml::ooxmlsdk::schemas::schemas_openxmlformats_org_drawingml_2006_wordprocessing_drawing as wp;
use paged_ooxml::ooxmlsdk::schemas::schemas_openxmlformats_org_wordprocessingml_2006_main as wml;
use paged_ooxml::ooxmlsdk::simple_type::{
    HpsMeasureValue, OnOffValue, SignedHpsMeasureValue, SignedTwipsMeasureValue, TwipsMeasureValue,
};
use paged_ooxml::{parse_root, part_dir, rels, resolve_target, OoxmlError, OpcPackage};

/// Import context threaded through the body mapping: the numbering resolver and
/// the image (media-part) resolver.
struct ImportCtx<'a> {
    numbering: NumberingTable,
    images: ImageResolver<'a>,
    /// Paragraph styles' own `w:numPr` and `basedOn`, for a paragraph whose
    /// list comes from its style (Word applies a style's numbering).
    style_lists: StyleLists,
    /// Legacy VML drawings (`w:pict`) that are not an inline picture, and so
    /// are not placed (floating shapes, text boxes, org charts…).
    unplaced_vml: std::cell::Cell<u32>,
}

/// A paragraph style's own `w:numPr` (`numId`, `ilvl`) and its `basedOn`.
type StyleList = (Option<(i32, u8)>, Option<String>);

/// Per paragraph style: its [`StyleList`], plus the document's default
/// paragraph style (`w:default="1"`).
#[derive(Default)]
struct StyleLists {
    styles: std::collections::HashMap<String, StyleList>,
    default_style: Option<String>,
}

impl StyleLists {
    fn from_styles(styles: &wml::Styles) -> Self {
        let mut out = StyleLists::default();
        for s in &styles.style {
            if !matches!(s.r#type, Some(wml::StyleValues::Paragraph) | None) {
                continue;
            }
            let Some(id) = s.style_id.clone() else {
                continue;
            };
            if s.default.is_some() && on(&s.default) {
                out.default_style = Some(id.clone());
            }
            let num = s
                .style_paragraph_properties
                .as_deref()
                .and_then(|pp| pp.numbering_properties.as_deref())
                .and_then(|np| {
                    let id = np.numbering_id.as_ref()?.val;
                    let lvl = np
                        .numbering_level_reference
                        .as_ref()
                        .map_or(0, |l| l.val as u8);
                    Some((id, lvl))
                });
            out.styles
                .insert(id, (num, s.based_on.as_ref().map(|b| b.val.clone())));
        }
        out
    }

    /// The numbering a paragraph in `style` inherits: the nearest style in
    /// its chain (the default paragraph style when it names none) with a
    /// `w:numPr`. `numId` 0 there means "no list" and ends the walk.
    fn inherited(&self, style: Option<&str>) -> Option<(i32, u8)> {
        let mut next = style.or(self.default_style.as_deref());
        for _ in 0..32 {
            let (num, based_on) = self.styles.get(next?)?;
            if let Some(n) = num {
                return (n.0 != 0).then_some(*n);
            }
            next = based_on.as_deref();
        }
        None
    }
}

impl ImportCtx<'_> {
    /// A `w:hyperlink`'s target: the external URL its `r:id` resolves to, or
    /// `#anchor` for an internal bookmark.
    fn hyperlink_target(&self, h: &wml::Hyperlink) -> Option<String> {
        if let Some(id) = &h.id {
            if let Some(rel) = self.images.rels.by_id(id) {
                return Some(rel.target.clone());
            }
        }
        h.anchor.as_ref().map(|a| format!("#{a}"))
    }
}

/// Resolves a drawing's `r:embed` rel id to its media bytes + MIME type.
struct ImageResolver<'a> {
    /// The main document part's relationships (where `r:embed` ids resolve).
    rels: rels::Relationships,
    /// The OPC package (holds the `word/media/…` byte parts).
    package: &'a OpcPackage,
    /// The directory of the main document part (for resolving media targets).
    base_dir: String,
}

impl ImageResolver<'_> {
    fn resolve(&self, embed_id: &str) -> Option<(Vec<u8>, String)> {
        let rel = self.rels.by_id(embed_id)?;
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.eq_ignore_ascii_case("External"))
        {
            return None; // external (linked) images aren't embedded bytes
        }
        let target = resolve_target(&self.base_dir, &rel.target);
        let bytes = self.package.part(&target)?.to_vec();
        Some((bytes, mime_for(&target)))
    }
}

/// A media part name -> MIME type (by extension).
fn mime_for(part_name: &str) -> String {
    let ext = part_name
        .rsplit('.')
        .next()
        .unwrap_or("")
        .to_ascii_lowercase();
    match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        _ => "application/octet-stream",
    }
    .to_string()
}

/// Import a `.docx`/`.dotx` package into the semantic model.
pub fn import_docx(bytes: &[u8]) -> Result<DocxDocument, OoxmlError> {
    import_docx_with_package(bytes).map(|(doc, _pkg, _main)| doc)
}

/// Import a `.docx`/`.dotx` and ALSO return the retained [`OpcPackage`] and the
/// main-document part name — the two things M2 edited save-back needs to patch
/// `word/document.xml` in place and re-emit every other part verbatim. The plain
/// [`import_docx`] discards both.
pub fn import_docx_with_package(
    bytes: &[u8],
) -> Result<(DocxDocument, OpcPackage, String), OoxmlError> {
    let pkg = OpcPackage::read(bytes)?;

    // Resolve the main document part via the root relationships.
    let main_part = main_document_part(&pkg).unwrap_or_else(|| "word/document.xml".to_string());
    let doc_bytes = pkg.require(&main_part)?;
    let wml_doc: wml::Document = parse_root(&main_part, doc_bytes)?;

    // The styles part is optional.
    let wml_styles = styles_part(&pkg, &main_part)
        .and_then(|name| pkg.part(&name).map(|b| (name, b)))
        .and_then(|(name, b)| parse_root::<wml::Styles>(&name, b).ok());
    let styles = wml_styles.as_ref().map(map_styles).unwrap_or_default();
    let style_lists = wml_styles
        .as_ref()
        .map(StyleLists::from_styles)
        .unwrap_or_default();

    // numbering.xml (optional) -> a resolver from (numId, level) to a marker.
    let numbering = numbering_part(&pkg, &main_part)
        .and_then(|name| pkg.part(&name).map(|b| (name, b)))
        .and_then(|(name, b)| parse_root::<wml::Numbering>(&name, b).ok())
        .map(|n| NumberingTable::from_numbering(&n))
        .unwrap_or_default();

    // Body mapping borrows the package (image resolution); scope the borrow so
    // the package can be moved into the return value once mapping is done.
    // settings.xml (optional): even/odd headers and the document-wide note
    // numbering.
    let settings = settings_part(&pkg, &main_part)
        .and_then(|name| pkg.part(&name).map(|b| (name, b)))
        .and_then(|(name, b)| parse_root::<wml::Settings>(&name, b).ok());
    let even_and_odd_headers = settings
        .as_ref()
        .and_then(|s| s.even_and_odd_headers.as_ref())
        .is_some_and(|e| on(&e.val));
    let footnote_props = settings
        .as_ref()
        .and_then(|s| s.footnote_document_wide_properties.as_deref())
        .map(|f| {
            note_props(
                f.footnote_position.as_ref().map(|p| p.val.to_string()),
                f.numbering_format.as_ref(),
                f.numbering_start.as_ref(),
                f.numbering_restart.as_ref(),
            )
        })
        .unwrap_or_default();
    let endnote_props = settings
        .as_ref()
        .and_then(|s| s.endnote_document_wide_properties.as_deref())
        .map(|e| {
            note_props(
                e.endnote_position.as_ref().map(|p| p.val.to_string()),
                e.numbering_format.as_ref(),
                e.numbering_start.as_ref(),
                e.numbering_restart.as_ref(),
            )
        })
        .unwrap_or_default();

    let auto_hyphenation = settings
        .as_ref()
        .and_then(|s| s.auto_hyphenation.as_ref())
        .is_some_and(|a| on(&a.val));

    let (body, sections, notes, headers_footers, unplaced_vml) = {
        let doc_rels = pkg
            .part(&rels::rels_part_name(&main_part))
            .map(rels::Relationships::parse)
            .unwrap_or_default();
        let ctx = ImportCtx {
            numbering,
            style_lists,
            unplaced_vml: std::cell::Cell::new(0),
            images: ImageResolver {
                rels: doc_rels,
                package: &pkg,
                base_dir: part_dir(&main_part).to_string(),
            },
        };
        let (body, mut sections, refs) = wml_doc
            .body
            .as_deref()
            .map(|b| map_body(b, &ctx))
            .unwrap_or_default();
        let notes = map_notes(&pkg, &main_part, &ctx);
        let headers_footers = map_headers_footers(&pkg, &main_part, &mut sections, &refs, &ctx);
        (
            body,
            sections,
            notes,
            headers_footers,
            ctx.unplaced_vml.get(),
        )
    };

    let doc = DocxDocument {
        notes,
        headers_footers,
        even_and_odd_headers,
        footnote_props,
        endnote_props,
        body,
        styles,
        sections,
        auto_hyphenation,
        unplaced_vml,
    };
    Ok((doc, pkg, main_part))
}

/// Parse the footnotes + endnotes parts (both optional) into `docx-core` notes.
/// Word ships two SEPARATOR pseudo-notes (`w:type` separator /
/// continuationSeparator, conventionally ids -1 and 0) that carry no real
/// content — those are skipped.
fn map_notes(pkg: &OpcPackage, main_part: &str, ctx: &ImportCtx) -> Vec<Note> {
    fn is_separator(t: &Option<wml::FootnoteEndnoteValues>) -> bool {
        matches!(
            t,
            Some(wml::FootnoteEndnoteValues::Separator)
                | Some(wml::FootnoteEndnoteValues::ContinuationSeparator)
                | Some(wml::FootnoteEndnoteValues::ContinuationNotice)
        )
    }
    let mut out = Vec::new();

    if let Some(name) = note_part(pkg, main_part, "/footnotes") {
        if let Some(parsed) = pkg
            .part(&name)
            .and_then(|b| parse_root::<wml::Footnotes>(&name, b).ok())
        {
            for f in &parsed.footnote {
                if is_separator(&f.r#type) {
                    continue;
                }
                out.push(Note {
                    id: f.id,
                    endnote: false,
                    paragraphs: note_paragraphs(&f.footnote_choice, ctx),
                });
            }
        }
    }
    if let Some(name) = note_part(pkg, main_part, "/endnotes") {
        if let Some(parsed) = pkg
            .part(&name)
            .and_then(|b| parse_root::<wml::Endnotes>(&name, b).ok())
        {
            for e in &parsed.endnote {
                if is_separator(&e.r#type) {
                    continue;
                }
                out.push(Note {
                    id: e.id,
                    endnote: true,
                    paragraphs: endnote_paragraphs(&e.endnote_choice, ctx),
                });
            }
        }
    }
    out
}

/// A footnote's body paragraphs. Cell provenance is `None` and the body ordinal
/// is 0: note content lives in its own part, so it is not save-back patchable.
fn note_paragraphs(choices: &[wml::FootnoteChoice], ctx: &ImportCtx) -> Vec<Paragraph> {
    choices
        .iter()
        .filter_map(|c| match c {
            wml::FootnoteChoice::Paragraph(p) => Some(map_paragraph(p, ctx, 0, None)),
            _ => None,
        })
        .collect()
}

/// An endnote's body paragraphs (the endnote choice is a distinct generated enum).
fn endnote_paragraphs(choices: &[wml::EndnoteChoice], ctx: &ImportCtx) -> Vec<Paragraph> {
    choices
        .iter()
        .filter_map(|c| match c {
            wml::EndnoteChoice::Paragraph(p) => Some(map_paragraph(p, ctx, 0, None)),
            _ => None,
        })
        .collect()
}

/// One `w:headerReference` / `w:footerReference`: footer?, kind, `r:id`.
type HeaderFooterRefIn = (bool, HeaderFooterKind, String);

/// A section's header/footer references, as written.
fn section_refs(sp: &wml::SectionProperties) -> Vec<HeaderFooterRefIn> {
    fn kind(v: wml::HeaderFooterValues) -> HeaderFooterKind {
        match v {
            wml::HeaderFooterValues::Default => HeaderFooterKind::Default,
            wml::HeaderFooterValues::First => HeaderFooterKind::First,
            wml::HeaderFooterValues::Even => HeaderFooterKind::Even,
        }
    }
    sp.section_properties_choice
        .iter()
        .map(|c| match c {
            wml::SectionPropertiesChoice::HeaderReference(h) => {
                (false, kind(h.r#type), h.id.clone())
            }
            wml::SectionPropertiesChoice::FooterReference(f) => {
                (true, kind(f.r#type), f.id.clone())
            }
        })
        .collect()
}

/// Parse every header/footer part ANY section references (thoughts ADR 033;
/// each section's references live on its own `sectPr`), once per part, and
/// record on each section which part it shows per kind. A section without a
/// reference of a kind shows the previous section's (Word's rule, ECMA-376
/// §17.10.5); in the first section that is a blank one.
fn map_headers_footers(
    pkg: &OpcPackage,
    main_part: &str,
    sections: &mut [Section],
    refs: &[Vec<HeaderFooterRefIn>],
    ctx: &ImportCtx,
) -> Vec<HeaderFooter> {
    let rels = pkg
        .part(&rels::rels_part_name(main_part))
        .map(rels::Relationships::parse)
        .unwrap_or_default();
    let base = part_dir(main_part);
    let mut out: Vec<HeaderFooter> = Vec::new();
    let mut prev = (HeaderFooterSet::default(), HeaderFooterSet::default());
    for (k, section) in sections.iter_mut().enumerate() {
        // Inherited unless the section says otherwise.
        let inherit = |r: Option<HeaderFooterRef>| {
            r.map(|r| HeaderFooterRef {
                inherited: true,
                ..r
            })
        };
        let mut headers = HeaderFooterSet {
            default: inherit(prev.0.default),
            first: inherit(prev.0.first),
            even: inherit(prev.0.even),
        };
        let mut footers = HeaderFooterSet {
            default: inherit(prev.1.default),
            first: inherit(prev.1.first),
            even: inherit(prev.1.even),
        };
        for (footer, kind, id) in refs.get(k).map(Vec::as_slice).unwrap_or_default() {
            let Some(rel) = rels.by_id(id) else { continue };
            let name = resolve_target(base, &rel.target);
            let index = match out.iter().position(|h| h.part == name) {
                Some(i) => i,
                None => {
                    let Some(bytes) = pkg.part(&name) else {
                        continue;
                    };
                    out.push(HeaderFooter {
                        footer: *footer,
                        kind: Some(
                            match kind {
                                HeaderFooterKind::Default => "default",
                                HeaderFooterKind::First => "first",
                                HeaderFooterKind::Even => "even",
                            }
                            .to_string(),
                        ),
                        part: name.clone(),
                        section: k,
                        paragraphs: header_footer_paragraphs(*footer, &name, bytes, ctx),
                    });
                    out.len() - 1
                }
            };
            let set = if *footer { &mut footers } else { &mut headers };
            let slot = match kind {
                HeaderFooterKind::Default => &mut set.default,
                HeaderFooterKind::First => &mut set.first,
                HeaderFooterKind::Even => &mut set.even,
            };
            *slot = Some(HeaderFooterRef {
                index,
                inherited: false,
            });
        }
        section.headers = headers;
        section.footers = footers;
        prev = (headers, footers);
    }
    out
}

/// A header (`w:hdr`) or footer (`w:ftr`) part's paragraphs. Not save-back
/// patchable yet (the body ordinal is 0, no cell provenance).
fn header_footer_paragraphs(
    footer: bool,
    name: &str,
    bytes: &[u8],
    ctx: &ImportCtx,
) -> Vec<Paragraph> {
    if footer {
        parse_root::<wml::Footer>(name, bytes)
            .ok()
            .map(|f| {
                f.footer_choice
                    .iter()
                    .filter_map(|c| match c {
                        wml::FooterChoice::Paragraph(p) => Some(map_paragraph(p, ctx, 0, None)),
                        _ => None,
                    })
                    .collect()
            })
            .unwrap_or_default()
    } else {
        parse_root::<wml::Header>(name, bytes)
            .ok()
            .map(|h| {
                h.header_choice
                    .iter()
                    .filter_map(|c| match c {
                        wml::HeaderChoice::Paragraph(p) => Some(map_paragraph(p, ctx, 0, None)),
                        _ => None,
                    })
                    .collect()
            })
            .unwrap_or_default()
    }
}

/// The settings part name (via the main document's `.rels`).
fn settings_part(pkg: &OpcPackage, main_part: &str) -> Option<String> {
    let rels_bytes = pkg.part(&rels::rels_part_name(main_part))?;
    let rels = rels::Relationships::parse(rels_bytes);
    let r = rels.by_type_suffix("/settings")?;
    Some(resolve_target(part_dir(main_part), &r.target))
}

/// `w:footnotePr` / `w:endnotePr` children -> [`NoteProps`].
fn note_props(
    pos: Option<String>,
    fmt: Option<&wml::NumberingFormat>,
    start: Option<&wml::NumberingStart>,
    restart: Option<&wml::NumberingRestart>,
) -> NoteProps {
    NoteProps {
        num_fmt: fmt.map(|f| f.val.to_string()),
        num_start: start.map(|s| u32::from(s.val)),
        num_restart: restart.map(|r| match r.val {
            wml::RestartNumberValues::Continuous => NoteRestart::Continuous,
            wml::RestartNumberValues::EachSection => NoteRestart::EachSection,
            wml::RestartNumberValues::EachPage => NoteRestart::EachPage,
        }),
        pos,
    }
}

/// A notes part name (via the main document's `.rels`), by relationship suffix.
fn note_part(pkg: &OpcPackage, main_part: &str, suffix: &str) -> Option<String> {
    let rels_bytes = pkg.part(&rels::rels_part_name(main_part))?;
    let rels = rels::Relationships::parse(rels_bytes);
    let r = rels.by_type_suffix(suffix)?;
    Some(resolve_target(part_dir(main_part), &r.target))
}

/// The numbering part name (via the main document's `.rels`).
fn numbering_part(pkg: &OpcPackage, main_part: &str) -> Option<String> {
    let rels_bytes = pkg.part(&rels::rels_part_name(main_part))?;
    let rels = rels::Relationships::parse(rels_bytes);
    let r = rels.by_type_suffix("/numbering")?;
    Some(resolve_target(part_dir(main_part), &r.target))
}

/// The main document part name (via `_rels/.rels` `officeDocument`).
fn main_document_part(pkg: &OpcPackage) -> Option<String> {
    let root_rels = pkg.part("_rels/.rels")?;
    let rels = rels::Relationships::parse(root_rels);
    let r = rels.by_type_suffix("/officeDocument")?;
    Some(resolve_target("", &r.target))
}

/// The styles part name (via the main document's `.rels`).
fn styles_part(pkg: &OpcPackage, main_part: &str) -> Option<String> {
    let rels_name = rels::rels_part_name(main_part);
    let rels_bytes = pkg.part(&rels_name)?;
    let rels = rels::Relationships::parse(rels_bytes);
    let r = rels.by_type_suffix("/styles")?;
    Some(resolve_target(part_dir(main_part), &r.target))
}

// ---------------------------------------------------------------------------
// Body

/// A resolver from a paragraph's `(numId, ilvl)` to a [`ListMarker`], built once
/// from `numbering.xml`. `numId -> abstractNumId -> level -> (numFmt, lvlText)`.
#[derive(Default)]
struct NumberingTable {
    /// numId -> abstractNumId.
    num_to_abstract: std::collections::HashMap<i32, i32>,
    /// abstractNumId -> (level -> (format, lvlText, the level's w:ind)).
    abstract_levels: std::collections::HashMap<i32, std::collections::HashMap<u8, Level>>,
}

/// One numbering level: its format, `w:lvlText` and indents.
type Level = (wml::NumberFormatValues, Option<String>, LevelIndent);

/// A numbering level's own `w:pPr/w:ind` (twips): left, first line, hanging.
type LevelIndent = (Option<i32>, Option<i32>, Option<i32>);

impl NumberingTable {
    fn from_numbering(n: &wml::Numbering) -> Self {
        let mut t = NumberingTable::default();
        for a in &n.abstract_num {
            let mut levels = std::collections::HashMap::new();
            for lvl in &a.level {
                let ilvl = lvl.level_index as u8;
                let fmt = lvl
                    .numbering_format
                    .as_ref()
                    .map(|f| f.val)
                    .unwrap_or(wml::NumberFormatValues::Decimal);
                let text = lvl.level_text.as_ref().and_then(|t| t.val.clone());
                let indent = lvl
                    .previous_paragraph_properties
                    .as_deref()
                    .and_then(|pp| pp.indentation.as_ref())
                    .map_or((None, None, None), |ind| {
                        (
                            ind.left.as_ref().or(ind.start.as_ref()).and_then(stwips),
                            ind.first_line.as_ref().and_then(twips_u),
                            ind.hanging.as_ref().and_then(stwips),
                        )
                    });
                levels.insert(ilvl, (fmt, text, indent));
            }
            t.abstract_levels.insert(a.abstract_number_id, levels);
        }
        for inst in &n.numbering_instance {
            t.num_to_abstract
                .insert(inst.number_id, inst.abstract_num_id.val);
        }
        t
    }

    fn resolve(&self, num_id: i32, level: u8) -> Option<ListMarker> {
        let abstract_id = self.num_to_abstract.get(&num_id)?;
        let (fmt, text, (left, first, hanging)) =
            self.abstract_levels.get(abstract_id)?.get(&level)?;
        let marker = |kind, bullet_char, number_format| ListMarker {
            kind,
            level,
            bullet_char,
            number_format,
            left_indent: *left,
            first_line_indent: *first,
            hanging_indent: *hanging,
        };
        Some(match fmt {
            wml::NumberFormatValues::Bullet => marker(
                ListKind::Bullet,
                Some(normalize_bullet(text.as_deref())),
                None,
            ),
            wml::NumberFormatValues::None => {
                marker(ListKind::Bullet, Some("\u{2022}".into()), None)
            }
            other => marker(
                ListKind::Numbered,
                None,
                Some(numbering_sample(other).to_string()),
            ),
        })
    }
}

/// Normalize a `w:lvlText` bullet glyph to a renderable Unicode character.
/// Word bullets are usually Symbol/Wingdings code points; map the common ones to
/// their Unicode equivalents so they render in the paragraph font.
fn normalize_bullet(text: Option<&str>) -> String {
    let first = text.and_then(|s| s.chars().next());
    match first {
        Some('\u{F0B7}') | Some('\u{2022}') | None => "\u{2022}".into(), // •
        Some('\u{F0A7}') | Some('\u{25AA}') => "\u{25AA}".into(),        // ▪
        Some('\u{F06E}') | Some('\u{25A0}') => "\u{25A0}".into(),        // ■
        Some('o') | Some('\u{25E6}') => "\u{25E6}".into(),               // ◦
        Some('\u{F0D8}') | Some('\u{2023}') => "\u{2023}".into(),        // ‣
        Some(c) if (c as u32) >= 0xF000 => "\u{2022}".into(),            // other symbol-font glyph
        Some(c) => c.to_string(),
    }
}

/// Word `w:numFmt` -> the IDML numbering-format sample the engine's
/// `format_number` reads (it keys off the head before the first comma).
fn numbering_sample(fmt: &wml::NumberFormatValues) -> &'static str {
    use wml::NumberFormatValues as F;
    match fmt {
        F::UpperRoman => "I, II, III, IV...",
        F::LowerRoman => "i, ii, iii, iv...",
        F::UpperLetter => "A, B, C, D...",
        F::LowerLetter => "a, b, c, d...",
        _ => "1, 2, 3, 4...",
    }
}

/// The body's blocks, its sections, and each section's header/footer
/// references (parallel to the sections).
type MappedBody = (Vec<Block>, Vec<Section>, Vec<Vec<HeaderFooterRefIn>>);

fn map_body(body: &wml::Body, ctx: &ImportCtx) -> MappedBody {
    let mut blocks = Vec::new();
    // Ordinal among the direct `<w:p>` children of `<w:body>` — the save-back
    // patcher's paragraph key. Only body paragraphs advance it (tables and any
    // dropped body child do not), matching "the Nth `<w:p>` under `<w:body>`".
    let mut para_ord = 0u32;
    // Ordinal among the direct `<w:tbl>` children of `<w:body>` (cell provenance).
    let mut table_ord = 0u32;
    // ADR 029 — a paragraph whose `pPr` carries a `sectPr` ENDS a section
    // (that `sectPr` describes the section it closes); the body-level
    // `sectPr` describes the last one.
    let mut sections = Vec::new();
    let mut refs = Vec::new();
    let mut section_start = 0usize;
    for choice in &body.body_choice {
        match choice {
            wml::BodyChoice::Paragraph(p) => {
                blocks.push(Block::Paragraph(map_paragraph(p, ctx, para_ord, None)));
                para_ord += 1;
                if let Some(sp) = p
                    .paragraph_properties
                    .as_deref()
                    .and_then(|pp| pp.section_properties.as_deref())
                {
                    let mut section = map_section(sp);
                    section.first_block = section_start;
                    sections.push(section);
                    refs.push(section_refs(sp));
                    section_start = blocks.len();
                }
            }
            wml::BodyChoice::Table(t) => {
                blocks.push(Block::Table(map_table(t, ctx, table_ord)));
                table_ord += 1;
            }
            _ => {}
        }
    }
    if let Some(sp) = body.section_properties.as_deref() {
        let mut section = map_section(sp);
        section.first_block = section_start;
        sections.push(section);
        refs.push(section_refs(sp));
    }
    (blocks, sections, refs)
}

fn map_paragraph(
    p: &wml::Paragraph,
    ctx: &ImportCtx,
    para_ord: u32,
    source_cell: Option<CellPath>,
) -> Paragraph {
    let (style_id, props) = match p.paragraph_properties.as_deref() {
        Some(pp) => (
            pp.paragraph_style_id.as_ref().map(|s| s.val.clone()),
            para_props(
                &pp.justification,
                &pp.indentation,
                &pp.spacing_between_lines,
                Keeps {
                    next: pp.keep_next.as_ref().map(|v| on(&v.val)),
                    lines: pp.keep_lines.as_ref().map(|v| on(&v.val)),
                    widow_control: pp.widow_control.as_ref().map(|v| on(&v.val)),
                },
                pp.page_break_before.as_ref().map(|v| on(&v.val)),
                &pp.tabs,
            ),
        ),
        None => (None, ParaProps::default()),
    };

    // Resolve w:numPr -> a list marker through numbering.xml. With no numPr
    // of its own the paragraph takes its style chain's (Word applies a
    // paragraph style's numbering); its own numPr, even `numId` 0, wins.
    let own_num = p
        .paragraph_properties
        .as_deref()
        .and_then(|pp| pp.numbering_properties.as_deref());
    let list = match own_num {
        Some(np) => np.numbering_id.as_ref().and_then(|id| {
            let level = np
                .numbering_level_reference
                .as_ref()
                .map(|l| l.val as u8)
                .unwrap_or(0);
            ctx.numbering.resolve(id.val, level)
        }),
        None => ctx
            .style_lists
            .inherited(
                p.paragraph_properties
                    .as_deref()
                    .and_then(|pp| pp.paragraph_style_id.as_ref())
                    .map(|s| s.val.as_str()),
            )
            .and_then(|(id, level)| ctx.numbering.resolve(id, level)),
    };

    let mut runs = Vec::new();
    // Complex-field state: a stack so (rare) nested fields resolve correctly.
    // A run bearing a `fldChar`/`instrText` is a control run (no display text);
    // runs in an active HYPERLINK field's result become links.
    let mut fields: Vec<FieldFrame> = Vec::new();
    // Ordinal of the source `<w:r>` among the paragraph's DIRECT children — the
    // key the save-back patcher locates on. EVERY `ParagraphChoice::WRun`
    // consumes one (control/field-char runs included), so the count tracks the
    // real XML `<w:r>` child index; `<w:hyperlink>`/`<w:fldSimple>`-wrapped runs
    // sit on a different path and are marked non-patchable instead.
    let mut wrun_ord = 0u32;
    // Ordinals of the WRAPPER children (`<w:hyperlink>` / `<w:fldSimple>`), so a
    // wrapped run keeps a locatable address of its own.
    let mut link_ord = 0u32;
    let mut field_ord = 0u32;
    for choice in &p.paragraph_choice {
        match choice {
            wml::ParagraphChoice::WRun(r) => {
                let ord = wrun_ord;
                wrun_ord += 1;
                if let Some(kind) = run_field_char(r) {
                    match kind {
                        wml::FieldCharValues::Begin => fields.push(FieldFrame::default()),
                        wml::FieldCharValues::Separate => {
                            if let Some(f) = fields.last_mut() {
                                f.result_url = parse_hyperlink_instr(&f.instruction);
                                f.result_field = field_name(&f.instruction);
                                f.separated = true;
                            }
                        }
                        wml::FieldCharValues::End => {
                            fields.pop();
                        }
                    }
                    continue;
                }
                if let Some(instr) = run_instr_text(r) {
                    if let Some(f) = fields.last_mut() {
                        f.instruction.push_str(&instr);
                    }
                    continue;
                }
                let mut run = map_run(r, ctx);
                run.source = Some(RunSource::DirectRun(ord));
                if let Some(f) = fields.last() {
                    if f.separated {
                        if let Some(url) = &f.result_url {
                            run.hyperlink = Some(url.clone());
                        }
                        run.field = f.result_field.clone();
                    }
                }
                runs.push(run);
            }
            wml::ParagraphChoice::Hyperlink(h) => {
                let target = ctx.hyperlink_target(h);
                let mut inner = 0u32;
                for hc in &h.hyperlink_choice {
                    if let wml::HyperlinkChoice::WRun(r) = hc {
                        let mut run = map_run(r, ctx);
                        run.hyperlink = target.clone();
                        run.source = Some(RunSource::Hyperlink {
                            link_ord,
                            run_ord: inner,
                        });
                        inner += 1;
                        runs.push(run);
                    }
                }
                link_ord += 1;
            }
            // `w:fldSimple` — the single-element field form. If it's a
            // HYPERLINK, its inner display runs become links.
            wml::ParagraphChoice::SimpleField(fs) => {
                let url = parse_hyperlink_instr(&fs.instruction);
                let mut inner = 0u32;
                for c in &fs.simple_field_choice {
                    if let wml::SimpleFieldChoice::WRun(r) = c {
                        let mut run = map_run(r, ctx);
                        run.hyperlink = url.clone();
                        run.field = field_name(&fs.instruction);
                        run.source = Some(RunSource::Field {
                            field_ord,
                            run_ord: inner,
                        });
                        inner += 1;
                        runs.push(run);
                    }
                }
                field_ord += 1;
            }
            _ => {}
        }
    }

    Paragraph {
        style_id,
        props,
        runs,
        list,
        source_para_ord: para_ord,
        source_cell,
    }
}

fn map_table(t: &wml::Table, ctx: &ImportCtx, table_ord: u32) -> docx_core::Table {
    let column_widths = t
        .table_grid
        .as_deref()
        .map(|g| {
            g.grid_column
                .iter()
                .filter_map(|c| c.width.as_ref().and_then(twips_u))
                .collect()
        })
        .unwrap_or_default();

    let mut rows = Vec::new();
    let mut row_ord = 0u32;
    for tc in &t.table_choice2 {
        if let wml::TableChoice2::TableRow(tr) = tc {
            let mut cell_ord = 0u32;
            let mut cells = Vec::new();
            for rc in &tr.table_row_choice {
                if let wml::TableRowChoice::TableCell(c) = rc {
                    cells.push(map_cell(c, ctx, table_ord, row_ord, cell_ord));
                    cell_ord += 1;
                }
            }
            rows.push(docx_core::TableRow { cells });
            row_ord += 1;
        }
    }
    docx_core::Table {
        column_widths,
        rows,
    }
}

fn map_cell(
    c: &wml::TableCell,
    ctx: &ImportCtx,
    table_ord: u32,
    row: u32,
    cell: u32,
) -> docx_core::TableCell {
    let props = c.table_cell_properties.as_deref();
    let grid_span = props
        .and_then(|p| p.grid_span.as_ref())
        .map(|g| g.val.max(1) as u32)
        .unwrap_or(1);
    let v_merge = match props.and_then(|p| p.vertical_merge.as_ref()) {
        None => docx_core::VMerge::None,
        // A `w:vMerge` with `val="restart"` starts a span; anything else
        // (`val="continue"` or an absent val) continues it.
        Some(vm) => match vm.val {
            Some(wml::MergedCellValues::Restart) => docx_core::VMerge::Restart,
            _ => docx_core::VMerge::Continue,
        },
    };
    // Cell paragraphs carry a `CellPath` (tbl/tr/tc/p ordinals) so save-back can
    // locate them; `source_para_ord` is meaningless here.
    let mut paragraphs = Vec::new();
    let mut para = 0u32;
    for cc in &c.table_cell_choice {
        if let wml::TableCellChoice::Paragraph(p) = cc {
            let path = CellPath {
                table_ord,
                row,
                cell,
                para,
            };
            paragraphs.push(map_paragraph(p, ctx, 0, Some(path)));
            para += 1;
        }
    }
    docx_core::TableCell {
        paragraphs,
        grid_span,
        v_merge,
    }
}

/// One frame of a complex field (`w:fldChar begin … instrText … separate …
/// result … end`). Instruction text accumulates between `begin` and `separate`;
/// runs between `separate` and `end` are the field result.
#[derive(Default)]
struct FieldFrame {
    instruction: String,
    separated: bool,
    /// The resolved external URL if this is a `HYPERLINK` field (else `None`).
    result_url: Option<String>,
    /// The field's NAME (instruction's first token, upper-cased) — `PAGE`,
    /// `DATE`, `REF`, … Stamped on the result runs so lowering can report which
    /// field kinds became frozen text.
    result_field: Option<String>,
}

/// The `w:fldChar` type carried by a run, if any (a control run — no display text).
fn run_field_char(r: &wml::Run) -> Option<wml::FieldCharValues> {
    r.run_choice.iter().find_map(|c| match c {
        wml::RunChoice::FieldChar(fc) => Some(fc.field_char_type),
        _ => None,
    })
}

/// The concatenated `w:instrText` (field-code) text carried by a run, if any.
fn run_instr_text(r: &wml::Run) -> Option<String> {
    let mut s = String::new();
    for c in &r.run_choice {
        if let wml::RunChoice::FieldCode(fc) = c {
            if let Some(t) = &fc.0.xml_content {
                s.push_str(t);
            }
        }
    }
    (!s.is_empty()).then_some(s)
}

/// Parse a field instruction and return the EXTERNAL URL when it is a
/// `HYPERLINK "url"` field. An internal `HYPERLINK \l "bookmark"` link returns
/// `None` (styled-only, mirroring the `#anchor` case — the core hyperlink door
/// registers URL destinations, not text anchors). Word splits the instruction
/// across several `w:instrText` runs, so this parses the accumulated string.
/// A field instruction's NAME: its first whitespace-separated token, upper-cased
/// (`" PAGE  \\* MERGEFORMAT "` → `PAGE`). `None` for an empty instruction.
fn field_name(instr: &str) -> Option<String> {
    instr
        .split_whitespace()
        .next()
        .map(|t| t.to_ascii_uppercase())
        .filter(|t| !t.is_empty())
}

fn parse_hyperlink_instr(instr: &str) -> Option<String> {
    let rest = instr.trim().strip_prefix("HYPERLINK")?;
    let quote = rest.find('"')?;
    // A `\l` switch BEFORE the first quoted argument means an internal
    // bookmark target (no external URL); skip it.
    if rest[..quote].split_whitespace().any(|t| t == "\\l") {
        return None;
    }
    let after = &rest[quote + 1..];
    let end = after.find('"')?;
    let url = &after[..end];
    (!url.is_empty()).then(|| url.to_string())
}

fn map_run(r: &wml::Run, ctx: &ImportCtx) -> Run {
    let mut props = RunProps::default();
    let mut style_id = None;
    if let Some(rpr) = r.run_properties.as_deref() {
        for c in &rpr.run_properties_choice {
            apply_run_property_choice(c, &mut props, &mut style_id);
        }
    }
    let mut text = String::new();
    let mut breaks = Vec::new();
    let mut images = Vec::new();
    let mut other_drawings = 0u32;
    let mut note_ref = None;
    let mut ptabs = Vec::new();
    let mut symbols = Vec::new();
    for c in &r.run_choice {
        match c {
            wml::RunChoice::Text(t) => {
                if let Some(s) = &t.0.xml_content {
                    text.push_str(s);
                }
            }
            wml::RunChoice::TabChar => text.push('\t'),
            // ADR 028/029 — a page or column break is a pagination
            // instruction, not text: it is recorded at its char offset. A
            // typeless or `textWrapping` break (and `w:cr`) is a line break
            // WITHIN the paragraph: U+2028, the engine's forced line break
            // (one char in the contiguous style space, 3 UTF-8 bytes in
            // insertText's), never `\n`, which would split the paragraph.
            wml::RunChoice::Break(b) => match b.r#type {
                Some(wml::BreakValues::Page) => breaks.push(RunBreak {
                    at: text.chars().count(),
                    kind: BreakKind::Page,
                }),
                Some(wml::BreakValues::Column) => breaks.push(RunBreak {
                    at: text.chars().count(),
                    kind: BreakKind::Column,
                }),
                _ => text.push(LINE_BREAK),
            },
            wml::RunChoice::CarriageReturn => text.push(LINE_BREAK),
            wml::RunChoice::NoBreakHyphen => text.push('\u{2011}'),
            // Word's optional hyphen: the one place its word may break.
            wml::RunChoice::SoftHyphen => text.push(SOFT_HYPHEN),
            // A symbol-font character: its Unicode equivalent, where it has
            // one; recorded either way, so the lowering can say what it
            // drew in the run's font and what it could not carry.
            wml::RunChoice::SymbolChar(sym) | wml::RunChoice::SymbolCharExt(sym) => {
                let code = sym.char.clone().unwrap_or_default();
                let font = sym.font.clone();
                let ch = docx_core::symbol_char(font.as_deref(), &code);
                symbols.push(RunSymbol {
                    at: text.chars().count(),
                    font,
                    code,
                    char: ch,
                });
                if let Some(ch) = ch {
                    text.push(ch);
                }
            }
            // An absolute-position tab: a tab in the text, with where it goes
            // kept for the lowering.
            wml::RunChoice::PositionalTab(p) => {
                use wml::AbsolutePositionTabAlignmentValues as A;
                use wml::AbsolutePositionTabLeaderCharValues as L;
                use wml::AbsolutePositionTabPositioningBaseValues as B;
                ptabs.push(PositionalTab {
                    at: text.chars().count(),
                    alignment: match p.alignment {
                        A::Left => PtabAlignment::Left,
                        A::Center => PtabAlignment::Center,
                        A::Right => PtabAlignment::Right,
                    },
                    relative_to: match p.relative_to {
                        B::Margin => PtabBase::Margin,
                        B::Indent => PtabBase::Indent,
                    },
                    leader: match p.leader {
                        L::None => None,
                        L::Dot => Some(".".into()),
                        L::Hyphen => Some("-".into()),
                        L::Underscore => Some("_".into()),
                        L::MiddleDot => Some("\u{00B7}".into()),
                    },
                });
                text.push('\t');
            }
            // Every picture in the run; anything else a drawing can be
            // (shape, text box, chart) is counted, not carried.
            wml::RunChoice::Drawing(d) => match map_drawing(d, ctx) {
                Some(img) => images.push(img),
                None => other_drawings += 1,
            },
            // A legacy VML picture (`w:pict`): a picture where Word gives it
            // room in the flow, else counted as not placed.
            wml::RunChoice::Picture(p) => match map_pict(p, ctx) {
                Some(img) => images.push(img),
                None => ctx.unplaced_vml.set(ctx.unplaced_vml.get() + 1),
            },
            // A footnote/endnote reference mark — the run carries no text; the
            // note body lives in the notes part, keyed by this id.
            wml::RunChoice::FootnoteReference(f) => note_ref = Some(f.id),
            wml::RunChoice::EndnoteReference(e) => note_ref = Some(e.id),
            _ => {}
        }
    }
    Run {
        style_id,
        props,
        text,
        breaks,
        ptabs,
        symbols,
        images,
        other_drawings,
        hyperlink: None,
        // The caller (map_paragraph) stamps the real provenance from context.
        source: None,
        note_ref,
        field: None,
    }
}

/// Extract an [`Image`] from a `w:drawing`: the intrinsic extent + the picture
/// blip's `r:embed` rel id, walked typed (Inline/Anchor → a:graphic →
/// graphicData → pic:pic → blipFill → blip@embed), resolved to media bytes.
/// (Typed navigation avoids linking `ooxmlsdk`'s serializer, keeping the wasm
/// lean — `to_xml()` would pull in the whole schema's `write_to` codegen.)
fn map_drawing(d: &wml::Drawing, ctx: &ImportCtx) -> Option<Image> {
    let (width_emu, height_emu, graphic, float) = match d.drawing_choice.as_ref()? {
        wml::DrawingChoice::Inline(i) => (i.extent.cx, i.extent.cy, &i.graphic, None),
        wml::DrawingChoice::Anchor(a) => (a.extent.cx, a.extent.cy, &a.graphic, Some(map_float(a))),
    };
    let embed_id = blip_embed(graphic)?;
    let (bytes, mime) = ctx.images.resolve(embed_id)?;
    Some(Image {
        bytes,
        mime,
        width_emu,
        height_emu,
        float,
    })
}

/// A `wp:anchor`'s position and wrap (thoughts ADR 035), as written.
fn map_float(a: &wp::Anchor) -> Float {
    let mut f = Float {
        dist_top: i64::from(a.distance_from_top.unwrap_or(0)),
        dist_bottom: i64::from(a.distance_from_bottom.unwrap_or(0)),
        dist_left: i64::from(a.distance_from_left.unwrap_or(0)),
        dist_right: i64::from(a.distance_from_right.unwrap_or(0)),
        behind_doc: a.behind_doc.as_bool(),
        allow_overlap: a.allow_overlap.as_bool(),
        layout_in_cell: a.layout_in_cell.as_bool(),
        locked: a.locked.as_bool(),
        relative_height: a.relative_height.unwrap_or(0),
        ..Float::default()
    };
    if a.simple_pos.is_some_and(|b| b.as_bool()) {
        f.simple_pos = a
            .simple_position
            .as_ref()
            .map(|p| (coordinate(&p.x), coordinate(&p.y)));
    }
    f.horizontal = a.horizontal_position.as_deref().map(|h| {
        use wp::HorizontalPositionChoice as C;
        let mut p = FloatPosition {
            relative_from: h.relative_from.to_string(),
            offset: None,
            align: None,
            percent: None,
        };
        match &h.horizontal_position_choice {
            Some(C::HorizontalAlignment(v)) => p.align = Some(v.to_string()),
            Some(C::PositionOffset(v)) => p.offset = Some(i64::from(*v)),
            Some(C::PercentagePositionHeightOffset(v)) => p.percent = percent(&v.to_string()),
            None => {}
        }
        p
    });
    f.vertical = a.vertical_position.as_deref().map(|v| {
        use wp::VerticalPositionChoice as C;
        let mut p = FloatPosition {
            relative_from: v.relative_from.to_string(),
            offset: None,
            align: None,
            percent: None,
        };
        match &v.vertical_position_choice {
            Some(C::VerticalAlignment(a)) => p.align = Some(a.to_string()),
            Some(C::PositionOffset(o)) => p.offset = Some(i64::from(*o)),
            Some(C::PercentagePositionVerticalOffset(o)) => p.percent = percent(&o.to_string()),
            None => {}
        }
        p
    });
    use wp::AnchorChoice as W;
    match &a.anchor_choice {
        None | Some(W::WrapNone) => f.wrap = FloatWrap::None,
        Some(W::WrapSquare(w)) => {
            f.wrap = FloatWrap::Square;
            f.wrap_text = Some(w.wrap_text.to_string());
            if let Some(v) = w.distance_from_top {
                f.dist_top = i64::from(v);
            }
            if let Some(v) = w.distance_from_bottom {
                f.dist_bottom = i64::from(v);
            }
            if let Some(v) = w.distance_from_left {
                f.dist_left = i64::from(v);
            }
            if let Some(v) = w.distance_from_right {
                f.dist_right = i64::from(v);
            }
        }
        Some(W::WrapTight(w)) => {
            f.wrap = FloatWrap::Tight;
            f.wrap_text = Some(w.wrap_text.to_string());
            if let Some(v) = w.distance_from_left {
                f.dist_left = i64::from(v);
            }
            if let Some(v) = w.distance_from_right {
                f.dist_right = i64::from(v);
            }
        }
        Some(W::WrapThrough(w)) => {
            f.wrap = FloatWrap::Through;
            f.wrap_text = Some(w.wrap_text.to_string());
            if let Some(v) = w.distance_from_left {
                f.dist_left = i64::from(v);
            }
            if let Some(v) = w.distance_from_right {
                f.dist_right = i64::from(v);
            }
        }
        Some(W::WrapTopBottom(w)) => {
            f.wrap = FloatWrap::TopAndBottom;
            if let Some(v) = w.distance_from_top {
                f.dist_top = i64::from(v);
            }
            if let Some(v) = w.distance_from_bottom {
                f.dist_bottom = i64::from(v);
            }
        }
    }
    f
}

/// A DrawingML coordinate as EMU (its XML text is an integer, or a
/// universal measure we do not expect in `wp:simplePos`).
fn coordinate<T: std::fmt::Display>(v: &T) -> i64 {
    v.to_string().trim().parse().unwrap_or(0)
}

/// A `wp14:pctPos*Offset` (thousandths of a percent) as written.
fn percent(s: &str) -> Option<i64> {
    s.trim().trim_end_matches('%').parse().ok()
}

/// An [`Image`] from a legacy VML picture (`w:pict` → `v:shape` →
/// `v:imagedata r:id`), Word 97–2003's picture: its size is the shape's CSS
/// `width`/`height`. An inline picture lowers, and so does a floating one
/// whose text wraps top and bottom (`w10:wrap type="topAndBottom"`): Word
/// gives it a band of its own, so the text resumes below it, as after an
/// inline picture (`fixtures/real-docx.word.json`, F01). Any other floating
/// VML, or VML that is not a picture, returns `None` and is counted as
/// unplaced.
fn map_pict(p: &wml::Picture, ctx: &ImportCtx) -> Option<Image> {
    p.picture_choice.iter().find_map(|c| {
        let wml::PictureChoice::Shape(shape) = c else {
            return None;
        };
        let style = shape.style.as_deref().unwrap_or("");
        let css = |key: &str| {
            style.split(';').find_map(|decl| {
                let (k, v) = decl.split_once(':')?;
                (k.trim() == key).then(|| v.trim())
            })
        };
        let own_band = shape.shape_choice.iter().any(|sc| {
            matches!(sc, vml::ShapeChoice::TextWrap(w)
                if matches!(w.r#type, Some(w10::WrapValues::TopAndBottom) | None))
        });
        if css("position") == Some("absolute") && !own_band {
            return None;
        }
        let rel_id = shape.shape_choice.iter().find_map(|sc| match sc {
            vml::ShapeChoice::ImageData(d) => d.relationship_id.as_deref(),
            _ => None,
        })?;
        let width_emu = css_length_emu(css("width")?)?;
        let height_emu = css_length_emu(css("height")?)?;
        let (bytes, mime) = ctx.images.resolve(rel_id)?;
        // Inline in the flow, as Word lays both kinds (a floating one's band
        // is its own line's worth of room).
        Some(Image {
            bytes,
            mime,
            width_emu,
            height_emu,
            float: None,
        })
    })
}

/// A VML/CSS length (`481.1pt`, `2in`, `3cm`, `20mm`, `96px`) in EMU.
fn css_length_emu(v: &str) -> Option<i64> {
    let split = v
        .find(|c: char| !(c.is_ascii_digit() || c == '.' || c == '-'))
        .unwrap_or(v.len());
    let (num, unit) = v.split_at(split);
    let n: f64 = num.parse().ok()?;
    let per = match unit.trim() {
        "pt" => 12700.0,
        "in" => 914400.0,
        "cm" => 360000.0,
        "mm" => 36000.0,
        "px" | "" => 9525.0,
        _ => return None,
    };
    Some((n * per).round() as i64)
}

/// The `r:embed` rel id of the first picture blip in a DrawingML graphic.
fn blip_embed(graphic: &aml::Graphic) -> Option<&str> {
    for choice in &graphic.graphic_data.graphic_data_choice {
        if let aml::GraphicDataChoice::Picture(pic) = choice {
            let blip = pic.blip_fill.as_deref()?.blip.as_deref()?;
            return blip.embed.as_deref();
        }
    }
    None
}

/// Apply one `w:rPr` child (the run's choice-vector form) to [`RunProps`].
fn apply_run_property_choice(
    c: &wml::RunPropertiesChoice,
    props: &mut RunProps,
    style_id: &mut Option<String>,
) {
    use wml::RunPropertiesChoice as C;
    match c {
        C::RunStyle(s) => *style_id = Some(s.val.clone()),
        C::RunFonts(f) => props.font = f.ascii.clone().or_else(|| props.font.take()),
        C::Bold(b) => props.bold = Some(on(&b.val)),
        C::Italic(i) => props.italic = Some(on(&i.val)),
        C::Caps(v) => props.caps = Some(on(&v.val)),
        C::SmallCaps(v) => props.small_caps = Some(on(&v.val)),
        C::Strike(v) => props.strike = Some(on(&v.val)),
        C::Color(col) => props.color = color_hex(&col.val),
        C::FontSize(sz) => props.size_half_pts = hps(&sz.val),
        C::Underline(u) => props.underline = Some(underline_on(u)),
        C::VerticalTextAlignment(v) => props.vert_align = Some(vert_align(&v.val)),
        C::Position(p) => props.baseline_half_pts = signed_hps(&p.val),
        _ => {}
    }
}

// ---------------------------------------------------------------------------
// Styles

fn map_styles(styles: &wml::Styles) -> StyleCatalog {
    let mut out = StyleCatalog::default();
    if let Some(dd) = styles.doc_defaults.as_deref() {
        // docDefaults — the document-wide base run/paragraph properties (Word's
        // Normal defaults: font, size, spacing). `docx-lower` turns a non-empty
        // Defaults into a base style every un-based style inherits from.
        if let Some(rpd) = dd.run_properties_default.as_deref() {
            if let Some(base) = rpd.run_properties_base_style.as_deref() {
                out.doc_defaults.run = run_props_base(base);
            }
        }
        if let Some(ppd) = dd.paragraph_properties_default.as_deref() {
            if let Some(base) = ppd.paragraph_properties_base_style.as_deref() {
                out.doc_defaults.para = para_props(
                    &base.justification,
                    &base.indentation,
                    &base.spacing_between_lines,
                    Keeps {
                        next: None,
                        lines: None,
                        widow_control: base.widow_control.as_ref().map(|v| on(&v.val)),
                    },
                    None,
                    &None,
                );
            }
        }
    }
    for s in &styles.style {
        if let Some(style) = map_style(s) {
            out.styles.push(style);
        }
    }
    out
}

/// Read the docDefaults run base style (`w:rPrDefault/w:rPr`) — a named-field
/// shape carrying the subset of run properties Word emits as defaults.
fn run_props_base(rpr: &wml::RunPropertiesBaseStyle) -> RunProps {
    let mut props = RunProps::default();
    if let Some(f) = &rpr.run_fonts {
        props.font = f.ascii.clone();
    }
    if let Some(b) = &rpr.bold {
        props.bold = Some(on(&b.val));
    }
    if let Some(i) = &rpr.italic {
        props.italic = Some(on(&i.val));
    }
    if let Some(c) = &rpr.color {
        props.color = color_hex(&c.val);
    }
    if let Some(sz) = &rpr.font_size {
        props.size_half_pts = hps(&sz.val);
    }
    if let Some(u) = &rpr.underline {
        props.underline = Some(underline_on(u));
    }
    props
}

fn map_style(s: &wml::Style) -> Option<Style> {
    let style_id = s.style_id.clone()?;
    let kind = match s.r#type {
        Some(wml::StyleValues::Paragraph) => StyleKind::Paragraph,
        Some(wml::StyleValues::Character) => StyleKind::Character,
        Some(wml::StyleValues::Table) => StyleKind::Table,
        Some(wml::StyleValues::Numbering) => StyleKind::Numbering,
        _ => StyleKind::Paragraph,
    };
    let para = s
        .style_paragraph_properties
        .as_deref()
        .map(|pp| {
            para_props(
                &pp.justification,
                &pp.indentation,
                &pp.spacing_between_lines,
                Keeps {
                    next: pp.keep_next.as_ref().map(|v| on(&v.val)),
                    lines: pp.keep_lines.as_ref().map(|v| on(&v.val)),
                    widow_control: pp.widow_control.as_ref().map(|v| on(&v.val)),
                },
                pp.page_break_before.as_ref().map(|v| on(&v.val)),
                &pp.tabs,
            )
        })
        .unwrap_or_default();
    let run = s
        .style_run_properties
        .as_deref()
        .map(run_props_named)
        .unwrap_or_default();

    Some(Style {
        style_id,
        name: s.style_name.as_ref().map(|n| n.val.clone()),
        kind,
        based_on: s.based_on.as_ref().map(|b| b.val.clone()),
        is_default: s.default.is_some() && on(&s.default),
        para,
        run,
    })
}

/// Read a style's `w:rPr` (the *named-field* form) into [`RunProps`].
fn run_props_named(rpr: &wml::StyleRunProperties) -> RunProps {
    let mut props = RunProps::default();
    if let Some(f) = &rpr.run_fonts {
        props.font = f.ascii.clone();
    }
    if let Some(b) = &rpr.bold {
        props.bold = Some(on(&b.val));
    }
    if let Some(i) = &rpr.italic {
        props.italic = Some(on(&i.val));
    }
    if let Some(v) = &rpr.caps {
        props.caps = Some(on(&v.val));
    }
    if let Some(v) = &rpr.small_caps {
        props.small_caps = Some(on(&v.val));
    }
    if let Some(v) = &rpr.strike {
        props.strike = Some(on(&v.val));
    }
    if let Some(c) = &rpr.color {
        props.color = color_hex(&c.val);
    }
    if let Some(sz) = &rpr.font_size {
        props.size_half_pts = hps(&sz.val);
    }
    if let Some(u) = &rpr.underline {
        props.underline = Some(underline_on(u));
    }
    if let Some(v) = &rpr.vertical_text_alignment {
        props.vert_align = Some(vert_align(&v.val));
    }
    if let Some(p) = &rpr.position {
        props.baseline_half_pts = signed_hps(&p.val);
    }
    props
}

// ---------------------------------------------------------------------------
// Paragraph properties (shared by w:pPr and w:style/w:pPr — same field types)

/// A paragraph's `w:keepNext`, `w:keepLines` and `w:widowControl`, each
/// `None` when absent and its `w:val` when present (`w:val="0"` turns an
/// inherited one off).
struct Keeps {
    next: Option<bool>,
    lines: Option<bool>,
    widow_control: Option<bool>,
}

fn para_props(
    justification: &Option<wml::Justification>,
    indentation: &Option<wml::Indentation>,
    spacing: &Option<wml::SpacingBetweenLines>,
    keeps: Keeps,
    page_break_before: Option<bool>,
    tabs: &Option<wml::Tabs>,
) -> ParaProps {
    let mut p = ParaProps::default();
    if let Some(j) = justification {
        p.justification = map_justification(&j.val);
    }
    if let Some(ind) = indentation {
        p.left_indent = ind.left.as_ref().or(ind.start.as_ref()).and_then(stwips);
        p.right_indent = ind.right.as_ref().or(ind.end.as_ref()).and_then(stwips);
        p.first_line_indent = ind.first_line.as_ref().and_then(twips_u);
        p.hanging_indent = ind.hanging.as_ref().and_then(stwips);
    }
    if let Some(sp) = spacing {
        p.space_before = sp.before.as_ref().and_then(stwips);
        p.space_after = sp.after.as_ref().and_then(stwips);
        if let Some(value) = sp.line.as_ref().and_then(stwips) {
            use wml::LineSpacingRuleValues as R;
            p.line_spacing = Some(LineSpacing {
                value,
                rule: match sp.line_rule {
                    Some(R::Exact) => LineRule::Exact,
                    Some(R::AtLeast) => LineRule::AtLeast,
                    _ => LineRule::Auto,
                },
            });
        }
    }
    if keeps.next == Some(true) {
        p.keep_next = Some(true);
    }
    p.keep_lines = keeps.lines;
    p.widow_control = keeps.widow_control;
    p.page_break_before = page_break_before;
    if let Some(t) = tabs {
        for ts in &t.tab_stop {
            // A "clear" stop (alignment `None`) removes the inherited stop at
            // its position; the lowering merges the chain.
            if let Some(position) = stwips(&ts.position) {
                p.tabs.push(TabStop {
                    position,
                    alignment: tab_alignment(&ts.val),
                    leader: ts.leader.as_ref().and_then(tab_leader),
                });
            }
        }
    }
    p
}

fn map_section(sp: &wml::SectionProperties) -> Section {
    let mut s = Section::default();
    if let Some(t) = &sp.section_type {
        use wml::SectionMarkValues as M;
        s.kind = match t.val {
            M::NextPage => SectionKind::NextPage,
            M::Continuous => SectionKind::Continuous,
            M::EvenPage => SectionKind::EvenPage,
            M::OddPage => SectionKind::OddPage,
            M::NextColumn => SectionKind::NextColumn,
        };
    }
    if let Some(ps) = &sp.page_size {
        if let Some(w) = ps.width.as_ref().and_then(twips_u) {
            s.page_width = w;
        }
        if let Some(h) = ps.height.as_ref().and_then(twips_u) {
            s.page_height = h;
        }
    }
    if let Some(pm) = &sp.page_margin {
        if let Some(v) = pm.top.as_ref().and_then(stwips) {
            s.margin_top = v;
        }
        if let Some(v) = pm.bottom.as_ref().and_then(stwips) {
            s.margin_bottom = v;
        }
        if let Some(v) = pm.left.as_ref().and_then(twips_u) {
            s.margin_left = v;
        }
        if let Some(v) = pm.right.as_ref().and_then(twips_u) {
            s.margin_right = v;
        }
        s.header_distance = pm.header.as_ref().and_then(twips_u);
        s.footer_distance = pm.footer.as_ref().and_then(twips_u);
    }
    s.title_page = sp.title_page.as_ref().is_some_and(|t| on(&t.val));
    if let Some(n) = &sp.page_number_type {
        s.page_number_start = n.start;
        s.page_number_format = n.format.map(|f| f.to_string());
    }
    if let Some(f) = sp.footnote_properties.as_deref() {
        s.footnote_props = note_props(
            f.footnote_position.as_ref().map(|p| p.val.to_string()),
            f.numbering_format.as_ref(),
            f.numbering_start.as_ref(),
            f.numbering_restart.as_ref(),
        );
    }
    if let Some(e) = sp.endnote_properties.as_deref() {
        s.endnote_props = note_props(
            e.endnote_position.as_ref().map(|p| p.val.to_string()),
            e.numbering_format.as_ref(),
            e.numbering_start.as_ref(),
            e.numbering_restart.as_ref(),
        );
    }
    if let Some(cols) = &sp.columns {
        if let Some(n) = cols.column_count {
            s.columns = (n as u32).max(1);
        }
        if let Some(v) = cols.space.as_ref().and_then(twips_u) {
            s.column_space = v;
        }
        let equal = cols.equal_width.is_none() || on(&cols.equal_width);
        if !equal && !cols.column.is_empty() {
            s.column_widths = cols
                .column
                .iter()
                .map(|c| {
                    (
                        c.width.as_ref().and_then(stwips).unwrap_or(0),
                        c.space.as_ref().and_then(stwips).unwrap_or(0),
                    )
                })
                .collect();
            s.columns = s.column_widths.len() as u32;
        }
    }
    s
}

// ---------------------------------------------------------------------------
// Value helpers

/// An `OnOff` toggle: absent `val` on a present element means "on".
fn on(v: &Option<OnOffValue>) -> bool {
    match v {
        Some(o) => matches!(
            o,
            OnOffValue::True | OnOffValue::On | OnOffValue::One | OnOffValue::Empty
        ),
        None => true,
    }
}

fn hps(v: &HpsMeasureValue) -> Option<u32> {
    match v {
        HpsMeasureValue::HalfPoints(n) => Some(*n as u32),
        HpsMeasureValue::UniversalMeasure(_) => None,
    }
}

/// Signed half-points (`w:position`), or `None` for a universal measure.
fn signed_hps(v: &SignedHpsMeasureValue) -> Option<i32> {
    match v {
        SignedHpsMeasureValue::HalfPoints(n) => Some(*n as i32),
        SignedHpsMeasureValue::UniversalMeasure(_) => None,
    }
}

fn stwips(v: &SignedTwipsMeasureValue) -> Option<i32> {
    match v {
        SignedTwipsMeasureValue::Twips(n) => Some(*n as i32),
        SignedTwipsMeasureValue::UniversalMeasure(_) => None,
    }
}

fn twips_u(v: &TwipsMeasureValue) -> Option<i32> {
    match v {
        TwipsMeasureValue::Twips(n) => Some(*n as i32),
        TwipsMeasureValue::UniversalMeasure(_) => None,
    }
}

/// `w:color/@w:val` normalized to `RRGGBB`, or `None` for `auto`/theme colors.
fn color_hex(val: &Option<String>) -> Option<String> {
    let v = val.as_ref()?;
    if v.eq_ignore_ascii_case("auto") || v.len() != 6 {
        return None;
    }
    Some(v.to_ascii_uppercase())
}

/// `w:u` is "on" unless it explicitly says `val="none"`. A present `<w:u/>` with
/// no `val`, or any real underline style (single/double/…), reads as on.
fn underline_on(u: &wml::Underline) -> bool {
    !matches!(u.val, Some(wml::UnderlineValues::None))
}

/// `w:tab/@w:val` -> a paged tab alignment string, or `None` for `clear`.
fn tab_alignment(v: &wml::TabStopValues) -> Option<String> {
    use wml::TabStopValues as T;
    Some(
        match v {
            T::Left | T::Start | T::Number => "left",
            T::Center => "center",
            T::Right | T::End => "right",
            T::Decimal => "decimal",
            T::Bar => "bar",
            T::Clear => return None,
        }
        .to_string(),
    )
}

/// `w:tab/@w:leader` as the native leader string.
fn tab_leader(v: &wml::TabStopLeaderCharValues) -> Option<String> {
    use wml::TabStopLeaderCharValues as L;
    Some(
        match v {
            L::None => return None,
            L::Dot => ".",
            L::Hyphen => "-",
            L::Underscore | L::Heavy => "_",
            L::MiddleDot => "\u{B7}",
        }
        .to_string(),
    )
}

fn vert_align(v: &wml::VerticalPositionValues) -> VertAlign {
    match v {
        wml::VerticalPositionValues::Superscript => VertAlign::Superscript,
        wml::VerticalPositionValues::Subscript => VertAlign::Subscript,
        wml::VerticalPositionValues::Baseline => VertAlign::Baseline,
    }
}

fn map_justification(v: &wml::JustificationValues) -> Option<Justification> {
    use wml::JustificationValues as J;
    Some(match v {
        J::Left => Justification::Left,
        J::Start => Justification::Start,
        J::Center => Justification::Center,
        J::Right => Justification::Right,
        J::End => Justification::End,
        J::Both => Justification::Both,
        J::Distribute | J::ThaiDistribute => Justification::Distribute,
        _ => return None,
    })
}
