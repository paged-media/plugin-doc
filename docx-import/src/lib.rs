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
    Block, DocxDocument, Image, Justification, ListKind, ListMarker, ParaProps, Paragraph, Run,
    RunProps, Section, Style, StyleCatalog, StyleKind, TabStop, VertAlign,
};
use paged_ooxml::ooxmlsdk::schemas::schemas_openxmlformats_org_drawingml_2006_main as aml;
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
    let pkg = OpcPackage::read(bytes)?;

    // Resolve the main document part via the root relationships.
    let main_part = main_document_part(&pkg).unwrap_or_else(|| "word/document.xml".to_string());
    let doc_bytes = pkg.require(&main_part)?;
    let wml_doc: wml::Document = parse_root(&main_part, doc_bytes)?;

    // The styles part is optional.
    let styles = styles_part(&pkg, &main_part)
        .and_then(|name| pkg.part(&name).map(|b| (name, b)))
        .and_then(|(name, b)| parse_root::<wml::Styles>(&name, b).ok())
        .map(|s| map_styles(&s))
        .unwrap_or_default();

    // numbering.xml (optional) -> a resolver from (numId, level) to a marker.
    let numbering = numbering_part(&pkg, &main_part)
        .and_then(|name| pkg.part(&name).map(|b| (name, b)))
        .and_then(|(name, b)| parse_root::<wml::Numbering>(&name, b).ok())
        .map(|n| NumberingTable::from_numbering(&n))
        .unwrap_or_default();

    // The main document's rels resolve `r:embed` image refs to media parts.
    let doc_rels = pkg
        .part(&rels::rels_part_name(&main_part))
        .map(rels::Relationships::parse)
        .unwrap_or_default();
    let ctx = ImportCtx {
        numbering,
        images: ImageResolver {
            rels: doc_rels,
            package: &pkg,
            base_dir: part_dir(&main_part).to_string(),
        },
    };

    let (body, sections) = wml_doc
        .body
        .as_deref()
        .map(|b| map_body(b, &ctx))
        .unwrap_or_else(|| (Vec::new(), Vec::new()));

    Ok(DocxDocument {
        body,
        styles,
        sections,
    })
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
    /// abstractNumId -> (level -> (format, lvlText)).
    abstract_levels: std::collections::HashMap<
        i32,
        std::collections::HashMap<u8, (wml::NumberFormatValues, Option<String>)>,
    >,
}

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
                levels.insert(ilvl, (fmt, text));
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
        let (fmt, text) = self.abstract_levels.get(abstract_id)?.get(&level)?;
        Some(match fmt {
            wml::NumberFormatValues::Bullet => ListMarker {
                kind: ListKind::Bullet,
                level,
                bullet_char: Some(normalize_bullet(text.as_deref())),
                number_format: None,
            },
            wml::NumberFormatValues::None => ListMarker {
                kind: ListKind::Bullet,
                level,
                bullet_char: Some("\u{2022}".into()),
                number_format: None,
            },
            other => ListMarker {
                kind: ListKind::Numbered,
                level,
                bullet_char: None,
                number_format: Some(numbering_sample(other).to_string()),
            },
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

fn map_body(body: &wml::Body, ctx: &ImportCtx) -> (Vec<Block>, Vec<Section>) {
    let mut blocks = Vec::new();
    for choice in &body.body_choice {
        match choice {
            wml::BodyChoice::Paragraph(p) => blocks.push(Block::Paragraph(map_paragraph(p, ctx))),
            wml::BodyChoice::Table(t) => {
                blocks.push(Block::Table(map_table(t, ctx)));
            }
            _ => {}
        }
    }
    let sections = body
        .section_properties
        .as_deref()
        .map(|sp| vec![map_section(sp)])
        .unwrap_or_default();
    (blocks, sections)
}

fn map_paragraph(p: &wml::Paragraph, ctx: &ImportCtx) -> Paragraph {
    let (style_id, props) = match p.paragraph_properties.as_deref() {
        Some(pp) => (
            pp.paragraph_style_id.as_ref().map(|s| s.val.clone()),
            para_props(
                &pp.justification,
                &pp.indentation,
                &pp.spacing_between_lines,
                pp.keep_next.is_some(),
                pp.keep_lines.is_some(),
                &pp.tabs,
            ),
        ),
        None => (None, ParaProps::default()),
    };

    // Resolve w:numPr -> a list marker through numbering.xml.
    let list = p
        .paragraph_properties
        .as_deref()
        .and_then(|pp| pp.numbering_properties.as_deref())
        .and_then(|np| {
            let num_id = np.numbering_id.as_ref()?.val;
            let level = np
                .numbering_level_reference
                .as_ref()
                .map(|l| l.val as u8)
                .unwrap_or(0);
            ctx.numbering.resolve(num_id, level)
        });

    let mut runs = Vec::new();
    // Complex-field state: a stack so (rare) nested fields resolve correctly.
    // A run bearing a `fldChar`/`instrText` is a control run (no display text);
    // runs in an active HYPERLINK field's result become links.
    let mut fields: Vec<FieldFrame> = Vec::new();
    for choice in &p.paragraph_choice {
        match choice {
            wml::ParagraphChoice::WRun(r) => {
                if let Some(kind) = run_field_char(r) {
                    match kind {
                        wml::FieldCharValues::Begin => fields.push(FieldFrame::default()),
                        wml::FieldCharValues::Separate => {
                            if let Some(f) = fields.last_mut() {
                                f.result_url = parse_hyperlink_instr(&f.instruction);
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
                if let Some(f) = fields.last() {
                    if f.separated {
                        if let Some(url) = &f.result_url {
                            run.hyperlink = Some(url.clone());
                        }
                    }
                }
                runs.push(run);
            }
            wml::ParagraphChoice::Hyperlink(h) => {
                let target = ctx.hyperlink_target(h);
                for hc in &h.hyperlink_choice {
                    if let wml::HyperlinkChoice::WRun(r) = hc {
                        let mut run = map_run(r, ctx);
                        run.hyperlink = target.clone();
                        runs.push(run);
                    }
                }
            }
            // `w:fldSimple` — the single-element field form. If it's a
            // HYPERLINK, its inner display runs become links.
            wml::ParagraphChoice::SimpleField(fs) => {
                let url = parse_hyperlink_instr(&fs.instruction);
                for c in &fs.simple_field_choice {
                    if let wml::SimpleFieldChoice::WRun(r) = c {
                        let mut run = map_run(r, ctx);
                        run.hyperlink = url.clone();
                        runs.push(run);
                    }
                }
            }
            _ => {}
        }
    }

    Paragraph {
        style_id,
        props,
        runs,
        list,
    }
}

fn map_table(t: &wml::Table, ctx: &ImportCtx) -> docx_core::Table {
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
    for tc in &t.table_choice2 {
        if let wml::TableChoice2::TableRow(tr) = tc {
            let cells = tr
                .table_row_choice
                .iter()
                .filter_map(|rc| match rc {
                    wml::TableRowChoice::TableCell(c) => Some(map_cell(c, ctx)),
                    _ => None,
                })
                .collect();
            rows.push(docx_core::TableRow { cells });
        }
    }
    docx_core::Table {
        column_widths,
        rows,
    }
}

fn map_cell(c: &wml::TableCell, ctx: &ImportCtx) -> docx_core::TableCell {
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
    let paragraphs = c
        .table_cell_choice
        .iter()
        .filter_map(|cc| match cc {
            wml::TableCellChoice::Paragraph(p) => Some(map_paragraph(p, ctx)),
            _ => None,
        })
        .collect();
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
    let mut image = None;
    for c in &r.run_choice {
        match c {
            wml::RunChoice::Text(t) => {
                if let Some(s) = &t.0.xml_content {
                    text.push_str(s);
                }
            }
            wml::RunChoice::TabChar => text.push('\t'),
            wml::RunChoice::Break(_) | wml::RunChoice::CarriageReturn => text.push('\n'),
            wml::RunChoice::NoBreakHyphen => text.push('\u{2011}'),
            wml::RunChoice::Drawing(d) => {
                if image.is_none() {
                    image = map_drawing(d, ctx);
                }
            }
            _ => {}
        }
    }
    Run {
        style_id,
        props,
        text,
        image,
        hyperlink: None,
    }
}

/// Extract an [`Image`] from a `w:drawing`: the intrinsic extent + the picture
/// blip's `r:embed` rel id, walked typed (Inline/Anchor → a:graphic →
/// graphicData → pic:pic → blipFill → blip@embed), resolved to media bytes.
/// (Typed navigation avoids linking `ooxmlsdk`'s serializer, keeping the wasm
/// lean — `to_xml()` would pull in the whole schema's `write_to` codegen.)
fn map_drawing(d: &wml::Drawing, ctx: &ImportCtx) -> Option<Image> {
    let (width_emu, height_emu, graphic) = match d.drawing_choice.as_ref()? {
        wml::DrawingChoice::Inline(i) => (i.extent.cx, i.extent.cy, &i.graphic),
        wml::DrawingChoice::Anchor(a) => (a.extent.cx, a.extent.cy, &a.graphic),
    };
    let embed_id = blip_embed(graphic)?;
    let (bytes, mime) = ctx.images.resolve(embed_id)?;
    Some(Image {
        bytes,
        mime,
        width_emu,
        height_emu,
    })
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
                    false,
                    false,
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
                pp.keep_next.is_some(),
                pp.keep_lines.is_some(),
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

fn para_props(
    justification: &Option<wml::Justification>,
    indentation: &Option<wml::Indentation>,
    spacing: &Option<wml::SpacingBetweenLines>,
    keep_next: bool,
    keep_lines: bool,
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
    }
    if keep_next {
        p.keep_next = Some(true);
    }
    if keep_lines {
        p.keep_lines = Some(true);
    }
    if let Some(t) = tabs {
        for ts in &t.tab_stop {
            // A "clear" stop removes an inherited tab — no alignment, not carried.
            if let Some(alignment) = tab_alignment(&ts.val) {
                if let Some(position) = stwips(&ts.position) {
                    p.tabs.push(TabStop {
                        position,
                        alignment: Some(alignment),
                        // Leader glyphs (dot/hyphen) are a later-tier refinement.
                        leader: None,
                    });
                }
            }
        }
    }
    p
}

fn map_section(sp: &wml::SectionProperties) -> Section {
    let mut s = Section::default();
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
    }
    if let Some(cols) = &sp.columns {
        if let Some(n) = cols.column_count {
            s.columns = (n as u32).max(1);
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
