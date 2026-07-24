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
    Block, DocxDocument, Justification, ParaProps, Paragraph, Run, RunProps, Section, Style,
    StyleCatalog, StyleKind, TabStop, VertAlign,
};
use paged_ooxml::ooxmlsdk::schemas::schemas_openxmlformats_org_wordprocessingml_2006_main as wml;
use paged_ooxml::ooxmlsdk::simple_type::{
    HpsMeasureValue, OnOffValue, SignedTwipsMeasureValue, TwipsMeasureValue,
};
use paged_ooxml::{parse_root, part_dir, rels, resolve_target, OoxmlError, OpcPackage};

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

    let (body, sections) = wml_doc
        .body
        .as_deref()
        .map(map_body)
        .unwrap_or_else(|| (Vec::new(), Vec::new()));

    Ok(DocxDocument {
        body,
        styles,
        sections,
    })
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

fn map_body(body: &wml::Body) -> (Vec<Block>, Vec<Section>) {
    let mut blocks = Vec::new();
    for choice in &body.body_choice {
        match choice {
            wml::BodyChoice::Paragraph(p) => blocks.push(Block::Paragraph(map_paragraph(p))),
            wml::BodyChoice::Table(_) => {
                // Tables are a Tier-2 construct; carry an empty stub so
                // docx-lower can emit an honest diagnostic without losing order.
                blocks.push(Block::Table(docx_core::Table::default()));
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

fn map_paragraph(p: &wml::Paragraph) -> Paragraph {
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

    let mut runs = Vec::new();
    for choice in &p.paragraph_choice {
        match choice {
            wml::ParagraphChoice::WRun(r) => runs.push(map_run(r)),
            wml::ParagraphChoice::Hyperlink(h) => {
                for hc in &h.hyperlink_choice {
                    if let wml::HyperlinkChoice::WRun(r) = hc {
                        runs.push(map_run(r));
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
    }
}

fn map_run(r: &wml::Run) -> Run {
    let mut props = RunProps::default();
    let mut style_id = None;
    if let Some(rpr) = r.run_properties.as_deref() {
        for c in &rpr.run_properties_choice {
            apply_run_property_choice(c, &mut props, &mut style_id);
        }
    }
    let mut text = String::new();
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
            _ => {}
        }
    }
    Run {
        style_id,
        props,
        text,
    }
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
