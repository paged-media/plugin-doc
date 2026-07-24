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

//! # docx-lower — Tier-0 lowering of a Word document to the native model
//!
//! Pure `docx-core` -> [`ir::LoweredDoc`]. No SDK, no core dependency: the
//! output is a plugin-local IR the TS host-model turns into `host.document.mutate`
//! ops (the `sheet-lower` -> `sheet-host-model` split).
//!
//! ## Why styles carry direct formatting
//!
//! The host's only range-styling op is `applyStyle(named style)` — there is no
//! "set character property on a text range" mutation. So **direct** Word
//! formatting (a bold word in a Normal paragraph) is lowered by *synthesizing* a
//! named style that carries the override (deduped by property signature, `basedOn`
//! the referenced style) and applying it. Word documents are style-heavy, so this
//! stays bounded; it is also exactly how the native model represents the result.

use std::collections::HashMap;

use docx_core::{
    Block, DocxDocument, Justification, ListKind, ListMarker, ParaProps, Run, RunProps, Section,
    Style, StyleKind, VertAlign,
};

pub mod ir;

use ir::{
    Diagnostic, LoweredBlock, LoweredCell, LoweredDoc, LoweredImage, LoweredParagraph, LoweredRun,
    LoweredSection, LoweredStory, LoweredStyle, LoweredSwatch, LoweredTabStop, LoweredTable,
    PropValue, StyleCollection, StyleProp,
};

const PARA_PREFIX: &str = "ParagraphStyle/docx-";
const CHAR_PREFIX: &str = "CharacterStyle/docx-";

/// Twips (1/1440 inch) -> points (1/72 inch). 20 twips per point.
fn twip_to_pt(twips: i32) -> f32 {
    twips as f32 / 20.0
}

/// Half-points -> points.
fn half_pt_to_pt(half: u32) -> f32 {
    half as f32 / 2.0
}

/// EMU (English Metric Units) -> points. 914400 EMU/inch, 72 pt/inch ⇒ 12700
/// EMU/pt.
fn emu_to_pt(emu: i64) -> f32 {
    emu as f32 / 12700.0
}

/// Lower an image to an anchored-frame placement, embedding the media bytes as a
/// self-contained `data:` URI (Tier-2 v1; large images should later use a part
/// reference instead of an inline base64 payload).
fn lower_image(img: &docx_core::Image) -> LoweredImage {
    use base64::Engine;
    let b64 = base64::engine::general_purpose::STANDARD.encode(&img.bytes);
    LoweredImage {
        width_pt: emu_to_pt(img.width_emu),
        height_pt: emu_to_pt(img.height_emu),
        uri: format!("data:{};base64,{}", img.mime, b64),
    }
}

/// Word `w:jc` -> the IDML justification string paged expects
/// (`Justification::as_idml`).
fn justification_idml(j: Justification) -> &'static str {
    match j {
        Justification::Left | Justification::Start => "LeftAlign",
        Justification::Center => "CenterAlign",
        Justification::Right | Justification::End => "RightAlign",
        Justification::Both => "LeftJustified",
        Justification::Distribute => "FullyJustified",
    }
}

/// Lower a whole Word document to the native IR.
pub fn lower(doc: &DocxDocument) -> LoweredDoc {
    let mut ctx = Lowering::default();

    // 0. docDefaults -> a base paragraph style every un-based style + un-styled
    //    paragraph inherits from (Word's Normal defaults: font, size, spacing).
    ctx.install_doc_defaults(&doc.styles.doc_defaults);

    // 1. The Word style catalog -> native styles (topologically ordered).
    ctx.lower_style_catalog(&doc.styles.styles);

    // 2. The body -> a native story of blocks (paragraphs + tables) in order,
    //    synthesizing styles for direct formatting.
    let mut blocks = Vec::new();
    for (idx, block) in doc.body.iter().enumerate() {
        match block {
            Block::Paragraph(p) => {
                blocks.push(LoweredBlock::Paragraph(ctx.lower_paragraph(p, idx as u32)));
            }
            Block::Table(t) => {
                blocks.push(LoweredBlock::Table(ctx.lower_table(t)));
            }
        }
    }

    if ctx.hyperlinks > 0 {
        ctx.diagnostics.push(Diagnostic::info(
            format!(
                "{} hyperlink run(s) styled (blue + underline); the clickable link is \
                 preserved in the source .docx but not yet a native clickable link",
                ctx.hyperlinks
            ),
            3,
        ));
    }

    let section = lower_section(doc.sections.first());
    let styles = ctx.ordered_styles();

    LoweredDoc {
        swatches: ctx.swatches,
        styles,
        story: LoweredStory { blocks },
        section,
        diagnostics: ctx.diagnostics,
    }
}

/// Mutable lowering state: the style table (id -> def, plus insertion order),
/// the swatch registry, the synthesis dedup cache, and diagnostics.
#[derive(Default)]
struct Lowering {
    styles: HashMap<String, LoweredStyle>,
    style_order: Vec<String>,
    swatches: Vec<LoweredSwatch>,
    swatch_index: HashMap<String, String>,
    synth_cache: HashMap<String, String>,
    synth_counter: u32,
    diagnostics: Vec<Diagnostic>,
    /// The docDefaults base style id, if docDefaults carried any properties.
    /// Un-based Word styles + un-styled paragraphs fall back to it.
    default_base: Option<String>,
    /// Count of hyperlink runs styled (for the clickable-link diagnostic).
    hyperlinks: u32,
}

impl Lowering {
    /// Create the `docx-Default` base paragraph style from docDefaults (if it
    /// carries anything), so every un-based style and un-styled paragraph
    /// inherits Word's document defaults.
    fn install_doc_defaults(&mut self, defaults: &docx_core::Defaults) {
        let mut props = self.para_props(&defaults.para);
        props.extend(self.run_props(&defaults.run));
        if props.is_empty() {
            return;
        }
        let id = format!("{PARA_PREFIX}Default");
        self.record_style(LoweredStyle {
            id: id.clone(),
            name: "docx defaults".into(),
            collection: StyleCollection::Paragraph,
            based_on: None,
            props,
        });
        self.default_base = Some(id);
    }

    /// A paragraph `basedOn` fallback: the explicit parent, else the docDefaults
    /// base style. (The default style itself is created with `basedOn: None`
    /// directly, so this never produces a self-reference.)
    fn para_base(&self, explicit: Option<String>) -> Option<String> {
        explicit.or_else(|| self.default_base.clone())
    }

    fn record_style(&mut self, style: LoweredStyle) {
        if !self.styles.contains_key(&style.id) {
            self.style_order.push(style.id.clone());
        }
        self.styles.insert(style.id.clone(), style);
    }

    /// Return styles in an order where every `basedOn` parent precedes its
    /// children (the host requires the parent to exist at create time).
    fn ordered_styles(&self) -> Vec<LoweredStyle> {
        let mut out = Vec::with_capacity(self.style_order.len());
        let mut emitted: HashMap<&str, bool> = HashMap::new();
        // Depth-bounded DFS emit (defends against a based_on cycle in malformed
        // input: MAX_DEPTH stops runaway recursion).
        fn emit<'a>(
            id: &'a str,
            styles: &'a HashMap<String, LoweredStyle>,
            emitted: &mut HashMap<&'a str, bool>,
            out: &mut Vec<LoweredStyle>,
            depth: u8,
        ) {
            if depth > 32 || emitted.get(id).copied().unwrap_or(false) {
                return;
            }
            emitted.insert(id, true);
            if let Some(style) = styles.get(id) {
                if let Some(parent) = &style.based_on {
                    if styles.contains_key(parent) {
                        emit(parent, styles, emitted, out, depth + 1);
                    }
                }
                out.push(style.clone());
            }
        }
        for id in &self.style_order {
            emit(id, &self.styles, &mut emitted, &mut out, 0);
        }
        out
    }

    /// Register a color (by `RRGGBB`) and return its swatch token.
    fn swatch_for(&mut self, hex: &str) -> String {
        if let Some(id) = self.swatch_index.get(hex) {
            return id.clone();
        }
        let id = format!("Color/docx-{hex}");
        let (r, g, b) = parse_hex(hex).unwrap_or((0.0, 0.0, 0.0));
        self.swatches.push(LoweredSwatch {
            id: id.clone(),
            name: format!("docx {hex}"),
            space: "RGB".into(),
            value: vec![r, g, b],
        });
        self.swatch_index.insert(hex.to_string(), id.clone());
        id
    }

    fn lower_style_catalog(&mut self, styles: &[Style]) {
        for s in styles {
            match s.kind {
                StyleKind::Paragraph => {
                    let mut props = self.para_props(&s.para);
                    props.extend(self.run_props(&s.run));
                    self.record_style(LoweredStyle {
                        id: format!("{PARA_PREFIX}{}", sanitize(&s.style_id)),
                        name: s.name.clone().unwrap_or_else(|| s.style_id.clone()),
                        collection: StyleCollection::Paragraph,
                        based_on: self.para_base(
                            s.based_on
                                .as_ref()
                                .map(|b| format!("{PARA_PREFIX}{}", sanitize(b))),
                        ),
                        props,
                    });
                }
                StyleKind::Character => {
                    let props = self.run_props(&s.run);
                    self.record_style(LoweredStyle {
                        id: format!("{CHAR_PREFIX}{}", sanitize(&s.style_id)),
                        name: s.name.clone().unwrap_or_else(|| s.style_id.clone()),
                        collection: StyleCollection::Character,
                        based_on: s
                            .based_on
                            .as_ref()
                            .map(|b| format!("{CHAR_PREFIX}{}", sanitize(b))),
                        props,
                    });
                }
                StyleKind::Table | StyleKind::Numbering => {
                    self.diagnostics.push(Diagnostic::info(
                        format!(
                            "style '{}' ({:?}) is not a Tier-0 construct and was skipped",
                            s.style_id, s.kind
                        ),
                        3,
                    ));
                }
            }
        }
    }

    fn lower_paragraph(&mut self, p: &docx_core::Paragraph, source_index: u32) -> LoweredParagraph {
        let explicit = p
            .style_id
            .as_ref()
            .map(|id| format!("{PARA_PREFIX}{}", sanitize(id)));
        let mut props = self.para_props(&p.props);
        if let Some(list) = &p.list {
            props.extend(list_props(list));
        }
        let para_style_id = if props.is_empty() {
            // No direct formatting or list: apply the paragraph's style, or the
            // docDefaults base when the paragraph carries no style at all.
            self.para_base(explicit)
        } else {
            let base = self.para_base(explicit);
            Some(self.synthesize(StyleCollection::Paragraph, base, props))
        };

        let runs = p
            .runs
            .iter()
            .filter(|r| !r.text.is_empty())
            .map(|r| self.lower_run(r))
            .collect();

        // Inline images ride on their own (empty-text) runs; collect them as
        // anchored-frame placements for this paragraph.
        let images = p
            .runs
            .iter()
            .filter_map(|r| r.image.as_ref().map(lower_image))
            .collect();

        LoweredParagraph {
            para_style_id,
            runs,
            images,
            source_index,
        }
    }

    fn lower_run(&mut self, r: &Run) -> LoweredRun {
        let base = r
            .style_id
            .as_ref()
            .map(|id| format!("{CHAR_PREFIX}{}", sanitize(id)));
        let mut props = self.run_props(&r.props);

        // A hyperlink run gets the conventional look — blue + underline — unless
        // the run already sets those directly. (The clickable link itself is
        // preserved in the source .docx; a native clickable-hyperlink door is
        // future work.)
        if r.hyperlink.is_some() {
            self.hyperlinks += 1;
            if !props.iter().any(|p| p.path == "characterFillColor") {
                let swatch = self.swatch_for("0000FF");
                props.push(StyleProp {
                    path: "characterFillColor".into(),
                    value: PropValue::ColorRef(swatch),
                });
            }
            if !props.iter().any(|p| p.path == "characterUnderline") {
                props.push(boolean("characterUnderline", true));
            }
        }

        let char_style_id = if props.is_empty() {
            base
        } else {
            Some(self.synthesize(StyleCollection::Character, base, props))
        };
        LoweredRun {
            text: r.text.clone(),
            char_style_id,
        }
    }

    /// Lower a table: resolve the grid (gridSpan widens a cell across columns,
    /// vMerge merges cells down rows) into positioned cells with spans. A
    /// vMerge-continue cell is absorbed into its restart cell above (not emitted).
    fn lower_table(&mut self, t: &docx_core::Table) -> LoweredTable {
        let cols = if !t.column_widths.is_empty() {
            t.column_widths.len() as u32
        } else {
            t.rows
                .iter()
                .map(|r| r.cells.iter().map(|c| c.grid_span.max(1)).sum::<u32>())
                .max()
                .unwrap_or(1)
        };
        let column_widths_pt = t.column_widths.iter().map(|w| twip_to_pt(*w)).collect();

        let mut cells: Vec<LoweredCell> = Vec::new();
        // column index -> index into `cells` of the active vMerge restart cell.
        let mut vmerge_anchor: HashMap<u32, usize> = HashMap::new();
        for (r, row) in t.rows.iter().enumerate() {
            let mut col = 0u32;
            for cell in &row.cells {
                let span = cell.grid_span.max(1);
                match cell.v_merge {
                    docx_core::VMerge::Continue => {
                        if let Some(&idx) = vmerge_anchor.get(&col) {
                            cells[idx].row_span += 1;
                        }
                    }
                    _ => {
                        let paragraphs = cell
                            .paragraphs
                            .iter()
                            .map(|p| self.lower_paragraph(p, 0))
                            .collect();
                        let idx = cells.len();
                        cells.push(LoweredCell {
                            row: r as u32,
                            col,
                            row_span: 1,
                            col_span: span,
                            paragraphs,
                        });
                        if cell.v_merge == docx_core::VMerge::Restart {
                            vmerge_anchor.insert(col, idx);
                        } else {
                            vmerge_anchor.remove(&col);
                        }
                    }
                }
                col += span;
            }
        }

        LoweredTable {
            rows: t.rows.len() as u32,
            cols,
            column_widths_pt,
            cells,
        }
    }

    /// Synthesize (or reuse) a named style carrying direct formatting.
    fn synthesize(
        &mut self,
        collection: StyleCollection,
        based_on: Option<String>,
        props: Vec<StyleProp>,
    ) -> String {
        let sig = synth_signature(collection, &based_on, &props);
        if let Some(id) = self.synth_cache.get(&sig) {
            return id.clone();
        }
        self.synth_counter += 1;
        let n = self.synth_counter;
        let id = match collection {
            StyleCollection::Paragraph => format!("{PARA_PREFIX}auto-p{n}"),
            StyleCollection::Character => format!("{CHAR_PREFIX}auto-c{n}"),
        };
        self.record_style(LoweredStyle {
            id: id.clone(),
            name: format!("docx direct format {n}"),
            collection,
            based_on,
            props,
        });
        self.synth_cache.insert(sig, id.clone());
        id
    }

    fn para_props(&mut self, p: &ParaProps) -> Vec<StyleProp> {
        let mut out = Vec::new();
        if let Some(j) = p.justification {
            out.push(StyleProp {
                path: "paragraphJustification".into(),
                value: PropValue::Text(justification_idml(j).into()),
            });
        }
        if let Some(v) = p.left_indent {
            out.push(len("paragraphLeftIndent", twip_to_pt(v)));
        }
        if let Some(v) = p.right_indent {
            out.push(len("paragraphRightIndent", twip_to_pt(v)));
        }
        // A hanging indent is a negative first-line indent; it wins over an
        // explicit firstLine when both are (unusually) present.
        if let Some(v) = p.hanging_indent {
            out.push(len("paragraphFirstLineIndent", -twip_to_pt(v)));
        } else if let Some(v) = p.first_line_indent {
            out.push(len("paragraphFirstLineIndent", twip_to_pt(v)));
        }
        if let Some(v) = p.space_before {
            out.push(len("paragraphSpaceBefore", twip_to_pt(v)));
        }
        if let Some(v) = p.space_after {
            out.push(len("paragraphSpaceAfter", twip_to_pt(v)));
        }
        // Word's keepNext is a boolean; paged's keepWithNext is a line count, so
        // "on" maps to a single-line hold.
        if p.keep_next == Some(true) {
            out.push(len("paragraphKeepWithNext", 1.0));
        }
        if let Some(k) = p.keep_lines {
            out.push(boolean("paragraphKeepLinesTogether", k));
        }
        if !p.tabs.is_empty() {
            let stops = p
                .tabs
                .iter()
                .map(|t| LoweredTabStop {
                    position: twip_to_pt(t.position),
                    alignment: t.alignment.clone(),
                    alignment_character: None,
                    leader: t.leader.clone(),
                })
                .collect();
            out.push(StyleProp {
                path: "paragraphTabStops".into(),
                value: PropValue::TabStops(stops),
            });
        }
        out
    }

    fn run_props(&mut self, r: &RunProps) -> Vec<StyleProp> {
        let mut out = Vec::new();
        if let Some(font) = &r.font {
            out.push(StyleProp {
                path: "characterFontFamily".into(),
                value: PropValue::Text(font.clone()),
            });
        }
        // Word's separate bold/italic toggles collapse to one paged font style.
        if r.bold.is_some() || r.italic.is_some() {
            let style = match (r.bold.unwrap_or(false), r.italic.unwrap_or(false)) {
                (true, true) => "Bold Italic",
                (true, false) => "Bold",
                (false, true) => "Italic",
                (false, false) => "Regular",
            };
            out.push(StyleProp {
                path: "characterFontStyle".into(),
                value: PropValue::Text(style.into()),
            });
        }
        if let Some(half) = r.size_half_pts {
            out.push(len("characterFontSize", half_pt_to_pt(half)));
        }
        if let Some(hex) = &r.color {
            let swatch = self.swatch_for(hex);
            out.push(StyleProp {
                path: "characterFillColor".into(),
                value: PropValue::ColorRef(swatch),
            });
        }
        if let Some(u) = r.underline {
            out.push(boolean("characterUnderline", u));
        }
        if let Some(s) = r.strike {
            out.push(boolean("characterStrikethru", s));
        }
        match r.vert_align {
            Some(VertAlign::Superscript) => out.push(StyleProp {
                path: "characterPosition".into(),
                value: PropValue::Text("Superscript".into()),
            }),
            Some(VertAlign::Subscript) => out.push(StyleProp {
                path: "characterPosition".into(),
                value: PropValue::Text("Subscript".into()),
            }),
            _ => {}
        }
        // Capitalization: small caps wins over all caps when both are set.
        if r.small_caps == Some(true) {
            out.push(StyleProp {
                path: "characterCase".into(),
                value: PropValue::Text("SmallCaps".into()),
            });
        } else if r.caps == Some(true) {
            out.push(StyleProp {
                path: "characterCase".into(),
                value: PropValue::Text("AllCaps".into()),
            });
        }
        // Baseline shift: Word half-points (signed) -> points.
        if let Some(half) = r.baseline_half_pts {
            out.push(len("characterBaselineShift", half as f32 / 2.0));
        }
        out
    }
}

fn len(path: &str, pt: f32) -> StyleProp {
    StyleProp {
        path: path.into(),
        value: PropValue::Length(pt),
    }
}

fn boolean(path: &str, v: bool) -> StyleProp {
    StyleProp {
        path: path.into(),
        value: PropValue::Bool(v),
    }
}

/// The style props that turn a paragraph into a native list item: the list type
/// (which the engine's renderer gates marker emission + auto-numbering on), the
/// bullet glyph or numbering format, and a per-level left indent. Emitted after
/// the direct paragraph props so the list indent wins over an inherited one.
fn list_props(list: &ListMarker) -> Vec<StyleProp> {
    let mut out = Vec::new();
    let list_type = match list.kind {
        ListKind::Bullet => "BulletList",
        ListKind::Numbered => "NumberedList",
    };
    out.push(StyleProp {
        path: "paragraphListType".into(),
        value: PropValue::Text(list_type.into()),
    });
    if let Some(ch) = &list.bullet_char {
        out.push(StyleProp {
            path: "paragraphBulletCharacter".into(),
            value: PropValue::Text(ch.clone()),
        });
    }
    if let Some(fmt) = &list.number_format {
        out.push(StyleProp {
            path: "paragraphNumberingFormat".into(),
            value: PropValue::Text(fmt.clone()),
        });
    }
    // Each level indents by 18 pt (¼ inch) — a reasonable default when the list
    // definition's own indent metrics are not (yet) carried.
    out.push(len("paragraphLeftIndent", (list.level as f32 + 1.0) * 18.0));
    out
}

fn lower_section(section: Option<&Section>) -> LoweredSection {
    let s = section.cloned().unwrap_or_default();
    LoweredSection {
        page_width_pt: twip_to_pt(s.page_width),
        page_height_pt: twip_to_pt(s.page_height),
        margin_top_pt: twip_to_pt(s.margin_top),
        margin_bottom_pt: twip_to_pt(s.margin_bottom),
        margin_left_pt: twip_to_pt(s.margin_left),
        margin_right_pt: twip_to_pt(s.margin_right),
        columns: s.columns.max(1),
    }
}

/// A stable dedup key for a synthesized style.
fn synth_signature(
    collection: StyleCollection,
    based_on: &Option<String>,
    props: &[StyleProp],
) -> String {
    let mut sig = format!("{collection:?}|{}|", based_on.as_deref().unwrap_or(""));
    for p in props {
        sig.push_str(&p.path);
        sig.push('=');
        match &p.value {
            PropValue::Text(t) => sig.push_str(t),
            PropValue::Length(l) => sig.push_str(&format!("{l}")),
            PropValue::Bool(b) => sig.push_str(if *b { "1" } else { "0" }),
            PropValue::ColorRef(c) => sig.push_str(c),
            PropValue::TabStops(stops) => {
                for s in stops {
                    sig.push_str(&format!(
                        "{}:{}:{},",
                        s.position,
                        s.alignment.as_deref().unwrap_or(""),
                        s.leader.as_deref().unwrap_or("")
                    ));
                }
            }
        }
        sig.push(';');
    }
    sig
}

/// Sanitize a Word style id into a token-safe suffix (`/` and whitespace would
/// break the `Collection/id` token grammar).
fn sanitize(id: &str) -> String {
    id.chars()
        .map(|c| if c.is_ascii_alphanumeric() { c } else { '-' })
        .collect()
}

/// Parse `RRGGBB` into `[r, g, b]` on 0–255.
fn parse_hex(hex: &str) -> Option<(f32, f32, f32)> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some((r as f32, g as f32, b as f32))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_core::*;

    fn run(text: &str, props: RunProps) -> Run {
        Run {
            style_id: None,
            props,
            image: None,
            hyperlink: None,
            text: text.into(),
        }
    }

    #[test]
    fn lowers_styles_paragraphs_and_synthesizes_direct_bold() {
        let mut doc = DocxDocument::default();
        doc.styles.styles.push(Style {
            style_id: "Heading1".into(),
            name: Some("heading 1".into()),
            kind: StyleKind::Paragraph,
            based_on: Some("Normal".into()),
            para: ParaProps {
                justification: Some(Justification::Center),
                ..Default::default()
            },
            run: RunProps {
                size_half_pts: Some(48),
                ..Default::default()
            },
        });
        doc.styles.styles.insert(
            0,
            Style {
                style_id: "Normal".into(),
                name: Some("Normal".into()),
                kind: StyleKind::Paragraph,
                ..Default::default()
            },
        );
        doc.body.push(Block::Paragraph(Paragraph {
            style_id: Some("Heading1".into()),
            props: ParaProps::default(),
            runs: vec![
                run("Hello ", RunProps::default()),
                run(
                    "bold",
                    RunProps {
                        bold: Some(true),
                        color: Some("FF0000".into()),
                        ..Default::default()
                    },
                ),
            ],
            list: None,
        }));

        let lowered = lower(&doc);

        // Heading1 references Normal, which must be created first.
        let ids: Vec<&str> = lowered.styles.iter().map(|s| s.id.as_str()).collect();
        let normal = ids.iter().position(|s| s.ends_with("docx-Normal")).unwrap();
        let heading = ids
            .iter()
            .position(|s| s.ends_with("docx-Heading1"))
            .unwrap();
        assert!(normal < heading, "based_on parent must precede child");

        // The bold+red run synthesized a character style and a swatch.
        assert_eq!(lowered.swatches.len(), 1);
        assert_eq!(lowered.swatches[0].value, vec![255.0, 0.0, 0.0]);
        let para = &lowered.story.paragraphs()[0];
        assert!(para
            .para_style_id
            .as_deref()
            .unwrap()
            .ends_with("docx-Heading1"));
        assert_eq!(para.runs.len(), 2);
        assert!(para.runs[0].char_style_id.is_none());
        let synth = para.runs[1].char_style_id.as_deref().unwrap();
        assert!(synth.starts_with(CHAR_PREFIX));
        let synth_style = lowered.styles.iter().find(|s| s.id == synth).unwrap();
        assert!(synth_style
            .props
            .iter()
            .any(|p| p.path == "characterFontStyle" && p.value == PropValue::Text("Bold".into())));
    }

    #[test]
    fn caps_and_baseline_shift_lower_to_character_props() {
        let mut doc = DocxDocument::default();
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![run(
                "x",
                RunProps {
                    small_caps: Some(true),
                    baseline_half_pts: Some(6), // 6 half-pt -> 3 pt
                    ..Default::default()
                },
            )],
            ..Default::default()
        }));
        let ir = lower(&doc);
        let sid = ir.story.paragraphs()[0].runs[0]
            .char_style_id
            .clone()
            .unwrap();
        let style = ir.styles.iter().find(|s| s.id == sid).unwrap();
        assert!(style
            .props
            .iter()
            .any(|p| p.path == "characterCase" && p.value == PropValue::Text("SmallCaps".into())));
        assert!(style
            .props
            .iter()
            .any(|p| p.path == "characterBaselineShift" && p.value == PropValue::Length(3.0)));
    }

    #[test]
    fn dedups_identical_direct_formatting() {
        let mut doc = DocxDocument::default();
        let bold = RunProps {
            bold: Some(true),
            ..Default::default()
        };
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![run("a", bold.clone()), run("b", bold.clone())],
            ..Default::default()
        }));
        let lowered = lower(&doc);
        let s0 = lowered.story.paragraphs()[0].runs[0].char_style_id.clone();
        let s1 = lowered.story.paragraphs()[0].runs[1].char_style_id.clone();
        assert_eq!(s0, s1, "identical direct formatting reuses one synth style");
        let synth_count = lowered
            .styles
            .iter()
            .filter(|s| s.id.contains("auto-c"))
            .count();
        assert_eq!(synth_count, 1);
    }

    #[test]
    fn section_twips_convert_to_points() {
        let mut doc = DocxDocument::default();
        doc.sections.push(Section::default());
        let l = lower(&doc);
        assert_eq!(l.section.page_width_pt, 612.0);
        assert_eq!(l.section.page_height_pt, 792.0);
        assert_eq!(l.section.margin_left_pt, 72.0);
    }
}
