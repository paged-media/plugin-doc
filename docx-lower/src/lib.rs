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

use std::collections::{BTreeMap, BTreeSet};

use docx_core::{
    Block, BreakKind, DocxDocument, Justification, LineRule, LineSpacing, ListKind, ListMarker,
    ParaProps, Run, RunProps, Section, SectionKind, Style, StyleKind, VertAlign,
};

pub mod ir;
pub mod line_height;
pub mod sections;

pub use sections::LowerOptions;
use sections::{SectionColumns, SectionPlacement};

use ir::{
    Diagnostic, LoweredBlock, LoweredCell, LoweredDoc, LoweredImage, LoweredParagraph, LoweredRun,
    LoweredSection, LoweredSegment, LoweredStory, LoweredStyle, LoweredSwatch, LoweredTabStop,
    LoweredTable, PropValue, StyleCollection, StyleProp,
};

const PARA_PREFIX: &str = "ParagraphStyle/docx-";

/// The break-before rule's property path (thoughts ADR 028; settable from core
/// protocol 64 — older engines refuse it, and `applyStyleOps` then loses only
/// the break, with a warning). Its value is InDesign's `StartParagraph` string.
const START_PARAGRAPH: &str = "paragraphStartParagraph";

/// IDML's span/split columns (`SpanColumnType` and companions; settable from
/// core protocol 64, ADR 028 addendum). A mid-page column change lowers to
/// them (`sections`); an older engine refuses them, and the bundle reopens
/// the document with [`LowerOptions::mid_page_columns`] off.
pub const SPAN_COLUMN_TYPE: &str = "paragraphSpanColumnType";
pub const SPAN_SPLIT_COLUMN_COUNT: &str = "paragraphSpanSplitColumnCount";
pub const SPLIT_COLUMN_INSIDE_GUTTER: &str = "paragraphSplitColumnInsideGutter";
pub const SPLIT_COLUMN_OUTSIDE_GUTTER: &str = "paragraphSplitColumnOutsideGutter";

/// Word's single line spacing (`w:spacing w:line="240" w:lineRule="auto"`),
/// what Word lays when no line spacing is set anywhere (ADR 029; measured as
/// case L1 of `fixtures/line-spacing.word.json`).
const SINGLE: LineSpacing = LineSpacing {
    value: 240,
    rule: LineRule::Auto,
};

/// The `StartParagraph` value a page or column break lowers to. A column
/// break in a one-column section opens the next frame, which is the next
/// page, as Word does.
fn break_rule(kind: BreakKind) -> &'static str {
    match kind {
        BreakKind::Page => "NextPage",
        BreakKind::Column => "NextColumn",
    }
}

/// How much a `StartParagraph` rule asks for: odd/even pages imply a new page,
/// which implies a new column. Two rules on one paragraph keep the stronger.
fn rule_rank(rule: &str) -> u8 {
    match rule {
        "NextOddPage" | "NextEvenPage" => 3,
        "NextPage" => 2,
        "NextColumn" => 1,
        _ => 0,
    }
}

fn stronger(a: Option<&'static str>, b: Option<&'static str>) -> Option<&'static str> {
    match (a, b) {
        (Some(x), Some(y)) => Some(if rule_rank(y) > rule_rank(x) { y } else { x }),
        (x, None) => x,
        (None, y) => y,
    }
}
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

/// Blank lines that cannot keep different styles. A Word blank line (an
/// empty paragraph) pours as an empty native paragraph, styled by a caret
/// `applyStyle` at its offset in the contiguous character space (core
/// `65cf615`). It has no characters, so consecutive blank lines share one
/// offset, as do blank lines on either side of a table (whose host paragraph
/// has none either), and a caret styles all of them at once. The pour applies
/// the carets last, so they are never refused, and the LAST one wins: where
/// such a group has different styles in Word, they all take the last one's,
/// and their line heights may differ from Word's. Reported, not hidden.
fn blank_line_styles(
    blocks: &[LoweredBlock],
    starts_story: &dyn Fn(usize) -> bool,
    diagnostics: &mut Vec<Diagnostic>,
) {
    let blank = |b: &LoweredBlock| matches!(b, LoweredBlock::Paragraph(p) if p.runs.is_empty() && p.segments.is_empty());
    let mut group: Vec<(usize, Option<&str>)> = Vec::new();
    let mut flush = |group: &mut Vec<(usize, Option<&str>)>| {
        let last_styled = group.iter().rev().find_map(|(_, s)| *s);
        if group.iter().any(|(_, s)| *s != group[0].1) {
            let list: Vec<String> = group
                .iter()
                .map(|(i, s)| format!("{i} ({})", s.unwrap_or("no style")))
                .collect();
            diagnostics.push(Diagnostic::warning(
                format!(
                    "blank lines at body blocks {} have different paragraph styles; \
                     the engine styles blank lines at one position together, so they \
                     all take {} and may not have Word's line heights",
                    list.join(", "),
                    last_styled.unwrap_or("the default style"),
                ),
                1,
            ));
        }
        group.clear();
    };
    for (idx, block) in blocks.iter().enumerate() {
        if starts_story(idx) {
            flush(&mut group);
        }
        match block {
            LoweredBlock::Paragraph(p) if blank(block) => {
                group.push((idx, p.para_style_id.as_deref()));
            }
            // A table has no characters: blank lines on both sides meet.
            LoweredBlock::Table(_) => {}
            LoweredBlock::Paragraph(_) => flush(&mut group),
        }
    }
    flush(&mut group);

    // Inside a table cell the same holds in the cell's own offset space:
    // consecutive blank lines of one cell share a caret.
    for (idx, block) in blocks.iter().enumerate() {
        let LoweredBlock::Table(t) = block else {
            continue;
        };
        for cell in &t.cells {
            let mut group: Vec<(usize, Option<&str>)> = Vec::new();
            let mut flush = |group: &mut Vec<(usize, Option<&str>)>| {
                if group.iter().any(|(_, s)| *s != group[0].1) {
                    let last_styled = group.iter().rev().find_map(|(_, s)| *s);
                    let list: Vec<String> = group
                        .iter()
                        .map(|(i, s)| format!("{i} ({})", s.unwrap_or("no style")))
                        .collect();
                    diagnostics.push(Diagnostic::warning(
                        format!(
                            "blank lines in the table at body block {idx}, cell row {} column \
                             {}, paragraphs {} have different paragraph styles; the engine \
                             styles blank lines at one position together, so they all take \
                             {} and may not have Word's line heights",
                            cell.row,
                            cell.col,
                            list.join(", "),
                            last_styled.unwrap_or("the default style"),
                        ),
                        1,
                    ));
                }
                group.clear();
            };
            for (k, p) in cell.paragraphs.iter().enumerate() {
                if p.runs.is_empty() {
                    group.push((k, p.para_style_id.as_deref()));
                } else {
                    flush(&mut group);
                }
            }
            flush(&mut group);
        }
    }
}

/// Lower a whole Word document to the native IR, for an engine with every
/// property the lowering uses ([`LowerOptions::default`]).
pub fn lower(doc: &DocxDocument) -> LoweredDoc {
    lower_with(doc, LowerOptions::default())
}

/// Lower a whole Word document to the native IR with `options` (what the
/// target engine can lay out).
pub fn lower_with(doc: &DocxDocument, options: LowerOptions) -> LoweredDoc {
    let mut ctx = Lowering {
        word_styles: doc
            .styles
            .styles
            .iter()
            .map(|s| (s.style_id.clone(), s.clone()))
            .collect(),
        word_defaults: doc.styles.doc_defaults.clone(),
        ..Lowering::default()
    };

    // 0. docDefaults -> a base paragraph style every un-based style + un-styled
    //    paragraph inherits from (Word's Normal defaults: font, size, spacing).
    ctx.install_doc_defaults(&doc.styles.doc_defaults);

    // 1. The Word style catalog -> native styles (topologically ordered).
    ctx.lower_style_catalog(&doc.styles.styles);

    // 2. The body -> a native story of blocks (paragraphs + tables) in order,
    //    synthesizing styles for direct formatting.
    //
    //    ADR 028/029 — Word's breaks become the engine's break-before rule
    //    (`paragraphStartParagraph`) on the paragraph that must start over:
    //    an `oddPage`/`evenPage` section's first paragraph, a paragraph after
    //    a page/column break, or the part of a paragraph after a break inside
    //    it. (`nextPage` sections need no rule: each section is its own story
    //    on its own page, ADR 029 addendum.)
    //
    //    `continuous` / `nextColumn` sections JOIN the story before them when
    //    the native model can say what Word does (`sections` has the rules and
    //    Word's measurements): no break, a `NextColumn` rule, or the margin
    //    difference as paragraph indents.
    let placements = sections::place_sections_with(&doc.sections, options, &mut ctx.diagnostics);
    let section_starts: Vec<(usize, SectionKind, Option<&SectionPlacement>)> = doc
        .sections
        .iter()
        .enumerate()
        .map(|(k, s)| {
            // A section joining the story before it, or `None` for a story's
            // first section.
            let joined = placements
                .get(k)
                .filter(|p| k > 0 && placements.get(k - 1).map(|q| q.story) == Some(p.story));
            (s.first_block, s.kind, joined)
        })
        .collect();
    let starts_section = |idx: usize| section_starts.iter().find(|(first, _, _)| *first == idx);
    // A new STORY starts at `idx` (a section that does not join).
    let starts_story = |idx: usize| starts_section(idx).is_some_and(|(_, _, j)| j.is_none());
    let mut blocks = Vec::new();
    let mut carry: Option<BreakKind> = None;
    // The text width of the story being lowered (its first section's page
    // width less its margins).
    let mut story_width = 0.0f32;
    for (idx, block) in doc.body.iter().enumerate() {
        ctx.current_block = idx;
        if idx == 0 || starts_story(idx) {
            let default = docx_core::Section::default();
            let s = doc
                .sections
                .iter()
                .rev()
                .find(|s| s.first_block <= idx)
                .unwrap_or(&default);
            story_width = twip_to_pt(s.page_width - s.margin_left - s.margin_right);
        }
        let section_start = starts_section(idx);
        if let Some((_, _, joined)) = section_start {
            let (l, r) = joined.map_or((0.0, 0.0), |p| (p.indent_left_pt, p.indent_right_pt));
            ctx.section_indent = (l, r);
        }
        if idx == 0 || section_start.is_some() {
            // The section in force (the last one starting at or before
            // `idx`) and where it sits in its story's columns.
            let k = doc.sections.iter().rposition(|s| s.first_block <= idx);
            ctx.section_columns = k
                .and_then(|k| placements.get(k))
                .map_or(SectionColumns::Frame, |p| p.columns);
            let one_column = k.is_none_or(|k| doc.sections[k].columns <= 1);
            // A one-column section spans the story's whole text width.
            ctx.text_width = one_column.then_some(story_width);
        }
        if starts_story(idx) && idx > 0 {
            if let Some(kind) = carry.take() {
                ctx.diagnostics.push(Diagnostic::info(
                    format!(
                        "a {} break ends the section before body block {idx}; the next \
                         section starts a new page anyway, so the break adds nothing \
                         here (Word may add a blank page for it; not measured)",
                        if kind == BreakKind::Page {
                            "page"
                        } else {
                            "column"
                        }
                    ),
                    3,
                ));
            }
        }
        // The FIRST section's kind is not applied: it has nothing before it
        // to start after, and Word was not asked about it.
        let entry = match section_start.filter(|_| idx > 0) {
            Some((_, SectionKind::OddPage, _)) => Some("NextOddPage"),
            Some((_, SectionKind::EvenPage, _)) => Some("NextEvenPage"),
            Some((_, _, Some(p))) if p.next_column => Some("NextColumn"),
            _ => None,
        };
        match block {
            Block::Paragraph(p) => {
                // A trailing break carries to the next block when that is a
                // paragraph of the same STORY (a joined section's first
                // paragraph included).
                let next_in_section = matches!(doc.body.get(idx + 1), Some(Block::Paragraph(_)))
                    && !starts_story(idx + 1);
                let start = stronger(entry, carry.take().map(break_rule));
                let (para, tail) =
                    ctx.lower_body_paragraph(p, idx as u32, start, Some(next_in_section));
                carry = tail;
                blocks.push(LoweredBlock::Paragraph(para));
            }
            Block::Table(t) => {
                if let Some(rule) = entry {
                    ctx.diagnostics.push(Diagnostic::info(
                        format!(
                            "the section at body block {idx} starts with a table, which \
                             cannot carry its {rule} start; it starts on the section's \
                             own page whatever that page's parity"
                        ),
                        3,
                    ));
                }
                if ctx.section_columns != SectionColumns::Frame {
                    ctx.diagnostics.push(Diagnostic::info(
                        format!(
                            "the table at body block {idx} is in a section that changes the \
                             columns mid-page; a table cannot span or split columns, so it \
                             sits in the story's own column"
                        ),
                        3,
                    ));
                }
                if ctx.section_indent != (0.0, 0.0) {
                    ctx.diagnostics.push(Diagnostic::info(
                        format!(
                            "the table at body block {idx} is in a continuous section with \
                             other left/right margins; its paragraphs keep the margins of \
                             the page it continues"
                        ),
                        3,
                    ));
                }
                blocks.push(LoweredBlock::Table(ctx.lower_table(t)));
            }
        }
    }
    blank_line_styles(&blocks, &starts_story, &mut ctx.diagnostics);
    if ctx.breaks > 0 {
        ctx.diagnostics.push(Diagnostic::info(
            format!(
                "{} page/column break(s) lowered to the engine's break-before rule \
                 (paragraphStartParagraph); an engine before protocol 64 refuses the \
                 rule and only the breaks are lost",
                ctx.breaks
            ),
            3,
        ));
    }

    if ctx.hyperlinks > 0 {
        let styled_only = ctx.hyperlinks - ctx.clickable_links;
        let mut msg = format!(
            "{} hyperlink run(s) styled (blue + underline); {} became native \
             clickable link(s)",
            ctx.hyperlinks, ctx.clickable_links
        );
        if styled_only > 0 {
            msg.push_str(&format!(
                ", {styled_only} internal #anchor target(s) stay styled-only \
                 (native URL destinations only for now)"
            ));
        }
        ctx.diagnostics.push(Diagnostic::info(msg, 3));
    }

    // Footnotes / endnotes: the references and bodies are PARSED and preserved in
    // the source package, but the native model has no note construct (there is no
    // "insert footnote" mutation and no page-bottom note area), so nothing is
    // placed. Say so rather than dropping them silently — and never fake a
    // rendering by inlining note text into the flow.
    if !doc.notes.is_empty() {
        let footnotes = doc.notes.iter().filter(|n| !n.endnote).count();
        let endnotes = doc.notes.len() - footnotes;
        let refs: usize = doc
            .body
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => Some(p),
                _ => None,
            })
            .flat_map(|p| p.runs.iter())
            .filter(|r| r.note_ref.is_some())
            .count();
        let mut msg = String::new();
        if footnotes > 0 {
            msg.push_str(&format!("{footnotes} footnote(s)"));
        }
        if endnotes > 0 {
            if !msg.is_empty() {
                msg.push_str(" + ");
            }
            msg.push_str(&format!("{endnotes} endnote(s)"));
        }
        msg.push_str(&format!(
            " parsed ({refs} in-text reference(s)); the note text is preserved in \
             the source .docx and round-trips on save, but is NOT placed on the \
             page — the native model has no footnote construct yet"
        ));
        ctx.diagnostics.push(Diagnostic::info(msg, 3));
    }

    // Non-HYPERLINK FIELDS. Word computes a field's value and stores it as the
    // field's RESULT text; we keep that text (so the document still reads
    // correctly) but it is a FROZEN snapshot — the native model has no equivalent
    // for most field kinds, and nothing recomputes it. Name the kinds rather than
    // letting them look like ordinary text.
    {
        let mut kinds: Vec<&str> = doc
            .body
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => Some(p),
                _ => None,
            })
            .flat_map(|p| p.runs.iter())
            .filter_map(|r| r.field.as_deref())
            .filter(|f| *f != "HYPERLINK")
            .collect();
        kinds.sort_unstable();
        let total = kinds.len();
        kinds.dedup();
        if total > 0 {
            ctx.diagnostics.push(Diagnostic::info(
                format!(
                    "{total} field(s) ({}) lowered as their last-computed TEXT — a \
                     frozen snapshot. The value is preserved and round-trips on \
                     save, but nothing recomputes it (the native model has no \
                     equivalent for these field kinds)",
                    kinds.join(", ")
                ),
                3,
            ));
        }
    }

    // HEADERS / FOOTERS. Parsed and preserved in the source package, but the
    // native model places body content only — there is no header/footer story to
    // pour them into — so they are not rendered. Say so.
    if !doc.headers_footers.is_empty() {
        let headers = doc.headers_footers.iter().filter(|h| !h.footer).count();
        let footers = doc.headers_footers.len() - headers;
        let mut msg = String::new();
        if headers > 0 {
            msg.push_str(&format!("{headers} header(s)"));
        }
        if footers > 0 {
            if !msg.is_empty() {
                msg.push_str(" + ");
            }
            msg.push_str(&format!("{footers} footer(s)"));
        }
        msg.push_str(
            " parsed; their text is preserved in the source .docx and round-trips \
             on save, but is NOT placed on the page — the native model has no \
             header/footer story yet",
        );
        ctx.diagnostics.push(Diagnostic::info(msg, 3));
    }

    // Symbol characters (`<w:sym>`): carried as their Unicode equivalent,
    // drawn in the run's font, not the symbol font; or, with no
    // equivalent, not carried at all (the source keeps them, and an edit
    // to their run is refused).
    if !ctx.symbols_carried.is_empty() {
        let fonts: Vec<String> = ctx
            .symbols_carried
            .iter()
            .map(|(f, n)| format!("{n} from {f}"))
            .collect();
        ctx.diagnostics.push(Diagnostic::info(
            format!(
                "symbol character(s) ({}) carried as their Unicode equivalents, drawn in \
                 the run's font (with fallback) rather than in the symbol font",
                fonts.join(", ")
            ),
            3,
        ));
    }
    for (block, font, code) in &ctx.symbols_dropped {
        ctx.diagnostics.push(Diagnostic::warning(
            format!(
                "a symbol character in body block {block} ({font} {code}) has no Unicode \
                 equivalent and is not shown; the source keeps it, and an edit to its run \
                 is refused on save"
            ),
            1,
        ));
    }

    if !ctx.unmeasured_faces.is_empty() {
        let faces: Vec<&str> = ctx.unmeasured_faces.iter().map(String::as_str).collect();
        ctx.diagnostics.push(Diagnostic::info(
            format!(
                "line spacing in face(s) Word has not been measured on ({}): their \
                 single line is taken as Calibri's/Aptos's 1.2207 em, so line \
                 pitch may differ slightly from Word's",
                faces.join(", ")
            ),
            3,
        ));
    }

    let section = lower_section(doc.sections.first());
    let sections = if doc.sections.is_empty() {
        vec![section.clone()]
    } else {
        doc.sections
            .iter()
            .zip(&placements)
            .map(|(s, p)| LoweredSection {
                story: p.story,
                ..lower_section(Some(s))
            })
            .collect()
    };
    let styles = ctx.ordered_styles();

    LoweredDoc {
        swatches: ctx.swatches,
        styles,
        story: LoweredStory { blocks },
        section,
        sections,
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
    /// Count of hyperlink runs styled (blue + underline).
    hyperlinks: u32,
    /// Of those, how many carry an EXTERNAL target that becomes a native
    /// clickable link (internal `#anchor` targets stay styled-only for now).
    clickable_links: u32,
    /// The Word style catalog by id, to resolve inherited line spacing and
    /// run font/size through `basedOn` chains.
    word_styles: HashMap<String, Style>,
    /// Word's `docDefaults`, the root of every inheritance chain.
    word_defaults: docx_core::Defaults,
    /// The `characterLeading` each lowered paragraph style resolves to (its
    /// own or inherited), so a paragraph only overrides when it differs.
    style_leading: HashMap<String, f32>,
    /// Faces a line-spacing computation needed but Word was not measured on.
    unmeasured_faces: BTreeSet<String>,
    /// Page/column breaks (`w:br`) seen in body paragraphs.
    breaks: u32,
    /// The extra (left, right) indent in pt of the section being lowered:
    /// a continuous section with other margins, joined to the story before
    /// it (`sections`). Body paragraphs only.
    section_indent: (f32, f32),
    /// How the section being lowered sits in its story's columns (span or
    /// split columns for a mid-page column change, `sections`). Body
    /// paragraphs only.
    section_columns: SectionColumns,
    /// The text width in pt of the story being lowered (its first section's
    /// page width less its margins), or `None` for a multi-column section,
    /// whose column width the lowering does not know. Where an
    /// absolute-position tab goes (`ptab_stop`).
    text_width: Option<f32>,
    /// The body block being lowered, for diagnostics.
    current_block: usize,
    /// Symbol characters carried as their Unicode equivalent, per font.
    symbols_carried: BTreeMap<String, usize>,
    /// Symbol characters with no equivalent: (body block, font, code).
    symbols_dropped: Vec<(usize, String, String)>,
}

impl Lowering {
    /// Create the `docx-Default` base paragraph style from docDefaults (if it
    /// carries anything), so every un-based style and un-styled paragraph
    /// inherits Word's document defaults.
    fn install_doc_defaults(&mut self, defaults: &docx_core::Defaults) {
        let mut props = self.para_props(&defaults.para);
        props.extend(self.run_props(&defaults.run));
        if props.is_empty() && defaults.para.line_spacing.is_none() {
            // Nothing to carry: every style and paragraph resolves its own
            // leading (Word's single when nothing sets one).
            return;
        }
        let ls = defaults.para.line_spacing.unwrap_or(SINGLE);
        let single = self.single_line(defaults.run.font.as_deref(), defaults.run.size_half_pts);
        let leading = Some(line_height::line_pitch_pt(ls, single));
        if let Some(pt) = leading {
            props.push(len("characterLeading", pt));
        }
        let id = format!("{PARA_PREFIX}Default");
        if let Some(pt) = leading {
            self.style_leading.insert(id.clone(), pt);
        }
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
                    if s.para.page_break_before == Some(false)
                        && self.resolve_page_break_before(None, s.based_on.as_deref())
                    {
                        props.push(text(START_PARAGRAPH, "Anywhere"));
                    }
                    // Every paragraph style carries the leading it RESOLVES
                    // to: Word's auto/atLeast depend on the style's own
                    // (possibly inherited) font and size, which a child
                    // style can change without touching the spacing.
                    let id = format!("{PARA_PREFIX}{}", sanitize(&s.style_id));
                    if let Some(pt) = self.style_pitch(&s.style_id) {
                        props.push(len("characterLeading", pt));
                        self.style_leading.insert(id.clone(), pt);
                    }
                    props.extend(self.run_props(&s.run));
                    self.record_style(LoweredStyle {
                        id,
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

    /// A table-cell paragraph (breaks in cells are not lowered: Word ignores
    /// a page break inside a table cell's flow only partly, and the native
    /// cell has no break-before).
    fn lower_paragraph(&mut self, p: &docx_core::Paragraph, source_index: u32) -> LoweredParagraph {
        self.lower_body_paragraph(p, source_index, None, None).0
    }

    /// A body paragraph. `start` is a break-before rule the paragraph
    /// receives from outside (its section's odd/even start, or a break that
    /// ended the previous paragraph). `can_carry`: the next block is a
    /// paragraph of the same section, so a break at this paragraph's END is
    /// returned for that paragraph to start with (Word measured: the text
    /// after a trailing break opens the next page with no blank line,
    /// `fixtures/breaks.word.json` A18/A20) instead of splitting this one.
    /// `can_carry` is `None` for a table-cell paragraph, whose breaks are not
    /// lowered.
    fn lower_body_paragraph(
        &mut self,
        p: &docx_core::Paragraph,
        source_index: u32,
        start: Option<&'static str>,
        can_carry: Option<bool>,
    ) -> (LoweredParagraph, Option<BreakKind>) {
        let explicit = p
            .style_id
            .as_ref()
            .map(|id| format!("{PARA_PREFIX}{}", sanitize(id)));
        let mut props = self.para_props(&p.props);
        if p.props.page_break_before == Some(false)
            && self.resolve_page_break_before(None, p.style_id.as_deref())
        {
            props.push(text(START_PARAGRAPH, "Anywhere"));
        }

        // Page / column breaks in the runs, by where they sit in the
        // paragraph's text: at its START (the paragraph itself starts over),
        // at its END (the NEXT paragraph does), or INSIDE it (a split).
        let total: usize = p.runs.iter().map(|r| r.text.chars().count()).sum();
        let mut head: Option<&'static str> = None;
        let mut tail: Option<BreakKind> = None;
        let mut inside: std::collections::BTreeMap<usize, &'static str> = Default::default();
        if can_carry.is_some() {
            let mut offset = 0usize;
            for r in &p.runs {
                let len = r.text.chars().count();
                for b in &r.breaks {
                    self.breaks += 1;
                    let at = offset + b.at.min(len);
                    if at >= total {
                        tail = Some(match tail {
                            Some(BreakKind::Page) => BreakKind::Page,
                            _ => b.kind,
                        });
                    } else if at == 0 {
                        head = stronger(head, Some(break_rule(b.kind)));
                    } else {
                        let rule = break_rule(b.kind);
                        let e = inside.entry(at).or_insert(rule);
                        *e = stronger(Some(*e), Some(rule)).unwrap_or(rule);
                    }
                }
                offset += len;
            }
        }
        if let (Some(kind), Some(false)) = (tail, can_carry) {
            // Nothing in this section to carry the break to (a table follows,
            // or the section ends): the paragraph mark moves instead, as an
            // empty segment that opens the next page/column.
            inside.insert(total, break_rule(kind));
            tail = None;
        }

        // The rule this paragraph itself starts with. A paragraph that
        // already breaks before a page (`pageBreakBefore`) needs nothing
        // weaker on top.
        let own_page =
            self.resolve_page_break_before(p.props.page_break_before, p.style_id.as_deref());
        if let Some(rule) = stronger(start, head) {
            if !own_page || rule_rank(rule) > rule_rank("NextPage") {
                props.retain(|sp| sp.path != START_PARAGRAPH);
                props.push(text(START_PARAGRAPH, rule));
            }
        }
        if let Some(pt) = self.paragraph_pitch(p) {
            let inherited = self
                .para_base(explicit.clone())
                .and_then(|base| self.style_leading.get(&base).copied());
            if inherited.is_none_or(|v| (v - pt).abs() > 0.001) {
                props.push(len("characterLeading", pt));
            }
        }
        if let Some(list) = &p.list {
            props.extend(list_props(list));
        }
        if can_carry.is_some() && self.section_indent != (0.0, 0.0) {
            self.add_section_indent(p, &mut props);
        }
        if can_carry.is_some() {
            props.extend(section_column_props(self.section_columns));
        }
        if let Some(stops) = self.ptab_stop(p, can_carry.is_some()) {
            props.retain(|sp| sp.path != "paragraphTabStops");
            props.push(stops);
        }
        let mut para_style_id = if props.is_empty() {
            // No direct formatting or list: apply the paragraph's style, or the
            // docDefaults base when the paragraph carries no style at all.
            self.para_base(explicit)
        } else {
            let base = self.para_base(explicit);
            Some(self.synthesize(StyleCollection::Paragraph, base, props))
        };

        // A break inside the paragraph splits it into native paragraphs. Word
        // keeps it ONE paragraph, so the parts after a break are not new
        // paragraphs to Word: no space before, no first-line indent, no list
        // marker; and the part before it gets no space after.
        let mut segments = Vec::new();
        if !inside.is_empty() {
            let whole = para_style_id.clone();
            for (&at, &rule) in &inside {
                let mut seg = vec![
                    text(START_PARAGRAPH, rule),
                    len("paragraphSpaceBefore", 0.0),
                    len("paragraphFirstLineIndent", 0.0),
                ];
                if p.list.is_some() {
                    seg.push(text("paragraphListType", "NoList"));
                }
                segments.push(LoweredSegment {
                    at: at as u32,
                    para_style_id: Some(self.synthesize(
                        StyleCollection::Paragraph,
                        whole.clone(),
                        seg,
                    )),
                });
            }
            para_style_id = Some(self.synthesize(
                StyleCollection::Paragraph,
                whole,
                vec![len("paragraphSpaceAfter", 0.0)],
            ));
        }

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

        (
            LoweredParagraph {
                para_style_id,
                runs,
                images,
                source_index,
                segments,
            },
            tail,
        )
    }

    fn lower_run(&mut self, r: &Run) -> LoweredRun {
        for sym in &r.symbols {
            let font = sym.font.clone().unwrap_or_else(|| "no font".into());
            if sym.char.is_some() {
                *self.symbols_carried.entry(font).or_default() += 1;
            } else {
                self.symbols_dropped
                    .push((self.current_block, font, sym.code.clone()));
            }
        }
        let base = r
            .style_id
            .as_ref()
            .map(|id| format!("{CHAR_PREFIX}{}", sanitize(id)));
        let mut props = self.run_props(&r.props);

        // A hyperlink run gets the conventional look — blue + underline — unless
        // the run already sets those directly. The blue/underline is the
        // APPEARANCE; the CLICKABILITY comes from a native `insertHyperlink`
        // over the run range (emitted by the host-model from `hyperlink_url`).
        // Internal `#anchor` targets keep the look but no native link yet (the
        // core door registers URL destinations, not text anchors).
        let external_url = r
            .hyperlink
            .as_deref()
            .filter(|t| !t.starts_with('#'))
            .map(str::to_string);
        if r.hyperlink.is_some() {
            self.hyperlinks += 1;
            if external_url.is_some() {
                self.clickable_links += 1;
            }
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
            hyperlink_url: external_url,
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
        // Word line spacing (`w:spacing/@w:line`) is NOT lowered here: it
        // depends on the font and size the paragraph resolves to, so the
        // callers compute it (`style_pitch`, `paragraph_pitch`; see
        // `line_height` for the rule Word was measured to follow).
        // Word's keepNext is a boolean; paged's keepWithNext is a line count, so
        // "on" maps to a single-line hold.
        if p.keep_next == Some(true) {
            out.push(len("paragraphKeepWithNext", 1.0));
        }
        if let Some(k) = p.keep_lines {
            out.push(boolean("paragraphKeepLinesTogether", k));
        }
        // ADR 028/029 — `w:pageBreakBefore` is InDesign's StartParagraph
        // NextPage: both leave a paragraph that already opens a page where it
        // is (Word measured: `fixtures/breaks.word.json`, A15). An explicit
        // `w:val="0"` only matters over an inherited "on"; the callers, which
        // know the chain, emit `Anywhere` for that.
        if p.page_break_before == Some(true) {
            out.push(text(START_PARAGRAPH, "NextPage"));
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

    /// The tab stop an absolute-position tab (`<w:ptab>`) becomes. A ptab
    /// goes to a position of its own — the left/centre/right of the margins
    /// or of the paragraph's indents — not to the next tab stop. In the
    /// native model, the paragraph's ONE tab goes there when it is the
    /// paragraph's only tab stop, so a paragraph whose only tab is one ptab
    /// gets exactly that stop (its other stops served no tab). Where that
    /// cannot say it — other tabs in the paragraph, a table cell (whose
    /// margins the lowering does not know), a multi-column section — the
    /// ptab stays a plain tab to the next stop, and a diagnostic says so.
    fn ptab_stop(&mut self, p: &docx_core::Paragraph, body: bool) -> Option<StyleProp> {
        let ptabs: Vec<&docx_core::PositionalTab> =
            p.runs.iter().flat_map(|r| r.ptabs.iter()).collect();
        let first = *ptabs.first()?;
        let tabs: usize = p
            .runs
            .iter()
            .map(|r| r.text.chars().filter(|c| *c == '\t').count())
            .sum();
        let block = self.current_block;
        let refuse = |why: &str| {
            Diagnostic::info(
                format!(
                    "the absolute-position tab(s) in body block {block} became plain tabs \
                     to the next tab stop: {why}"
                ),
                3,
            )
        };
        let width = match self.text_width {
            _ if !body => {
                self.diagnostics.push(refuse(
                    "the paragraph is in a table cell, whose margins the lowering does not know",
                ));
                return None;
            }
            _ if tabs > 1 => {
                self.diagnostics.push(refuse(
                    "the paragraph has other tabs, which a tab stop at the ptab's position \
                     would also catch",
                ));
                return None;
            }
            None => {
                self.diagnostics.push(refuse(
                    "the section has several columns, whose width the lowering does not know",
                ));
                return None;
            }
            Some(w) => w,
        };
        let (dl, dr) = self.section_indent;
        let (left, right) = match first.relative_to {
            docx_core::PtabBase::Margin => (dl, width - dr),
            docx_core::PtabBase::Indent => {
                let chain = self.style_chain(p.style_id.as_deref());
                let indent = |pick: fn(&ParaProps) -> Option<i32>| -> f32 {
                    pick(&p.props)
                        .or_else(|| chain.iter().find_map(|s| pick(&s.para)))
                        .or_else(|| pick(&self.word_defaults.para))
                        .map_or(0.0, twip_to_pt)
                };
                (
                    dl + indent(|pp| pp.left_indent),
                    width - dr - indent(|pp| pp.right_indent),
                )
            }
        };
        let (position, alignment) = match first.alignment {
            docx_core::PtabAlignment::Left => (left, "left"),
            docx_core::PtabAlignment::Center => ((left + right) / 2.0, "center"),
            docx_core::PtabAlignment::Right => (right, "right"),
        };
        Some(StyleProp {
            path: "paragraphTabStops".into(),
            value: PropValue::TabStops(vec![LoweredTabStop {
                position,
                alignment: Some(alignment.into()),
                alignment_character: None,
                leader: first.leader.clone(),
            }]),
        })
    }

    /// ADR 029 — a continuous section with other left/right margins joins
    /// the story before it (`sections`): Word measures a paragraph's indents
    /// from its SECTION's margins, so the margin difference adds to the
    /// paragraph's own resolved indents (direct, list, style chain,
    /// docDefaults; the last one pushed wins, as the host applies them).
    fn add_section_indent(&self, p: &docx_core::Paragraph, props: &mut Vec<StyleProp>) {
        let (dl, dr) = self.section_indent;
        let chain = self.style_chain(p.style_id.as_deref());
        let resolved = |path: &str, pick: fn(&ParaProps) -> Option<i32>| -> f32 {
            props
                .iter()
                .rev()
                .find(|sp| sp.path == path)
                .and_then(|sp| match sp.value {
                    PropValue::Length(v) => Some(v),
                    _ => None,
                })
                .or_else(|| chain.iter().find_map(|s| pick(&s.para)).map(twip_to_pt))
                .or_else(|| pick(&self.word_defaults.para).map(twip_to_pt))
                .unwrap_or(0.0)
        };
        let left = resolved("paragraphLeftIndent", |pp| pp.left_indent) + dl;
        let right = resolved("paragraphRightIndent", |pp| pp.right_indent) + dr;
        if dl != 0.0 {
            props.retain(|sp| sp.path != "paragraphLeftIndent");
            props.push(len("paragraphLeftIndent", left));
        }
        if dr != 0.0 {
            props.retain(|sp| sp.path != "paragraphRightIndent");
            props.push(len("paragraphRightIndent", right));
        }
    }

    /// A paragraph style and its `basedOn` ancestors, nearest first
    /// (depth-bounded against malformed cycles).
    fn style_chain(&self, id: Option<&str>) -> Vec<&Style> {
        let mut out = Vec::new();
        let mut next = id;
        while let Some(id) = next {
            let Some(s) = self.word_styles.get(id) else {
                break;
            };
            if out.len() > 32 {
                break;
            }
            out.push(s);
            next = s.based_on.as_deref();
        }
        out
    }

    /// The line spacing a paragraph in `style` resolves to: `direct`, else
    /// the nearest style in the chain that sets it, else docDefaults, else
    /// Word's single (what Word lays when nothing sets it; the engine's own
    /// default, 120% of the point size, is not Word's).
    fn resolve_spacing(
        &self,
        direct: Option<LineSpacing>,
        style: Option<&str>,
    ) -> Option<LineSpacing> {
        direct
            .or_else(|| {
                self.style_chain(style)
                    .iter()
                    .find_map(|s| s.para.line_spacing)
            })
            .or(self.word_defaults.para.line_spacing)
            .or(Some(SINGLE))
    }

    /// Whether a paragraph with direct `pageBreakBefore` `direct` in `style`
    /// starts a new page: direct, else the nearest style that sets it.
    fn resolve_page_break_before(&self, direct: Option<bool>, style: Option<&str>) -> bool {
        direct
            .or_else(|| {
                self.style_chain(style)
                    .iter()
                    .find_map(|s| s.para.page_break_before)
            })
            .unwrap_or(false)
    }

    /// The font and size (half-points) a run resolves to: its direct props,
    /// its character style chain, the paragraph style chain, docDefaults.
    fn resolve_face(
        &self,
        run: Option<&Run>,
        para_style: Option<&str>,
    ) -> (Option<String>, Option<u32>) {
        let mut layers: Vec<&RunProps> = Vec::new();
        if let Some(r) = run {
            layers.push(&r.props);
            layers.extend(
                self.style_chain(r.style_id.as_deref())
                    .into_iter()
                    .map(|s| &s.run),
            );
        }
        layers.extend(self.style_chain(para_style).into_iter().map(|s| &s.run));
        layers.push(&self.word_defaults.run);
        let font = layers.iter().find_map(|l| l.font.clone());
        let size = layers.iter().find_map(|l| l.size_half_pts);
        (font, size)
    }

    /// Word's single line in points for `font` at `half_pts` (Word's default
    /// size is 10 pt), noting faces Word has not been measured on.
    fn single_line(&mut self, font: Option<&str>, half_pts: Option<u32>) -> f32 {
        if let Some(f) = font {
            if line_height::measured_single_em(f).is_none() {
                self.unmeasured_faces.insert(f.to_string());
            }
        }
        line_height::single_line_pt(font, half_pts.map_or(10.0, half_pt_to_pt))
    }

    /// The leading a paragraph style resolves to, if line spacing is set
    /// anywhere in its chain (or in docDefaults).
    fn style_pitch(&mut self, style_id: &str) -> Option<f32> {
        let ls = self.resolve_spacing(None, Some(style_id))?;
        let (font, size) = self.resolve_face(None, Some(style_id));
        let single = self.single_line(font.as_deref(), size);
        Some(line_height::line_pitch_pt(ls, single))
    }

    /// The leading Word lays a paragraph at: its resolved line spacing over
    /// the single line of its TALLEST run (Word sizes each line by its
    /// tallest run; one leading per paragraph is the native model's grain).
    fn paragraph_pitch(&mut self, p: &docx_core::Paragraph) -> Option<f32> {
        let style = p.style_id.as_deref();
        let ls = self.resolve_spacing(p.props.line_spacing, style)?;
        let faces: Vec<(Option<String>, Option<u32>)> = {
            let mut v: Vec<_> = p
                .runs
                .iter()
                .filter(|r| !r.text.is_empty())
                .map(|r| self.resolve_face(Some(r), style))
                .collect();
            if v.is_empty() {
                v.push(self.resolve_face(None, style));
            }
            v
        };
        let single = faces
            .iter()
            .map(|(font, size)| self.single_line(font.as_deref(), *size))
            .fold(0.0_f32, f32::max);
        Some(line_height::line_pitch_pt(ls, single))
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

fn text(path: &str, v: &str) -> StyleProp {
    StyleProp {
        path: path.into(),
        value: PropValue::Text(v.into()),
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
        first_block: s.first_block,
        story: 0,
    }
}

/// The span/split column properties of a paragraph in a section that sits
/// `columns` in its story (empty in the story's own columns). Min spaces are
/// left at the engine's 0: Word starts a section's text right under the
/// previous section's deepest line (`fixtures/continuous.word.json`).
fn section_column_props(columns: SectionColumns) -> Vec<StyleProp> {
    match columns {
        SectionColumns::Frame => Vec::new(),
        SectionColumns::SpanAll => vec![
            text(SPAN_COLUMN_TYPE, "SpanColumns"),
            text(SPAN_SPLIT_COLUMN_COUNT, "All"),
        ],
        SectionColumns::Split { count, inside_pt } => vec![
            text(SPAN_COLUMN_TYPE, "SplitColumns"),
            text(SPAN_SPLIT_COLUMN_COUNT, &count.to_string()),
            len(SPLIT_COLUMN_INSIDE_GUTTER, inside_pt),
            len(SPLIT_COLUMN_OUTSIDE_GUTTER, 0.0),
        ],
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

/// The lowered character-style token for a Word character `style_id` (the same
/// id `lower` mints, so save-back can build a `token → styleId` map to recover a
/// real `w:rStyle` — `sanitize` is lossy, so the map is the only safe inverse).
pub fn char_style_token(style_id: &str) -> String {
    format!("{CHAR_PREFIX}{}", sanitize(style_id))
}

/// The lowered PARAGRAPH-style token for a Word paragraph `style_id` (the twin of
/// [`char_style_token`]; `sanitize` is lossy, so save-back needs the map).
pub fn para_style_token(style_id: &str) -> String {
    format!("{PARA_PREFIX}{}", sanitize(style_id))
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
    #![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature
    use super::*;
    use docx_core::*;

    fn run(text: &str, props: RunProps) -> Run {
        Run {
            style_id: None,
            props,
            image: None,
            hyperlink: None,
            text: text.into(),
            ..Default::default()
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
            source_para_ord: 0,
            source_cell: None,
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

    fn ptab_run(text: &str, at: usize) -> Run {
        Run {
            ptabs: vec![PositionalTab {
                at,
                alignment: PtabAlignment::Right,
                relative_to: PtabBase::Margin,
                leader: None,
            }],
            ..run(text, RunProps::default())
        }
    }

    fn ptab_notes(l: &LoweredDoc) -> Vec<&str> {
        l.diagnostics
            .iter()
            .filter(|d| d.message.contains("absolute-position"))
            .map(|d| d.message.as_str())
            .collect()
    }

    #[test]
    fn a_ptab_lowers_to_one_tab_stop_at_the_right_margin__feat__plugin_doc_read_path() {
        let mut doc = DocxDocument::default();
        doc.sections.push(Section::default()); // 612 pt, 72 pt margins
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![ptab_run("a\tb", 1)],
            ..Default::default()
        }));
        let l = lower(&doc);
        let id = l.story.paragraphs()[0].para_style_id.clone().unwrap();
        let st = l.styles.iter().find(|s| s.id == id).unwrap();
        let stops = st
            .props
            .iter()
            .find_map(|p| match &p.value {
                PropValue::TabStops(t) => Some(t.clone()),
                _ => None,
            })
            .unwrap();
        assert_eq!(stops.len(), 1);
        assert_eq!(stops[0].position, 468.0);
        assert_eq!(stops[0].alignment.as_deref(), Some("right"));
        assert!(ptab_notes(&l).is_empty());
    }

    #[test]
    fn a_ptab_the_native_model_cannot_place_stays_a_tab_and_says_so__feat__plugin_doc_read_path() {
        // Another tab in the paragraph.
        let mut doc = DocxDocument::default();
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![ptab_run("a\tb\tc", 1)],
            ..Default::default()
        }));
        let l = lower(&doc);
        assert!(
            ptab_notes(&l)[0].contains("other tabs"),
            "{:?}",
            l.diagnostics
        );
        assert_eq!(l.story.paragraphs()[0].runs[0].text, "a\tb\tc");

        // Two columns.
        let mut doc = DocxDocument::default();
        doc.sections.push(Section {
            columns: 2,
            ..Section::default()
        });
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![ptab_run("a\tb", 1)],
            ..Default::default()
        }));
        assert!(ptab_notes(&lower(&doc))[0].contains("several columns"));

        // A table cell.
        let mut doc = DocxDocument::default();
        doc.body.push(Block::Table(Table {
            column_widths: vec![2000],
            rows: vec![TableRow {
                cells: vec![TableCell {
                    paragraphs: vec![Paragraph {
                        runs: vec![ptab_run("a\tb", 1)],
                        ..Default::default()
                    }],
                    grid_span: 1,
                    v_merge: VMerge::None,
                }],
            }],
        }));
        let l = lower(&doc);
        assert!(
            ptab_notes(&l)[0].contains("table cell"),
            "{:?}",
            l.diagnostics
        );
    }

    #[test]
    fn symbols_are_counted_and_the_ones_without_a_character_warned__feat__plugin_doc_read_path() {
        let mut doc = DocxDocument::default();
        let sym = |font: &str, code: &str, char: Option<char>| RunSymbol {
            at: 0,
            font: Some(font.into()),
            code: code.into(),
            char,
        };
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![Run {
                symbols: vec![
                    sym("Symbol", "F0B7", Some('\u{2022}')),
                    sym("Wingdings", "F0FF", None),
                ],
                ..run("\u{2022}", RunProps::default())
            }],
            ..Default::default()
        }));
        let l = lower(&doc);
        assert!(l
            .diagnostics
            .iter()
            .any(|d| d.severity == "info" && d.message.contains("1 from Symbol")));
        assert!(l.diagnostics.iter().any(
            |d| d.severity == "warning" && d.message.contains("body block 0 (Wingdings F0FF)")
        ));
    }
}
