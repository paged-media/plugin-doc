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

//! The **Lowered IR** — the contract between `docx-lower` (Rust) and
//! `@paged-media/doc-host-model` (TS).
//!
//! Every id in the IR is a fully-formed Paged token (`ParagraphStyle/docx-…`,
//! `CharacterStyle/docx-…`, `Color/docx-…`) so the host-model is a *dumb*
//! translator: it never invents ids, it only maps IR nodes to
//! `host.document.mutate(...)` ops. [`PropValue`] serializes to exactly the wire
//! `Value` union (`{ "type": "text"|"length"|"bool"|"colorRef", "value": … }`),
//! so a `StyleProp.value` drops straight into a `setStyleProperty`/`applyStyle`
//! payload with no re-shaping on the TS side.
//!
//! Serialized to JSON by `docx-js` and consumed verbatim by the bundle.

use serde::{Deserialize, Serialize};

/// The whole lowering of one Word document body.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredDoc {
    /// Colors to create (`createSwatch`) before styles reference them.
    pub swatches: Vec<LoweredSwatch>,
    /// Style catalog to create, **topologically ordered** so every `basedOn`
    /// parent precedes its children (Word styles + synthesized direct-format
    /// styles).
    pub styles: Vec<LoweredStyle>,
    /// The body poured as a single native story of paragraphs.
    pub story: LoweredStory,
    /// The first section's page geometry (points). Embedded placement uses it.
    pub section: LoweredSection,
    /// ADR 029 — EVERY section, in order, with its first block. Standalone
    /// open pours each section's blocks into that section's own story.
    #[serde(default)]
    pub sections: Vec<LoweredSection>,
    /// ADR 033 — `settings.xml` `w:evenAndOddHeaders`: even pages show the
    /// `even` header and footer.
    #[serde(default)]
    pub even_and_odd_headers: bool,
    /// Honest ADR-007 diagnostics for anything not lowered natively this pass.
    pub diagnostics: Vec<Diagnostic>,
}

/// A color to mint via `createSwatch`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredSwatch {
    /// `Color/docx-RRGGBB`.
    pub id: String,
    pub name: String,
    /// `"RGB"` this pass (CMYK is a later tier).
    pub space: String,
    /// Channel values in `space` — `[r, g, b]` on 0–255 (IDML convention).
    pub value: Vec<f32>,
}

/// A native style to create via `create{Paragraph,Character}Style` +
/// `setStyleProperty` per prop.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredStyle {
    /// Full token, e.g. `ParagraphStyle/docx-Heading1`.
    pub id: String,
    pub name: String,
    pub collection: StyleCollection,
    /// Another style's full token, or `None`.
    pub based_on: Option<String>,
    pub props: Vec<StyleProp>,
}

/// Which style collection a [`LoweredStyle`] belongs to (mirrors the host
/// `StyleCollection` wire strings we use this pass).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum StyleCollection {
    Paragraph,
    Character,
}

/// One `setStyleProperty` (or `applyStyle`-adjacent) property assignment.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StyleProp {
    /// A `PropertyPath` wire string, e.g. `"characterFontStyle"`.
    pub path: String,
    pub value: PropValue,
}

/// A property value that serializes to the host wire `Value` union verbatim.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum PropValue {
    Text(String),
    Length(f32),
    Bool(bool),
    ColorRef(String),
    /// `{ "type": "tabStops", "value": [TabStopSpec…] }`.
    TabStops(Vec<LoweredTabStop>),
}

/// A tab stop, shaped as the host `TabStopSpec` (position in points).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredTabStop {
    pub position: f32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment_character: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub leader: Option<String>,
}

/// The body as one native story: a sequence of blocks (paragraphs + tables) in
/// document order.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredStory {
    pub blocks: Vec<LoweredBlock>,
}

/// One top-level story block.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum LoweredBlock {
    Paragraph(LoweredParagraph),
    Table(LoweredTable),
}

/// A native table to build via `insertTable` + per-cell `insertText` +
/// `setCellSpan`. `rows`/`cols` size the grid; `cells` are the non-absorbed
/// (non-`vMerge`-continue) cells with their resolved grid position + spans.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredTable {
    pub rows: u32,
    pub cols: u32,
    /// Column widths in points (may be empty ⇒ let the engine auto-size).
    pub column_widths_pt: Vec<f32>,
    pub cells: Vec<LoweredCell>,
    /// Leading rows Word repeats on every page (`w:tblHeader`): the native
    /// table's header rows.
    #[serde(default)]
    pub header_rows: u32,
    /// Each row's least height in points (`w:trHeight`), 0 for a row that
    /// declares none; empty when no row does.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_heights_pt: Vec<f32>,
    /// When non-empty, the table is NOT built as a native table: it has a
    /// row taller than its page, which Word splits across pages and a
    /// native row never does (the whole row would be overset, its text
    /// gone). Its cells' paragraphs, in reading order, pour as body text
    /// instead; `rows` / `cols` / `cells` are then empty. Save-back leaves
    /// the Word table as it was.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub flow: Vec<LoweredParagraph>,
}

/// One table cell, addressed by its resolved grid position.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredCell {
    pub row: u32,
    pub col: u32,
    pub row_span: u32,
    pub col_span: u32,
    /// The cell's block content, lowered as paragraphs.
    pub paragraphs: Vec<LoweredParagraph>,
    /// Native cell insets (top, left, bottom, right, points) that give the
    /// cell Word's height and text width; see `Lowering::cell_insets`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub insets_pt: Option<[f32; 4]>,
    /// `w:vAlign` as the native vertical justification (`TopAlign` /
    /// `CenterAlign` / `BottomAlign`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<String>,
}

impl LoweredStory {
    /// Just the paragraph blocks, in order (convenience for tests/consumers that
    /// only care about body text; skips table blocks).
    pub fn paragraphs(&self) -> Vec<&LoweredParagraph> {
        self.blocks
            .iter()
            .filter_map(|b| match b {
                LoweredBlock::Paragraph(p) => Some(p),
                LoweredBlock::Table(_) => None,
            })
            .collect()
    }
}

/// A paragraph: an effective (Word or synthesized) paragraph style applied over
/// the paragraph range, plus its runs and any inline images.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredParagraph {
    /// Full `ParagraphStyle/…` token to `applyStyle` over the paragraph, or
    /// `None` to leave the default.
    pub para_style_id: Option<String>,
    pub runs: Vec<LoweredRun>,
    /// Inline images anchored to this paragraph (rendered via
    /// `insertAnchoredFrame` at the paragraph's story offset plus each
    /// image's [`LoweredImage::at`]).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub images: Vec<LoweredImage>,
    /// Provenance: the index of the source body block, kept for future
    /// targeted save-back (M2). Not used for rendering.
    pub source_index: u32,
    /// ADR 028/029 — where a page or column break INSIDE the Word paragraph
    /// splits it. The native model breaks only between paragraphs, so the pour
    /// starts a new native paragraph at each segment's `at`, styled with the
    /// segment's style (which carries the break-before rule). Empty for an
    /// unsplit paragraph. The block stays ONE block (one Word `w:p`), so
    /// save-back provenance is unchanged: the read-back folds the segments'
    /// native paragraphs back into this one.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub segments: Vec<LoweredSegment>,
}

/// The second or later part of a Word paragraph that a page or column break
/// splits (see [`LoweredParagraph::segments`]).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredSegment {
    /// The contiguous char offset into the paragraph's run text where this
    /// segment starts (it runs to the next segment's `at`, or the end).
    pub at: u32,
    /// Full `ParagraphStyle/…` token applied over this segment.
    pub para_style_id: Option<String>,
}

/// An image lowered to an anchored-frame placement.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredImage {
    /// The contiguous char offset into the paragraph's run text where the
    /// picture is addressed: where it sits in Word's paragraph (the chars of
    /// the runs before it), so the engine lays it out as a character of that
    /// line. The engine's contiguous address space cannot name a paragraph's
    /// END (that offset is the next paragraph's start), so a picture after
    /// the paragraph's last character is addressed one character earlier,
    /// and one alone in an empty paragraph falls to the next paragraph's
    /// start; the lowering reports both.
    #[serde(default)]
    pub at: u32,
    pub width_pt: f32,
    pub height_pt: f32,
    /// A self-contained `data:<mime>;base64,…` URI the anchored frame links to.
    pub uri: String,
    /// For a FLOATING Word drawing (`wp:anchor`), where Word positions it and
    /// how text wraps around it (thoughts ADR 035). It is still placed
    /// INLINE (at [`Self::at`]), with a diagnostic, until the engine
    /// can create a positioned, wrapped anchored object (RFI C-47/C-48);
    /// this carries what that later lowering needs. Absent for an inline
    /// picture.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub float: Option<LoweredFloat>,
}

/// A floating drawing's position and wrap, in points (thoughts ADR 035).
/// Word's vocabulary is kept as written (`relativeFrom`, `wrapSquare`,
/// `bothSides`): the mapping onto anchored-object settings is decided by the
/// lowering that places it.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredFloat {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<LoweredFloatPosition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<LoweredFloatPosition>,
    /// `wrapNone` / `wrapSquare` / `wrapTight` / `wrapThrough` /
    /// `wrapTopAndBottom`.
    pub wrap: String,
    /// `bothSides` / `left` / `right` / `largest` (square, tight, through).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    /// Distance from the text, top / bottom / left / right.
    pub dist_top_pt: f32,
    pub dist_bottom_pt: f32,
    pub dist_left_pt: f32,
    pub dist_right_pt: f32,
    pub behind_doc: bool,
    pub allow_overlap: bool,
    pub layout_in_cell: bool,
    pub locked: bool,
    /// z-order among the floats (higher is in front).
    pub relative_height: u32,
    /// `wp:simplePos` (x, y) from the page's top-left, when it is used.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub simple_pos_pt: Option<(f32, f32)>,
}

/// One axis of a float's position.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredFloatPosition {
    /// `page`, `margin`, `column`, `character`, `paragraph`, `line`, …
    pub relative_from: String,
    /// `wp:posOffset`, in points.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset_pt: Option<f32>,
    /// `wp:align`: `left` / `center` / `right` / `inside` / `outside` /
    /// `top` / `bottom`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    /// `wp14:pctPos*Offset`, in percent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub percent: Option<f32>,
}

/// A run: its text and an effective (Word or synthesized) character style
/// applied over the run range.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredRun {
    pub text: String,
    /// Full `CharacterStyle/…` token to `applyStyle` over the run, or `None`.
    pub char_style_id: Option<String>,
    /// When the run is a hyperlink, its resolved target URL. The host-model
    /// emits an `insertHyperlink` over the run range so it becomes a native
    /// clickable link (the blue+underline look still comes from `char_style_id`;
    /// this only carries the click target). `None` for ordinary runs.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyperlink_url: Option<String>,
}

/// Page geometry for the first section, in points.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredSection {
    pub page_width_pt: f32,
    pub page_height_pt: f32,
    pub margin_top_pt: f32,
    pub margin_bottom_pt: f32,
    pub margin_left_pt: f32,
    pub margin_right_pt: f32,
    pub columns: u32,
    /// ADR 029 — index into `story.blocks` of this section's first block
    /// (blocks map 1:1 to Word body blocks); it runs to the next section's.
    #[serde(default)]
    pub first_block: usize,
    /// ADR 029 — the native story (skeleton page) this section pours into.
    /// Consecutive sections share one when Word continues the later on the
    /// same page (`continuous` / `nextColumn`, see `sections`).
    #[serde(default)]
    pub story: usize,
    /// ADR 033 — the headers and footers this section shows (after Word's
    /// inheritance), its `titlePg`, numbering and distances. Carried, not
    /// yet placed (a diagnostic says so).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<LoweredHeaderFooter>,
    /// ADR 034 — the footnote numbering in force in this section: its own
    /// `w:footnotePr` alone (Word takes nothing from `settings.xml` or the
    /// previous section, `fixtures/footnote-numbering.word.json`). Absent
    /// when the section says nothing (Word's defaults: decimal from 1,
    /// continuous, page bottom).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_numbering: Option<LoweredNoteNumbering>,
    /// ADR 034 — the endnote numbering in force in this section.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_numbering: Option<LoweredNoteNumbering>,
}

/// One section's headers and footers (thoughts ADR 033): which header/footer
/// part each page kind shows, by part name (`word/header1.xml`).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredHeaderFooter {
    pub header: LoweredHeaderFooterParts,
    pub footer: LoweredHeaderFooterParts,
    /// `w:titlePg`: the section's first page shows the `first` pair.
    pub title_page: bool,
    /// `w:pgMar/@w:header`: the header's distance from the page top.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_distance_pt: Option<f32>,
    /// `w:pgMar/@w:footer`: the footer's distance from the page bottom.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_distance_pt: Option<f32>,
    /// `w:pgNumType/@w:start`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_number_start: Option<i32>,
    /// `w:pgNumType/@w:fmt` (`decimal`, `lowerRoman`, …).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_number_format: Option<String>,
}

/// The header (or footer) part of each kind; `None` is blank.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredHeaderFooterParts {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even: Option<String>,
}

/// Footnote or endnote numbering (thoughts ADR 034), Word's vocabulary. A
/// field is absent when neither the section nor the document says it (Word's
/// default applies).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredNoteNumbering {
    /// `decimal`, `lowerRoman`, `upperLetter`, `chicago`, …
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<u32>,
    /// `continuous` / `eachSect` / `eachPage`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<String>,
    /// `pageBottom` / `beneathText` / `sectEnd` / `docEnd`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<String>,
}

impl Default for LoweredSection {
    fn default() -> Self {
        // US Letter, 1-inch margins, single column (Word default), in points.
        LoweredSection {
            page_width_pt: 612.0,
            page_height_pt: 792.0,
            margin_top_pt: 72.0,
            margin_bottom_pt: 72.0,
            margin_left_pt: 72.0,
            margin_right_pt: 72.0,
            columns: 1,
            first_block: 0,
            story: 0,
            header_footer: None,
            footnote_numbering: None,
            endnote_numbering: None,
        }
    }
}

/// An honest diagnostic (ADR-007) — a construct not lowered natively this pass.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Diagnostic {
    /// `"info" | "warning" | "error"`.
    pub severity: String,
    pub message: String,
    /// The fidelity tier the construct belongs to (0–4).
    pub tier: u8,
}

impl Diagnostic {
    pub fn info(message: impl Into<String>, tier: u8) -> Self {
        Diagnostic {
            severity: "info".into(),
            message: message.into(),
            tier,
        }
    }

    pub fn warning(message: impl Into<String>, tier: u8) -> Self {
        Diagnostic {
            severity: "warning".into(),
            message: message.into(),
            tier,
        }
    }
}
