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
    /// The first section's page geometry (points).
    pub section: LoweredSection,
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
    /// `insertAnchoredFrame` at the paragraph's story offset).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub images: Vec<LoweredImage>,
    /// Provenance: the index of the source body block, kept for future
    /// targeted save-back (M2). Not used for rendering.
    pub source_index: u32,
}

/// An inline image lowered to an anchored-frame placement.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredImage {
    pub width_pt: f32,
    pub height_pt: f32,
    /// A self-contained `data:<mime>;base64,…` URI the anchored frame links to.
    pub uri: String,
}

/// A run: its text and an effective (Word or synthesized) character style
/// applied over the run range.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredRun {
    pub text: String,
    /// Full `CharacterStyle/…` token to `applyStyle` over the run, or `None`.
    pub char_style_id: Option<String>,
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
}
