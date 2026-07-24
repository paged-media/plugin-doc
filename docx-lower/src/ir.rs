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
}

/// The body as one native story.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredStory {
    pub paragraphs: Vec<LoweredParagraph>,
}

/// A paragraph: an effective (Word or synthesized) paragraph style applied over
/// the paragraph range, plus its runs.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LoweredParagraph {
    /// Full `ParagraphStyle/…` token to `applyStyle` over the paragraph, or
    /// `None` to leave the default.
    pub para_style_id: Option<String>,
    pub runs: Vec<LoweredRun>,
    /// Provenance: the index of the source body block, kept for future
    /// targeted save-back (M2). Not used for rendering.
    pub source_index: u32,
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
