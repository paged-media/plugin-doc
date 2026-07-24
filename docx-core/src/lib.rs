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

//! # docx-core — the frozen semantic WordprocessingML view
//!
//! A small, plugin-owned model of the parts of a `.docx` that Tier-0/1 lowering
//! reads: the document **body** (paragraphs and — as a Tier-2 stub — tables), the
//! **style catalog** (`styles.xml`), and **section** page geometry. It sits one
//! step away from `ooxmlsdk`'s code-generated typed DOM: `docx-import` maps the
//! enum-vector `ooxmlsdk` trees into these clean structs, and `docx-lower` reads
//! *only* these structs, so the lowering never touches `ooxmlsdk` directly (the
//! "wrap it, don't expose it raw" discipline of the spec §5.2).
//!
//! Units are Word's native units, unconverted: **twips** (1/1440 inch) for
//! lengths, **half-points** for font size, `RRGGBB` hex for colors. Conversion to
//! Paged's points happens in `docx-lower`.
//!
//! Everything is `Default` + `serde` so `docx-import` can build it incrementally
//! and tests can assert on it.

use serde::{Deserialize, Serialize};

/// A parsed Word document: the body in reading order, the style catalog, and the
/// sections (page geometry).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocxDocument {
    /// Body content in document order.
    pub body: Vec<Block>,
    /// The style catalog from `styles.xml` (`docDefaults` + named styles).
    pub styles: StyleCatalog,
    /// Section page geometry (Tier-1 partial). At least one is synthesized if the
    /// document omits `sectPr`.
    pub sections: Vec<Section>,
}

/// A top-level body block.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    /// A paragraph (`w:p`).
    Paragraph(Paragraph),
    /// A table (`w:tbl`) — Tier-2 stub; carried structurally, lowered minimally.
    Table(Table),
}

/// A Word paragraph (`w:p`): an applied paragraph style, direct paragraph
/// properties, and a sequence of runs.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Paragraph {
    /// `w:pPr/w:pStyle/@w:val` — the applied paragraph style id, if any.
    pub style_id: Option<String>,
    /// Direct paragraph formatting (`w:pPr`).
    pub props: ParaProps,
    /// The runs (`w:r`) in order. Non-run inline content (hyperlinks, fields) is
    /// flattened to its runs for Tier-0.
    pub runs: Vec<Run>,
}

/// A Word run (`w:r`): direct character formatting plus its text.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Run {
    /// `w:rPr/w:rStyle/@w:val` — the applied character style id, if any.
    pub style_id: Option<String>,
    /// Direct character formatting (`w:rPr`).
    pub props: RunProps,
    /// The concatenated text of the run's `w:t` children (tabs/breaks preserved
    /// as `\t` / `\n`).
    pub text: String,
}

/// Direct paragraph formatting. `None` means "inherit"; all lengths in twips.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct ParaProps {
    pub justification: Option<Justification>,
    pub left_indent: Option<i32>,
    pub right_indent: Option<i32>,
    pub first_line_indent: Option<i32>,
    /// A hanging indent (`w:ind/@w:hanging`) — stored as a positive twip value; it
    /// is the negative of a first-line indent.
    pub hanging_indent: Option<i32>,
    pub space_before: Option<i32>,
    pub space_after: Option<i32>,
    pub keep_next: Option<bool>,
    pub keep_lines: Option<bool>,
}

/// Direct character formatting. `None` means "inherit".
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct RunProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    pub caps: Option<bool>,
    pub small_caps: Option<bool>,
    /// `w:color/@w:val` as `RRGGBB` (never the sentinel `auto`).
    pub color: Option<String>,
    /// `w:sz/@w:val` in half-points.
    pub size_half_pts: Option<u32>,
    /// `w:rFonts/@w:ascii` — the primary Latin font family.
    pub font: Option<String>,
    pub vert_align: Option<VertAlign>,
}

/// Paragraph alignment (`w:jc`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Justification {
    Left,
    Center,
    Right,
    Both,
    Distribute,
    Start,
    End,
}

/// Run vertical alignment (`w:vertAlign`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum VertAlign {
    Baseline,
    Superscript,
    Subscript,
}

/// The style catalog (`styles.xml`).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct StyleCatalog {
    /// `docDefaults` — the document-wide default paragraph + run properties.
    pub doc_defaults: Defaults,
    /// Named styles in document order (paragraph, character, table, numbering).
    pub styles: Vec<Style>,
}

/// `w:docDefaults`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Defaults {
    pub para: ParaProps,
    pub run: RunProps,
}

/// A named Word style (`w:style`).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Style {
    /// `@w:styleId`.
    pub style_id: String,
    /// `w:name/@w:val` — the human-facing name (falls back to the id).
    pub name: Option<String>,
    /// `@w:type`.
    pub kind: StyleKind,
    /// `w:basedOn/@w:val`.
    pub based_on: Option<String>,
    /// Paragraph-level properties defined by the style (`w:pPr`).
    pub para: ParaProps,
    /// Character-level properties defined by the style (`w:rPr`).
    pub run: RunProps,
}

/// `w:style/@w:type`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum StyleKind {
    #[default]
    Paragraph,
    Character,
    Table,
    Numbering,
}

/// A table (`w:tbl`) — Tier-2 stub. Rows of cells, each cell holding paragraphs.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<TableRow>,
}

/// A table row (`w:tr`).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

/// A table cell (`w:tc`) — holds block content (paragraphs).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<Paragraph>,
}

/// Section page geometry (`w:sectPr`). Lengths in twips.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Section {
    /// `w:pgSz/@w:w`.
    pub page_width: i32,
    /// `w:pgSz/@w:h`.
    pub page_height: i32,
    pub margin_top: i32,
    pub margin_bottom: i32,
    pub margin_left: i32,
    pub margin_right: i32,
    /// `w:cols/@w:num` — column count (default 1).
    pub columns: u32,
}

impl Default for Section {
    /// US Letter, 1-inch margins, single column — the Word default when `sectPr`
    /// is absent.
    fn default() -> Self {
        Section {
            page_width: 12240,
            page_height: 15840,
            margin_top: 1440,
            margin_bottom: 1440,
            margin_left: 1440,
            margin_right: 1440,
            columns: 1,
        }
    }
}
