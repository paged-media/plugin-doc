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

pub mod symbols;
pub use symbols::symbol_char;

/// A parsed Word document: the body in reading order, the style catalog, and the
/// sections (page geometry).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocxDocument {
    /// Footnotes + endnotes from the notes parts, keyed by `w:id`
    /// ([`Run::note_ref`]). Empty when the document has none.
    #[serde(default)]
    pub notes: Vec<Note>,
    /// Header + footer parts referenced by the document's sections, one per
    /// PART (two references to one part share it). Each section says which
    /// of them it shows ([`Section::headers`] / [`Section::footers`]).
    #[serde(default)]
    pub headers_footers: Vec<HeaderFooter>,
    /// `settings.xml` `w:evenAndOddHeaders` (thoughts ADR 033): even pages
    /// show the `even` header and footer. Document-wide, unlike `titlePg`.
    #[serde(default)]
    pub even_and_odd_headers: bool,
    /// `settings.xml` `w:footnotePr`, as written (thoughts ADR 034). Word
    /// does NOT number by it: the marks follow each section's own
    /// `w:footnotePr` ([`DocxDocument::footnote_numbering`],
    /// `fixtures/footnote-numbering.word.json`). Kept for save-back and for
    /// a later reader that finds a use for it.
    #[serde(default)]
    pub footnote_props: NoteProps,
    /// `settings.xml` `w:endnotePr`, as [`DocxDocument::footnote_props`].
    #[serde(default)]
    pub endnote_props: NoteProps,
    /// Body content in document order.
    pub body: Vec<Block>,
    /// The style catalog from `styles.xml` (`docDefaults` + named styles).
    pub styles: StyleCatalog,
    /// Section page geometry (Tier-1 partial). At least one is synthesized if the
    /// document omits `sectPr`.
    pub sections: Vec<Section>,
    /// `w:settings/w:autoHyphenation`: Word hyphenates automatically only when
    /// the document turns it on (off by default, unlike the native engine).
    #[serde(default)]
    pub auto_hyphenation: bool,
    /// Legacy VML drawings (`w:pict`) the import could not place: floating
    /// shapes, text boxes, diagrams (only an inline VML picture becomes an
    /// [`Image`]). The lowering reports them.
    #[serde(default)]
    pub unplaced_vml: u32,
}

/// A top-level body block.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    /// A paragraph (`w:p`).
    Paragraph(Paragraph),
    /// A table (`w:tbl`) — Tier-2 stub; carried structurally, lowered minimally.
    Table(Table),
}

/// Where a TABLE-CELL paragraph lives in `word/document.xml` — the source path
/// the save-back patcher walks (`w:tbl` → `w:tr` → `w:tc` → `w:p`). All ordinals
/// are 0-based positions among siblings of that element type. `None` on body
/// paragraphs (they use [`Paragraph::source_para_ord`] instead).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct CellPath {
    /// The table's ordinal among the direct `<w:tbl>` children of `<w:body>`.
    pub table_ord: u32,
    /// The row's ordinal within the table.
    pub row: u32,
    /// The cell's ordinal within the row.
    pub cell: u32,
    /// The paragraph's ordinal within the cell.
    pub para: u32,
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
    /// `w:pPr/w:numPr` resolved through `numbering.xml` — the list marker this
    /// paragraph belongs to, if any.
    pub list: Option<ListMarker>,
    /// Provenance for M2 save-back: this paragraph's ordinal among the direct
    /// `<w:p>` children of `<w:body>` (0-based). Meaningless for table-cell
    /// paragraphs — those carry [`Paragraph::source_cell`] instead.
    #[serde(default)]
    pub source_para_ord: u32,
    /// Provenance for a TABLE-CELL paragraph: its `w:tbl`/`w:tr`/`w:tc`/`w:p`
    /// source path. `None` for body paragraphs.
    #[serde(default)]
    pub source_cell: Option<CellPath>,
}

/// A footnote or endnote (`w:footnote` / `w:endnote` in the notes part), keyed by
/// the `w:id` its in-text reference carries.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Note {
    /// `@w:id` — matches [`Run::note_ref`].
    pub id: i64,
    /// `true` for an endnote, `false` for a footnote.
    pub endnote: bool,
    /// The note's body paragraphs.
    pub paragraphs: Vec<Paragraph>,
}

/// A list marker, resolved from `w:numPr` + `numbering.xml` at import time so the
/// lowering stays pure (no `numbering.xml` access downstream).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ListMarker {
    pub kind: ListKind,
    /// The zero-based indent level (`w:ilvl`).
    pub level: u8,
    /// For bullets: the glyph from `w:lvlText` (e.g. `"•"`).
    pub bullet_char: Option<String>,
    /// For numbered lists: the IDML numbering-format sample (e.g. `"1, 2, 3, 4..."`,
    /// `"I, II, III, IV..."`), matching what the engine's `format_number` reads.
    pub number_format: Option<String>,
    /// The level's own `w:pPr/w:ind` (numbering.xml): where Word puts the
    /// list text (`left`) and the marker (`hanging` / `firstLine`), in twips.
    #[serde(default)]
    pub left_indent: Option<i32>,
    #[serde(default)]
    pub first_line_indent: Option<i32>,
    #[serde(default)]
    pub hanging_indent: Option<i32>,
}

/// Whether a list paragraph is bulleted or numbered.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ListKind {
    Bullet,
    Numbered,
}

/// Where a run came from in `word/document.xml` — the provenance the M2 save-back
/// patcher needs to locate the exact `<w:r>` to rewrite. Only a `DirectRun`
/// (a direct `<w:r>` child of the `<w:p>`, at ordinal `n`) is patchable in the
/// current increment; runs flattened out of `<w:hyperlink>`/`<w:fldSimple>`/complex
/// fields sit on a different locator path and are marked non-patchable.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum RunSource {
    /// The `n`-th direct `<w:r>` child of the paragraph (0-based). A COMPLEX
    /// field's result run is one of these — its URL lives in a separate
    /// `instrText` run, so editing its text is safe.
    DirectRun(u32),
    /// The `run_ord`-th `<w:r>` inside the `link_ord`-th `<w:hyperlink>` child of
    /// the paragraph. The `r:id` lives on the WRAPPER, so the run's own text and
    /// `<w:rPr>` are safely patchable through this path.
    Hyperlink { link_ord: u32, run_ord: u32 },
    /// The `run_ord`-th `<w:r>` inside the `field_ord`-th `<w:fldSimple>` child.
    /// The instruction lives on the wrapper's `w:instr` attribute.
    Field { field_ord: u32, run_ord: u32 },
}

/// A line break inside a paragraph (U+2028 LINE SEPARATOR): what a Word
/// text-wrapping `w:br` / `w:cr` is in [`Run::text`], and the engine's forced
/// line break in `insertText` text (the paragraph stays one paragraph).
pub const LINE_BREAK: char = '\u{2028}';

/// Word's optional hyphen (`<w:softHyphen/>`) in [`Run::text`]: U+00AD SOFT
/// HYPHEN, which the engine's composer takes as the only place its word may
/// break, drawing a hyphen there and nothing elsewhere.
pub const SOFT_HYPHEN: char = '\u{00AD}';

/// An absolute-position tab (`<w:ptab>`), which is a `\t` in [`Run::text`]
/// at char offset `at`: unlike a tab, it goes to a position of its own
/// rather than to the next tab stop.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PositionalTab {
    /// Char offset of its `\t` in the run's text.
    pub at: usize,
    /// `@w:alignment`: how the text after it sits at the position.
    pub alignment: PtabAlignment,
    /// `@w:relativeTo`: the position is the margins' or the indents'.
    pub relative_to: PtabBase,
    /// `@w:leader` as its display character (`None` for `none`).
    pub leader: Option<String>,
}

/// `w:ptab/@w:alignment`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PtabAlignment {
    /// At the left margin / indent.
    Left,
    /// Centred between the margins / indents.
    Center,
    /// At the right margin / indent.
    Right,
}

/// `w:ptab/@w:relativeTo`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PtabBase {
    Margin,
    Indent,
}

/// A symbol character (`<w:sym>`) in a run: its Unicode equivalent sits in
/// [`Run::text`] at char offset `at`, or, with no equivalent
/// ([`symbol_char`]), nothing does and `at` is where it was.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct RunSymbol {
    pub at: usize,
    /// `@w:font`.
    pub font: Option<String>,
    /// `@w:char`, as written (hex).
    pub code: String,
    /// The character carried in the text, if any.
    pub char: Option<char>,
}

/// A Word run (`w:r`): direct character formatting plus its text.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Run {
    /// `w:rPr/w:rStyle/@w:val` — the applied character style id, if any.
    pub style_id: Option<String>,
    /// Direct character formatting (`w:rPr`).
    pub props: RunProps,
    /// The concatenated text of the run's `w:t` children (tabs and
    /// absolute-position tabs as `\t`; text-wrapping breaks `w:br` / `w:cr`
    /// as [`LINE_BREAK`], U+2028, a line break inside the paragraph;
    /// `<w:noBreakHyphen/>` as U+2011, `<w:softHyphen/>` as [`SOFT_HYPHEN`],
    /// a `<w:sym>` as its Unicode equivalent where it has one). Page and
    /// column breaks are NOT in the text: they are [`Run::breaks`].
    pub text: String,
    /// `w:br w:type="page"|"column"` in this run, each at the char offset into
    /// [`Run::text`] where it sits (thoughts ADR 028/029: the content after
    /// one starts on a new page or column).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub breaks: Vec<RunBreak>,
    /// `<w:ptab>` in this run, each a `\t` in [`Run::text`].
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ptabs: Vec<PositionalTab>,
    /// `<w:sym>` in this run, carried or not ([`RunSymbol`]).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub symbols: Vec<RunSymbol>,
    /// The `w:drawing` pictures carried on this run, in order (`text` is
    /// empty for such a run).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub images: Vec<Image>,
    /// `w:drawing`s in this run that are not pictures (shapes, text boxes,
    /// charts, SmartArt) or whose picture cannot be resolved: not carried,
    /// counted so the lowering can say so.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub other_drawings: u32,
    /// When this run sits inside a `w:hyperlink`, its resolved target (an
    /// external URL, or `#anchor` for an internal bookmark). Styled blue +
    /// underline on lowering; the clickable link itself is preserved in the
    /// source `.docx` (a native clickable-hyperlink door is future work).
    pub hyperlink: Option<String>,
    /// Provenance for M2 save-back: which source `<w:r>` this run came from.
    /// `None` on runs not produced by the importer (e.g. test literals).
    #[serde(default)]
    pub source: Option<RunSource>,
    /// When this run carries a footnote/endnote reference (`w:footnoteReference`
    /// / `w:endnoteReference`), the referenced note's `w:id` — resolvable against
    /// [`DocxDocument::notes`].
    #[serde(default)]
    pub note_ref: Option<i64>,
    /// When this run is a FIELD's result, the field's name (the instruction's
    /// first token, upper-cased: `PAGE`, `DATE`, `NUMPAGES`, `REF`, …). The run's
    /// `text` is the value Word last computed — a frozen snapshot, since the
    /// native model has no equivalent for most field kinds. `HYPERLINK` fields
    /// are handled separately (see [`Run::hyperlink`]).
    #[serde(default)]
    pub field: Option<String>,
}

/// A page or column break inside a run (`w:br` with `w:type` `page` or
/// `column`; a typeless or `textWrapping` break is a line break, `\n` in the
/// text).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct RunBreak {
    /// Char (Unicode scalar) offset into the run's text: the break sits
    /// before the `at`-th char (`at == text.chars().count()` = after it all).
    pub at: usize,
    pub kind: BreakKind,
}

/// What a [`RunBreak`] starts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BreakKind {
    /// `w:br w:type="page"`: the next page.
    Page,
    /// `w:br w:type="column"`: the next column (the next page when the
    /// section has one column).
    Column,
}

/// A header or footer part's content (`w:hdr` / `w:ftr`), reached from a
/// section's `headerReference`/`footerReference`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HeaderFooter {
    /// `true` for a footer, `false` for a header.
    pub footer: bool,
    /// `w:type` on the first reference to it: `default` / `first` / `even`.
    pub kind: Option<String>,
    /// The part name (`word/header1.xml`), for save-back provenance.
    #[serde(default)]
    pub part: String,
    /// The section that first references it (0-based).
    #[serde(default)]
    pub section: usize,
    pub paragraphs: Vec<Paragraph>,
}

/// Which header (or footer) part a section shows, per `w:type`, as indices
/// into [`DocxDocument::headers_footers`] (thoughts ADR 033). `None` is a
/// blank one: no reference of that kind in this section or any before it.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct HeaderFooterSet {
    pub default: Option<HeaderFooterRef>,
    pub first: Option<HeaderFooterRef>,
    pub even: Option<HeaderFooterRef>,
}

/// One section's header or footer of one kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct HeaderFooterRef {
    /// Index into [`DocxDocument::headers_footers`].
    pub index: usize,
    /// `true` when the section has no reference of this kind and shows the
    /// previous section's (Word's rule, ECMA-376 §17.10.5).
    pub inherited: bool,
}

/// The kind of header/footer a page shows (`w:type`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum HeaderFooterKind {
    Default,
    First,
    Even,
}

impl HeaderFooterSet {
    pub fn get(&self, kind: HeaderFooterKind) -> Option<HeaderFooterRef> {
        match kind {
            HeaderFooterKind::Default => self.default,
            HeaderFooterKind::First => self.first,
            HeaderFooterKind::Even => self.even,
        }
    }

    pub fn is_empty(&self) -> bool {
        self.default.is_none() && self.first.is_none() && self.even.is_none()
    }
}

/// Footnote or endnote numbering (`w:footnotePr` / `w:endnotePr`, thoughts
/// ADR 034). `None` is "not said here" (inherit, or Word's default).
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct NoteProps {
    /// `w:numFmt/@w:val` as Word writes it (`decimal`, `lowerRoman`,
    /// `upperLetter`, `chicago`, …).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    /// `w:numStart/@w:val`: the first note's number.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<u32>,
    /// `w:numRestart/@w:val`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<NoteRestart>,
    /// `w:pos/@w:val` as Word writes it: footnotes `pageBottom` /
    /// `beneathText` / `sectEnd`; endnotes `sectEnd` / `docEnd`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<String>,
}

/// `w:numRestart/@w:val`: when note numbering starts over.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum NoteRestart {
    /// Never (Word's default).
    Continuous,
    /// At each section.
    EachSection,
    /// On each page (footnotes only).
    EachPage,
}

impl NoteProps {
    pub fn is_empty(&self) -> bool {
        *self == NoteProps::default()
    }

    /// The number Word gives a footnote under these (a section's own)
    /// numbering settings, from its 0-based ordinals among ALL the
    /// document's footnotes, its section's, and its page's (measured,
    /// `fixtures/footnote-numbering.word.json`): `numStart` (else 1) plus
    /// the ordinal since the restart point — the page for `eachPage`, the
    /// section for `eachSect`, and the DOCUMENT's start for `continuous`,
    /// whatever earlier sections restarted or started at.
    pub fn note_number(&self, in_document: u32, in_section: u32, on_page: u32) -> u32 {
        let start = self.num_start.unwrap_or(1);
        start
            + match self.num_restart.unwrap_or(NoteRestart::Continuous) {
                NoteRestart::Continuous => in_document,
                NoteRestart::EachSection => in_section,
                NoteRestart::EachPage => on_page,
            }
    }
}

/// A note number in Word's `w:numFmt` (`decimal` when `None`), for the
/// formats a note mark commonly takes: `decimal`, `lowerRoman`,
/// `upperRoman`, `lowerLetter`, `upperLetter` (Word's letters repeat:
/// 27 is `AA`). `None` for any other format.
pub fn format_note_number(n: u32, fmt: Option<&str>) -> Option<String> {
    fn roman(mut n: u32) -> String {
        const T: &[(u32, &str)] = &[
            (1000, "m"),
            (900, "cm"),
            (500, "d"),
            (400, "cd"),
            (100, "c"),
            (90, "xc"),
            (50, "l"),
            (40, "xl"),
            (10, "x"),
            (9, "ix"),
            (5, "v"),
            (4, "iv"),
            (1, "i"),
        ];
        let mut s = String::new();
        for (v, r) in T {
            while n >= *v {
                s.push_str(r);
                n -= v;
            }
        }
        s
    }
    fn letter(n: u32) -> String {
        if n == 0 {
            return String::new();
        }
        let c = char::from(b'a' + ((n - 1) % 26) as u8);
        std::iter::repeat_n(c, ((n - 1) / 26 + 1) as usize).collect()
    }
    Some(match fmt.unwrap_or("decimal") {
        "decimal" => n.to_string(),
        "lowerRoman" => roman(n),
        "upperRoman" => roman(n).to_uppercase(),
        "lowerLetter" => letter(n),
        "upperLetter" => letter(n).to_uppercase(),
        _ => return None,
    })
}

impl DocxDocument {
    /// The header (`footer == false`) or footer a page of section `section`
    /// shows, by Word's rule (ECMA-376 §17.10.1-6): the section's first page
    /// takes `first` when the section has `w:titlePg`; otherwise an even
    /// page takes `even` when the document has `w:evenAndOddHeaders`;
    /// otherwise `default`. `page_number` is the page's NUMBER (after any
    /// `w:pgNumType/@w:start`), which is what decides even and odd.
    /// `None` is a blank header.
    pub fn header_footer_for(
        &self,
        section: usize,
        footer: bool,
        first_of_section: bool,
        page_number: i32,
    ) -> Option<&HeaderFooter> {
        let s = self.sections.get(section)?;
        let set = if footer { &s.footers } else { &s.headers };
        let kind = if first_of_section && s.title_page {
            HeaderFooterKind::First
        } else if self.even_and_odd_headers && page_number % 2 == 0 {
            HeaderFooterKind::Even
        } else {
            HeaderFooterKind::Default
        };
        set.get(kind)
            .and_then(|r| self.headers_footers.get(r.index))
    }

    /// The footnote numbering in force in section `section`: its OWN
    /// `w:footnotePr`, alone. Word takes nothing from `settings.xml` nor
    /// from the previous section (measured,
    /// `fixtures/footnote-numbering.word.json`); what it leaves out is
    /// Word's default (decimal, from 1, continuous).
    pub fn footnote_numbering(&self, section: usize) -> NoteProps {
        self.sections
            .get(section)
            .map_or_else(NoteProps::default, |s| s.footnote_props.clone())
    }

    /// The endnote numbering in force in section `section`: its own
    /// `w:endnotePr`, by analogy with footnotes (endnotes not measured).
    pub fn endnote_numbering(&self, section: usize) -> NoteProps {
        self.sections
            .get(section)
            .map_or_else(NoteProps::default, |s| s.endnote_props.clone())
    }
}

/// An image (`w:drawing` → `wp:inline`/`wp:anchor` → a picture blip),
/// resolved to its media bytes + intrinsic size at import time. A floating
/// one (`wp:anchor`) carries its position and wrap in [`Image::float`].
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Image {
    /// The raw media bytes (PNG/JPEG/…) from `word/media/…`.
    pub bytes: Vec<u8>,
    /// The image MIME type (from the media part extension).
    pub mime: String,
    /// Intrinsic width in EMU (`wp:extent/@cx`; 914400 EMU/inch, 12700 EMU/pt).
    pub width_emu: i64,
    /// Intrinsic height in EMU (`wp:extent/@cy`).
    pub height_emu: i64,
    /// `Some` for a floating drawing (`wp:anchor`): where Word positions it
    /// and how text wraps around it (thoughts ADR 035). `None` for an inline
    /// one (`wp:inline`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub float: Option<Float>,
}

/// A floating drawing's placement (`wp:anchor`, thoughts ADR 035). Lengths in
/// EMU, as Word writes them.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct Float {
    /// `wp:positionH`.
    pub horizontal: Option<FloatPosition>,
    /// `wp:positionV`.
    pub vertical: Option<FloatPosition>,
    /// The wrap element (`wp:wrapSquare` / … / `wp:wrapNone`).
    pub wrap: FloatWrap,
    /// `@wrapText` of a square/tight/through wrap: `bothSides` / `left` /
    /// `right` / `largest`.
    pub wrap_text: Option<String>,
    /// `@distT` / `@distB` / `@distL` / `@distR` on `wp:anchor` (the wrap
    /// element's own, where it has them, win).
    pub dist_top: i64,
    pub dist_bottom: i64,
    pub dist_left: i64,
    pub dist_right: i64,
    /// `@behindDoc`: drawn behind the text.
    pub behind_doc: bool,
    /// `@allowOverlap`.
    pub allow_overlap: bool,
    /// `@layoutInCell`.
    pub layout_in_cell: bool,
    /// `@locked`.
    pub locked: bool,
    /// `@relativeHeight`: z-order among floats.
    pub relative_height: u32,
    /// `@simplePos="1"` with `wp:simplePos` (x, y) from the page's
    /// top-left: the position elements are ignored.
    pub simple_pos: Option<(i64, i64)>,
}

/// One axis of a float's position (`wp:positionH` / `wp:positionV`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FloatPosition {
    /// `@relativeFrom` as Word writes it (`page`, `margin`, `column`,
    /// `character`, `paragraph`, `line`, `leftMargin`, …).
    pub relative_from: String,
    /// `wp:posOffset` in EMU.
    pub offset: Option<i64>,
    /// `wp:align` (`left` / `center` / `right` / `inside` / `outside`, or
    /// `top` / `bottom` / …).
    pub align: Option<String>,
    /// `wp14:pctPosHOffset` / `wp14:pctPosVOffset`, in thousandths of a
    /// percent, as written.
    pub percent: Option<i64>,
}

/// How text wraps around a float.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum FloatWrap {
    /// `wp:wrapNone` (or no wrap element): in front of / behind the text.
    #[default]
    None,
    Square,
    Tight,
    Through,
    TopAndBottom,
}

impl FloatWrap {
    /// Word's own name for the wrap element.
    pub fn as_word(self) -> &'static str {
        match self {
            FloatWrap::None => "wrapNone",
            FloatWrap::Square => "wrapSquare",
            FloatWrap::Tight => "wrapTight",
            FloatWrap::Through => "wrapThrough",
            FloatWrap::TopAndBottom => "wrapTopAndBottom",
        }
    }
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
    /// `w:spacing/@w:line` + `@w:lineRule` (ADR 029). For `exact` and
    /// `atLeast` the value is twips; for `auto` it is 240ths of a line.
    #[serde(default)]
    pub line_spacing: Option<LineSpacing>,
    pub keep_next: Option<bool>,
    /// `w:keepLines`: every line of the paragraph on one page.
    pub keep_lines: Option<bool>,
    /// `w:widowControl`: no single first or last line of the paragraph alone
    /// on a page (Word moves a second line with it). Absent everywhere in
    /// the style hierarchy, it is off.
    #[serde(default)]
    pub widow_control: Option<bool>,
    /// `w:pageBreakBefore` (ADR 028/029): the paragraph starts a new page.
    /// `Some(false)` is an explicit `w:val="0"`, which turns off a style's.
    #[serde(default)]
    pub page_break_before: Option<bool>,
    /// `w:tabs` — explicit tab stops (empty = inherit).
    pub tabs: Vec<TabStop>,
}

/// A tab stop (`w:tab`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TabStop {
    /// `@w:pos` in twips.
    pub position: i32,
    /// `@w:val` alignment (`"left"`, `"center"`, `"right"`, `"decimal"`, …).
    /// `None` for a `"clear"` stop, which removes the stop an inherited style
    /// sets at the same position (Word MERGES a paragraph's stops with its
    /// style chain's).
    pub alignment: Option<String>,
    /// `@w:leader` as the native leader string (`"."` for `dot`, `"-"` for
    /// `hyphen`, `"_"` for `underscore`/`heavy`, `"\u{B7}"` for `middleDot`).
    pub leader: Option<String>,
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
    /// `w:position/@w:val` — baseline shift in half-points (signed; positive
    /// raises).
    pub baseline_half_pts: Option<i32>,
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
    /// `@w:default="1"`: the default style of its type. A paragraph that
    /// names no style is in the default paragraph style (Word's Normal).
    #[serde(default)]
    pub is_default: bool,
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

/// A table (`w:tbl`): a column grid + rows of cells.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Table {
    /// `w:tblGrid/w:gridCol/@w:w` — column widths in twips (defines column count).
    pub column_widths: Vec<i32>,
    pub rows: Vec<TableRow>,
}

/// A table row (`w:tr`).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

/// A table cell (`w:tc`) — block content (paragraphs) plus merge spans.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<Paragraph>,
    /// `w:tcPr/w:gridSpan/@w:val` — horizontal span (default 1).
    pub grid_span: u32,
    /// `w:tcPr/w:vMerge` — vertical merge role.
    pub v_merge: VMerge,
}

/// A cell's vertical-merge role (`w:vMerge`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum VMerge {
    /// Not vertically merged.
    #[default]
    None,
    /// `w:vMerge w:val="restart"` — the top cell of a vertical span.
    Restart,
    /// `w:vMerge` (or `val="continue"`) — a continuation cell absorbed by the
    /// restart cell above it.
    Continue,
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
    /// `w:cols/@w:space` — the gap between equal-width columns, in twips
    /// (Word's default 720).
    #[serde(default = "default_column_space")]
    pub column_space: i32,
    /// Unequal columns (`w:cols/@w:equalWidth="0"`): each `w:col` as
    /// `(w, space)` in twips. Empty for equal-width columns.
    #[serde(default)]
    pub column_widths: Vec<(i32, i32)>,
    /// Index into [`DocxDocument::body`] of this section's FIRST block. A
    /// section runs to the next section's `first_block` (the last one to the
    /// end of the body). Word marks a section's end with a `sectPr` inside its
    /// last paragraph's `pPr`; the body-level `sectPr` is the final section.
    #[serde(default)]
    pub first_block: usize,
    /// `w:type/@w:val` — how the section starts (thoughts ADR 029).
    #[serde(default)]
    pub kind: SectionKind,
    /// The headers this section shows, after Word's inheritance
    /// (thoughts ADR 033).
    #[serde(default)]
    pub headers: HeaderFooterSet,
    /// The footers this section shows, after Word's inheritance.
    #[serde(default)]
    pub footers: HeaderFooterSet,
    /// `w:titlePg`: the section's first page shows the `first` header and
    /// footer. Per section, not inherited.
    #[serde(default)]
    pub title_page: bool,
    /// `w:pgMar/@w:header`: the header's distance from the page top, twips.
    #[serde(default)]
    pub header_distance: Option<i32>,
    /// `w:pgMar/@w:footer`: the footer's distance from the page bottom.
    #[serde(default)]
    pub footer_distance: Option<i32>,
    /// `w:pgNumType/@w:start`: the section's page numbering starts over.
    #[serde(default)]
    pub page_number_start: Option<i32>,
    /// `w:pgNumType/@w:fmt` as Word writes it (`decimal`, `lowerRoman`, …).
    #[serde(default)]
    pub page_number_format: Option<String>,
    /// The section's own `w:footnotePr` (thoughts ADR 034).
    #[serde(default)]
    pub footnote_props: NoteProps,
    /// The section's own `w:endnotePr`.
    #[serde(default)]
    pub endnote_props: NoteProps,
}

/// Word line spacing (`w:spacing/@w:line` with its `@w:lineRule`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct LineSpacing {
    pub value: i32,
    pub rule: LineRule,
}

/// `w:spacing/@w:lineRule`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum LineRule {
    /// A multiple of single spacing (`value` in 240ths of a line).
    #[default]
    Auto,
    /// Exactly `value` twips.
    Exact,
    /// At least `value` twips.
    AtLeast,
}

/// `w:sectPr/w:type` — where a section starts.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum SectionKind {
    /// On a new page (Word's default).
    #[default]
    NextPage,
    /// On the same page, after the previous section's text.
    Continuous,
    /// On the next even page.
    EvenPage,
    /// On the next odd page.
    OddPage,
    /// In the next column.
    NextColumn,
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
            column_space: default_column_space(),
            column_widths: Vec::new(),
            first_block: 0,
            kind: SectionKind::NextPage,
            headers: HeaderFooterSet::default(),
            footers: HeaderFooterSet::default(),
            title_page: false,
            header_distance: None,
            footer_distance: None,
            page_number_start: None,
            page_number_format: None,
            footnote_props: NoteProps::default(),
            endnote_props: NoteProps::default(),
        }
    }
}

fn is_zero(n: &u32) -> bool {
    *n == 0
}

/// Word's gap between columns when `w:cols/@w:space` is absent (720 twips,
/// 0.5 in).
pub fn default_column_space() -> i32 {
    720
}
