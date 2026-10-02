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

//! Word's line height (ADR 029), as Word itself measured it.
//!
//! `w:spacing/@w:line` is relative to the FONT's single line, not to the
//! point size, so lowering it to the engine's `characterLeading` (a line
//! pitch in points) needs that single line. Word was asked
//! (`scripts/word-line-spacing-probe.sh` on `docx_conformance::
//! line_spacing_docx()`, recorded in `docx-conformance/fixtures/
//! line-spacing.word.json`), and its numbers say:
//!
//! - **single** = `(hhea.ascender − hhea.descender + hhea.lineGap) / unitsPerEm
//!   × size` of the face Word loads. Every one of 25 measured faces fits this
//!   to ±0.01 pt at 10 pt. The OS/2 win metrics do NOT (Aptos: win says
//!   12.85 pt, Word laid 12.21; Corbel: win 12.21, Word 12.08), nor do the
//!   typo metrics (Arial: typo 10.88, Word 11.50).
//! - **auto** `n` = single × n / 240 (Calibri 10 pt: 240 → 12.21, 276 →
//!   14.04, 360 → 18.32, 480 → 24.42).
//! - **exact** `n` = n / 20 pt, whatever the face.
//! - **atLeast** `n` = max(n / 20, single) (Calibri 10 pt atLeast 6 pt and
//!   12 pt both lay 12.21; Arial 10 pt atLeast 12 pt lays 12.00).
//!
//! WHY a table of measured faces and not the font file: this lowering is
//! pure (no font bytes reach it), and the font file on disk is not the face
//! Word uses — Word's bundled `times.ttf` has hhea lineGap 0 (1.107 em) yet
//! Word laid Times New Roman at 1.150 em, from another copy; its Century
//! Gothic and Book Antiqua likewise differ from the same-named files in its
//! own bundle. The table therefore holds the hhea ratio of the face Word
//! EMBEDDED in its PDF for each measured case, each confirmed by the
//! measured pitch.
//!
//! An UNLISTED face (or none: a theme font such as `minorHAnsi`) falls back
//! to 2500/2048 em: it is the line of both default theme body faces (Calibri
//! and Aptos, measured), and it is what Word laid for Inter, which it could
//! not load and silently replaced by Calibri. A face Word CAN load but that
//! is not listed will be off by its own hhea ratio's distance from that
//! (typically under 0.1 em); the lowering names such faces in a diagnostic.

use docx_core::{LineRule, LineSpacing};

/// Single line height per em of the default theme body faces (Calibri,
/// Aptos), and of the face Word substitutes for one it cannot load.
pub const FALLBACK_EM: f32 = 2500.0 / 2048.0;

/// `(family, hhea ascender − descender + lineGap in 2048ths of an em)` for
/// every face Word was measured on (all are 2048-unit faces).
const MEASURED: &[(&str, u16)] = &[
    ("Aptos", 2500),
    ("Arial", 2355),
    ("Arial Narrow", 2319),
    ("Book Antiqua", 2545),
    ("Calibri", 2500),
    ("Calibri Light", 2500),
    ("Cambria", 2401),
    ("Candara", 2500),
    ("Comic Sans MS", 2854),
    ("Consolas", 2398),
    ("Constantia", 2500),
    ("Corbel", 2473),
    ("Century Gothic", 2440),
    ("Courier New", 2320),
    ("Franklin Gothic Book", 2322),
    ("Garamond", 2304),
    ("Georgia", 2327),
    ("Gill Sans MT", 2375),
    ("Helvetica", 2355),
    ("Lucida Sans Unicode", 3147),
    ("Palatino Linotype", 2763),
    ("Symbol", 2510),
    ("Tahoma", 2472),
    ("Times New Roman", 2355),
    ("Trebuchet MS", 2378),
    ("Verdana", 2489),
    ("Wingdings", 2273),
];

/// Word's single line height for `font` in ems, or `None` when Word has not
/// been measured on that face (callers then use [`FALLBACK_EM`]).
pub fn measured_single_em(font: &str) -> Option<f32> {
    MEASURED
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(font.trim()))
        .map(|(_, units)| f32::from(*units) / 2048.0)
}

/// hhea ascenders (1/2048 em) of the measured faces, read from the font
/// files Word ships (fontTools, 2026-10-02). The native table puts a cell's
/// first baseline one ascent below its top inset.
const ASCENT: &[(&str, u16)] = &[
    ("Aptos", 1923),
    ("Arial", 1854),
    ("Arial Narrow", 1916),
    ("Book Antiqua", 1891),
    ("Calibri", 1950),
    ("Calibri Light", 1950),
    ("Cambria", 1946),
    ("Candara", 1484),
    ("Comic Sans MS", 2257),
    ("Consolas", 1521),
    ("Constantia", 1538),
    ("Corbel", 1523),
    ("Century Gothic", 2060),
    ("Courier New", 1705),
    ("Franklin Gothic Book", 1877),
    ("Garamond", 1765),
    ("Georgia", 1878),
    ("Gill Sans MT", 1903),
    ("Helvetica", 1577),
    ("Lucida Sans Unicode", 2246),
    ("Palatino Linotype", 2150),
    ("Tahoma", 2049),
    ("Times New Roman", 1825),
    ("Trebuchet MS", 1923),
    ("Verdana", 2059),
];

/// The ascent of `font` in ems (Calibri's when the face is unknown or the
/// theme's, like [`FALLBACK_EM`]).
pub fn ascent_em(font: Option<&str>) -> f32 {
    let units = font
        .and_then(|f| {
            ASCENT
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case(f.trim()))
        })
        .map_or(1950, |(_, u)| *u);
    f32::from(units) / 2048.0
}

/// Word's single line height in points for a run in `font` (`None` = theme
/// or default face) at `size_pt`.
pub fn single_line_pt(font: Option<&str>, size_pt: f32) -> f32 {
    font.and_then(measured_single_em).unwrap_or(FALLBACK_EM) * size_pt
}

/// The line pitch Word lays for `spacing` when the line's tallest run has a
/// single line of `single_pt`.
pub fn line_pitch_pt(spacing: LineSpacing, single_pt: f32) -> f32 {
    let twips = spacing.value as f32 / 20.0;
    match spacing.rule {
        LineRule::Exact => twips,
        LineRule::AtLeast => twips.max(single_pt),
        LineRule::Auto => single_pt * spacing.value as f32 / 240.0,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ls(value: i32, rule: LineRule) -> LineSpacing {
        LineSpacing { value, rule }
    }

    #[test]
    fn auto_is_a_multiple_of_the_faces_single_line() {
        let single = single_line_pt(Some("Calibri"), 10.0);
        assert!((single - 12.207).abs() < 0.001);
        assert!((line_pitch_pt(ls(480, LineRule::Auto), single) - 24.414).abs() < 0.001);
    }

    #[test]
    fn at_least_never_goes_below_single() {
        let single = single_line_pt(Some("arial"), 10.0);
        assert!((line_pitch_pt(ls(120, LineRule::AtLeast), single) - single).abs() < 1e-4);
        assert_eq!(line_pitch_pt(ls(240, LineRule::AtLeast), single), 12.0);
    }

    #[test]
    fn an_unmeasured_face_uses_the_theme_body_line() {
        assert_eq!(measured_single_em("Inter"), None);
        assert_eq!(single_line_pt(Some("Inter"), 10.0), FALLBACK_EM * 10.0);
        assert_eq!(single_line_pt(None, 10.0), FALLBACK_EM * 10.0);
    }
}
