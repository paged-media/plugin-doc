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

//! ADR 029 — which Word sections share a native story.
//!
//! Standalone open gives every story its own page, margin-box frame and grow
//! rule (`docx-skeleton`). A section that Word starts on a NEW page is a new
//! story. A section Word continues on the SAME page (`continuous`, or
//! `nextColumn`) joins the story before it when the native model can say
//! what Word does there. Word's answers, measured on
//! `docx_conformance::continuous_docx()` (`fixtures/continuous.word.json`):
//!
//! - same page size, margins and columns: the boundary is invisible; the
//!   section's lines follow the previous section's on the same page (and run
//!   on to the next page). → JOIN.
//! - other LEFT/RIGHT margins: applied at once, mid-page, and on every page
//!   after. → JOIN, with the margin difference as paragraph indents (exact in
//!   one column while the new text area lies inside the story's frame).
//! - other TOP/BOTTOM margins: the section continues on the same page under
//!   the old margins; the new ones apply from the next page. → JOIN (the page
//!   map is right where the section starts), with a diagnostic: the native
//!   story's later pages keep the earlier section's top/bottom margins.
//! - another PAGE SIZE: Word starts a new page. → a new story (no loss).
//! - other COLUMNS: Word changes the column count mid-page and balances the
//!   section's columns before a continuous break. The engine has no mid-page
//!   column change (no IDML `SpanColumnType` split/span columns in the model
//!   or on the wire) → a new story on a new page, with a warning.
//! - `nextColumn`, same geometry: the section opens the next column of the
//!   same page. → JOIN with a `NextColumn` break-before on its first
//!   paragraph.

use docx_core::{Section, SectionKind};

use crate::ir::Diagnostic;

/// Where one Word section goes in the native skeleton.
#[derive(Debug, Clone, PartialEq)]
pub struct SectionPlacement {
    /// Index of the native story (and skeleton page) the section pours into.
    /// Sections sharing a story are consecutive.
    pub story: usize,
    /// Extra left/right paragraph indent (pt) that puts the section's text
    /// between its own margins inside the story's frame (0 when the margins
    /// match the story's first section).
    pub indent_left_pt: f32,
    pub indent_right_pt: f32,
    /// The section joins its story at the next column (`nextColumn`).
    pub next_column: bool,
}

fn twips(v: i32) -> f32 {
    v as f32 / 20.0
}

/// Place every section (see the module docs for the rules and Word's
/// measurements); diagnostics for what the native model cannot say are
/// pushed to `diagnostics`.
pub fn place_sections(
    sections: &[Section],
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<SectionPlacement> {
    let mut out: Vec<SectionPlacement> = Vec::with_capacity(sections.len());
    // The first section of the current story: its frame is the story's.
    let mut base: Option<&Section> = None;
    for (k, sec) in sections.iter().enumerate() {
        let joins = matches!(sec.kind, SectionKind::Continuous | SectionKind::NextColumn);
        let story = out.last().map(|p| p.story).unwrap_or(0);
        let joined = match base.filter(|_| joins) {
            Some(b) => join(k, b, sec, diagnostics),
            None => None,
        };
        let placement = match joined {
            Some(mut p) => {
                p.story = story;
                p
            }
            None => {
                base = Some(sec);
                SectionPlacement {
                    story: if k == 0 { 0 } else { story + 1 },
                    indent_left_pt: 0.0,
                    indent_right_pt: 0.0,
                    next_column: false,
                }
            }
        };
        out.push(placement);
    }
    out
}

/// `sec` (index `k`, continuous or nextColumn) joining the story whose frame
/// is `base`'s, or `None` when it must open a new story.
fn join(
    k: usize,
    base: &Section,
    sec: &Section,
    diagnostics: &mut Vec<Diagnostic>,
) -> Option<SectionPlacement> {
    let kind = if sec.kind == SectionKind::NextColumn {
        "nextColumn"
    } else {
        "continuous"
    };
    if (sec.page_width, sec.page_height) != (base.page_width, base.page_height) {
        // Measured: Word starts a new page for another page size.
        return None;
    }
    let (base_cols, cols) = (base.columns.max(1), sec.columns.max(1));
    if cols != base_cols {
        diagnostics.push(Diagnostic::warning(
            format!(
                "section {} is {kind} and changes from {base_cols} to {cols} column(s): \
                 Word continues it on the same page; the engine has no mid-page column \
                 change (IDML split/span columns are not modelled), so it opens a new page",
                k + 1
            ),
            3,
        ));
        return None;
    }
    let left = twips(sec.margin_left - base.margin_left);
    let right = twips(sec.margin_right - base.margin_right);
    if left != 0.0 || right != 0.0 {
        if cols > 1 {
            diagnostics.push(Diagnostic::warning(
                format!(
                    "section {} is {kind} with other left/right margins in {cols} columns: \
                     Word narrows the columns mid-page; paragraph indents cannot narrow \
                     columns, so it opens a new page",
                    k + 1
                ),
                3,
            ));
            return None;
        }
        if left < 0.0 || right < 0.0 {
            diagnostics.push(Diagnostic::warning(
                format!(
                    "section {} is {kind} with narrower left/right margins than the page \
                     it continues: Word widens the text mid-page; the text cannot leave its \
                     frame, so it opens a new page",
                    k + 1
                ),
                3,
            ));
            return None;
        }
    }
    if (sec.margin_top, sec.margin_bottom) != (base.margin_top, base.margin_bottom) {
        diagnostics.push(Diagnostic::info(
            format!(
                "section {} is {kind} with other top/bottom margins: it continues on the \
                 same page like Word's, but its later pages keep the earlier section's \
                 top/bottom margins (Word applies the new ones from the next page)",
                k + 1
            ),
            3,
        ));
    }
    Some(SectionPlacement {
        story: 0, // the caller sets the story it joins
        indent_left_pt: left,
        indent_right_pt: right,
        next_column: sec.kind == SectionKind::NextColumn,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sec(kind: SectionKind, columns: u32, left: i32, right: i32) -> Section {
        Section {
            kind,
            columns,
            margin_left: left,
            margin_right: right,
            ..Section::default()
        }
    }

    fn stories(sections: &[Section]) -> (Vec<usize>, Vec<Diagnostic>) {
        let mut diags = Vec::new();
        let placed = place_sections(sections, &mut diags);
        (placed.iter().map(|p| p.story).collect(), diags)
    }

    #[test]
    fn a_first_continuous_section_still_opens_the_first_story() {
        let (s, d) = stories(&[sec(SectionKind::Continuous, 1, 1440, 1440)]);
        assert_eq!((s, d.len()), (vec![0], 0));
    }

    #[test]
    fn narrower_margins_cannot_join_and_say_so() {
        use SectionKind::*;
        let (s, d) = stories(&[sec(NextPage, 1, 1440, 1440), sec(Continuous, 1, 720, 1440)]);
        assert_eq!(s, vec![0, 1]);
        assert_eq!(d.len(), 1);
        assert_eq!(d[0].severity, "warning");
    }

    #[test]
    fn other_margins_in_columns_cannot_join_and_say_so() {
        use SectionKind::*;
        let (s, d) = stories(&[sec(NextPage, 2, 1440, 1440), sec(Continuous, 2, 2160, 1440)]);
        assert_eq!(s, vec![0, 1]);
        assert!(d[0].message.contains("2 columns"), "{}", d[0].message);
    }

    #[test]
    fn indents_measure_from_the_storys_first_section() {
        use SectionKind::*;
        let mut diags = Vec::new();
        let placed = place_sections(
            &[
                sec(NextPage, 1, 1440, 1440),
                sec(Continuous, 1, 2160, 1440),
                sec(Continuous, 1, 2880, 1800),
            ],
            &mut diags,
        );
        assert_eq!(
            placed
                .iter()
                .map(|p| (p.story, p.indent_left_pt, p.indent_right_pt))
                .collect::<Vec<_>>(),
            vec![(0, 0.0, 0.0), (0, 36.0, 0.0), (0, 72.0, 18.0)]
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn page_starting_kinds_never_join() {
        use SectionKind::*;
        let (s, _) = stories(&[
            sec(NextPage, 1, 1440, 1440),
            sec(NextPage, 1, 1440, 1440),
            sec(OddPage, 1, 1440, 1440),
            sec(EvenPage, 1, 1440, 1440),
            sec(NextColumn, 1, 1440, 1440),
        ]);
        assert_eq!(s, vec![0, 1, 2, 3, 3]);
    }
}
