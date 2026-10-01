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
//! `docx_conformance::continuous_docx()` (`fixtures/continuous.word.json`)
//! and `columns_docx()` (`fixtures/columns.word.json`):
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
//! - other COLUMNS: Word changes the column count (or gap) mid-page. A
//!   section before a `continuous` break is BALANCED by line count
//!   (`ceil(lines / k)` per column: 9 → 5 / 4, 7 in 3 → 3 / 3 / 1); one
//!   before a page-starting section, or at the document's end, is NOT (its
//!   lines fill column 1 first). A section that runs past the page fills
//!   that page's columns to the bottom and is balanced on its last page.
//!   → JOIN, with the engine's span/split columns (IDML `SpanColumnType`,
//!   protocol 64; core lays them out as InDesign does, ADR 028 addendum),
//!   decided per story ([`StoryColumns`]):
//!   - the story's LAST section is multi-column and Word leaves it
//!     unbalanced (a page-starting section or the document's end follows),
//!     and it is the story's only multi-column layout: the FRAME gets those
//!     columns and the one-column sections SPAN all of them. The engine
//!     balances the text above a span by line count and fills the columns
//!     after the last span in turn — Word's rule in both cases (B, I, M, O).
//!     The text above a span is balanced only when it opens its band, so a
//!     balanced section that runs in from the previous page is not;
//!   - every other story keeps one column and each multi-column section
//!     SPLITS it (`(column − (k − 1) × gap) / k` sub-columns, inside gutter
//!     = `w:space`, outside 0). A split block is balanced, also after it
//!     filled a page (J, L, and the first part of H, N, P). A split section
//!     that ends its story where Word leaves it unbalanced (two layouts in
//!     the story) is balanced on its last page (diagnosed). Two
//!     multi-column sections in a row with another count or gap are two
//!     split blocks: the engine ends a block where the count, the inside or
//!     the outside gutter changes, balances each on its own and starts the
//!     next directly below the deepest sub-column (core `d4311c7`, ADR 028
//!     addendum) — Word's map (H3, N3, P3).
//!     Two multi-column sections in a row with the SAME count and gap: Word
//!     balances each on its own (Q; R before a page-starting section). The
//!     engine has no boundary there (no setting changes, and its frame
//!     columns have none at all), so they share one column flow on the same
//!     page, with a warning. No honest native construct separates them: a
//!     gutter difference would move the text, and a paragraph between them
//!     would be content Word does not have.
//!
//!   An engine that refuses those properties (before protocol 64) gets the
//!   page-break lowering instead ([`LowerOptions::mid_page_columns`] off):
//!   every column change opens a new page, with a warning. The bundle
//!   learns which by the refusal of the style batch (`doc-bundle` open).
//! - unequal columns (`w:equalWidth="0"`): neither frame nor split columns
//!   can be unequal; they are laid out as equal columns with the first
//!   column's gap (Word: 180 pt + 90 pt, balanced 4 / 2), with a warning.
//! - `nextColumn`, same geometry: the section opens the next column of the
//!   same page. → JOIN with a `NextColumn` break-before on its first
//!   paragraph. With other columns (not measured) it opens a new page.

use docx_core::{Section, SectionKind};

use crate::ir::Diagnostic;

/// What the lowering may use (see [`LowerOptions::mid_page_columns`]).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LowerOptions {
    /// Lower a mid-page column change to span/split columns (protocol 64).
    /// Off, every column change opens a new page (the lowering for engines
    /// that refuse the properties).
    pub mid_page_columns: bool,
}

impl Default for LowerOptions {
    fn default() -> Self {
        LowerOptions {
            mid_page_columns: true,
        }
    }
}

/// How a section's paragraphs sit in its story's frame.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub enum SectionColumns {
    /// In the frame's own columns.
    #[default]
    Frame,
    /// Spanning all of the frame's columns (a one-column section in a
    /// multi-column story).
    SpanAll,
    /// Splitting the frame's one column into `count` sub-columns
    /// `inside_pt` apart.
    Split { count: u32, inside_pt: f32 },
}

/// The column layout of one native story's frame.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct StoryColumns {
    pub count: u32,
    pub gutter_pt: f32,
}

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
    /// How the section's paragraphs sit in the story's columns.
    pub columns: SectionColumns,
    /// The story frame's columns (the same for every section of a story).
    pub frame: StoryColumns,
}

fn twips(v: i32) -> f32 {
    v as f32 / 20.0
}

/// A section's column layout as the engine can lay it out: the count and
/// the gap in twips (unequal columns: the first column's gap).
fn layout(sec: &Section) -> (u32, i32) {
    let count = sec.columns.max(1);
    let gap = match sec.column_widths.first() {
        Some(&(_, space)) if count > 1 => space,
        _ => sec.column_space,
    };
    (count, if count > 1 { gap } else { 0 })
}

/// Place every section with the default [`LowerOptions`].
pub fn place_sections(
    sections: &[Section],
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<SectionPlacement> {
    place_sections_with(sections, LowerOptions::default(), diagnostics)
}

/// Place every section (see the module docs for the rules and Word's
/// measurements); diagnostics for what the native model cannot say are
/// pushed to `diagnostics`.
pub fn place_sections_with(
    sections: &[Section],
    options: LowerOptions,
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<SectionPlacement> {
    let mut out: Vec<SectionPlacement> = Vec::with_capacity(sections.len());
    // The first section of the current story: its frame is the story's.
    let mut base: Option<&Section> = None;
    let mut group_start = 0;
    for (k, sec) in sections.iter().enumerate() {
        if !sec.column_widths.is_empty() && sec.columns > 1 {
            let (n, gap) = layout(sec);
            diagnostics.push(Diagnostic::warning(
                format!(
                    "section {} has {n} columns of unequal width: the engine's columns are \
                     equal, so they are laid out as {n} equal columns {} pt apart (Word \
                     keeps the widths and balances them by width)",
                    k + 1,
                    twips(gap)
                ),
                3,
            ));
        }
        let joins = matches!(sec.kind, SectionKind::Continuous | SectionKind::NextColumn);
        let story = out.last().map(|p| p.story).unwrap_or(0);
        let joined = match (base.filter(|_| joins), k.checked_sub(1)) {
            (Some(b), Some(prev)) => join(k, b, &sections[prev], sec, options, diagnostics),
            _ => None,
        };
        let placement = match joined {
            Some(mut p) => {
                p.story = story;
                p
            }
            None => {
                finish_story(sections, &mut out, group_start..k, options, diagnostics);
                group_start = k;
                base = Some(sec);
                SectionPlacement {
                    story: if k == 0 { 0 } else { story + 1 },
                    indent_left_pt: 0.0,
                    indent_right_pt: 0.0,
                    next_column: false,
                    columns: SectionColumns::Frame,
                    frame: StoryColumns {
                        count: 1,
                        gutter_pt: 0.0,
                    },
                }
            }
        };
        out.push(placement);
    }
    let n = sections.len();
    finish_story(sections, &mut out, group_start..n, options, diagnostics);
    out
}

/// Decide the frame columns of the story made of `range` and how each of
/// its sections sits in them (the module docs).
fn finish_story(
    all: &[Section],
    out: &mut [SectionPlacement],
    range: std::ops::Range<usize>,
    options: LowerOptions,
    diagnostics: &mut Vec<Diagnostic>,
) {
    let first = range.start;
    let sections = &all[range.clone()];
    // Word balances a section's columns before a CONTINUOUS break on the
    // same page, also where the next section opens a new story here (a
    // limit, not Word's page): then the story's last section must end
    // balanced. (Before another page size Word starts a page; not asked.)
    let balanced_end = options.mid_page_columns
        && all.get(range.end).is_some_and(|s| {
            s.kind == SectionKind::Continuous
                && sections
                    .first()
                    .is_some_and(|h| (h.page_width, h.page_height) == (s.page_width, s.page_height))
        });
    let placed = &mut out[range];
    let Some(head) = sections.first() else {
        return;
    };
    let mut multi: Vec<(u32, i32)> = Vec::new();
    for s in sections {
        let l = layout(s);
        if l.0 > 1 && !multi.contains(&l) {
            multi.push(l);
        }
    }
    // Only the frame's own columns end UNBALANCED (after the last span the
    // engine fills them in turn), so a story whose last section Word leaves
    // unbalanced keeps them; every other story splits, because the engine
    // balances a split block also where it runs past a page (fill, then
    // balance the rest on the next), and a span balances the text above it
    // only when that text opens its band.
    let unbalanced_last = !balanced_end && sections.last().is_some_and(|s| layout(s).0 > 1);
    let frame = match multi.as_slice() {
        [] => StoryColumns {
            count: 1,
            gutter_pt: twips(head.column_space),
        },
        [(n, gap)] if unbalanced_last => StoryColumns {
            count: *n,
            gutter_pt: twips(*gap),
        },
        _ => StoryColumns {
            count: 1,
            gutter_pt: twips(head.column_space),
        },
    };
    for (s, p) in sections.iter().zip(placed.iter_mut()) {
        let (n, gap) = layout(s);
        p.frame = frame;
        p.columns = match (frame.count, n) {
            (f, 1) if f > 1 => SectionColumns::SpanAll,
            (1, n) if n > 1 => SectionColumns::Split {
                count: n,
                inside_pt: twips(gap),
            },
            _ => SectionColumns::Frame,
        };
    }
    if placed.iter().all(|p| p.columns == SectionColumns::Frame) {
        return;
    }
    let last = first + sections.len();
    let how = if frame.count > 1 {
        format!(
            "the story keeps {} columns and its one-column sections span them",
            frame.count
        )
    } else {
        "each multi-column section splits the story's one column into sub-columns".to_string()
    };
    diagnostics.push(Diagnostic::info(
        format!(
            "sections {}–{} change the columns mid-page, as Word does: {how} (span/split \
             columns, protocol 64)",
            first + 1,
            last
        ),
        3,
    ));
    if frame.count == 1 && unbalanced_last {
        diagnostics.push(Diagnostic::info(
            format!(
                "section {last} ends its story in sub-columns: Word fills its last page's \
                 columns in turn, the engine balances them"
            ),
            3,
        ));
    }
    if frame.count == 1
        && sections
            .iter()
            .skip(1)
            .any(|s| s.kind == SectionKind::NextColumn)
    {
        diagnostics.push(Diagnostic::warning(
            format!(
                "a nextColumn section in sections {}–{} sits in a one-column story with \
                 sub-columns: its column break opens the next page",
                first + 1,
                last
            ),
            3,
        ));
    }
}

/// `sec` (index `k`, continuous or nextColumn, after `prev`) joining the
/// story whose frame is `base`'s, or `None` when it must open a new story.
fn join(
    k: usize,
    base: &Section,
    prev: &Section,
    sec: &Section,
    options: LowerOptions,
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
    let ((prev_cols, prev_gap), (cols, gap)) = (layout(prev), layout(sec));
    if (prev_cols, prev_gap) != (cols, gap) {
        let change = if prev_cols == cols {
            format!("changes the gap between its {cols} columns")
        } else {
            format!("changes from {prev_cols} to {cols} column(s)")
        };
        let why = if sec.kind == SectionKind::NextColumn {
            Some("Word was not asked where a nextColumn section with other columns goes")
        } else if !options.mid_page_columns {
            Some(
                "this engine cannot change columns mid-page (span/split columns need \
                 protocol 64)",
            )
        } else {
            None
        };
        if let Some(why) = why {
            diagnostics.push(Diagnostic::warning(
                format!(
                    "section {} is {kind} and {change}: Word continues it on the same page; \
                     {why}, so it opens a new page",
                    k + 1
                ),
                3,
            ));
            return None;
        }
    }
    if (prev_cols, prev_gap) == (cols, gap) && cols > 1 && sec.kind == SectionKind::Continuous {
        // Measured (columns_docx Q, R): Word balances each of the two
        // sections on its own. The engine starts a new split block only
        // where the count or a gutter changes, and its frame columns have
        // no boundary at all, so the two share one column flow. Kept on
        // the same page (a page break would move every later line).
        diagnostics.push(Diagnostic::warning(
            format!(
                "section {} is continuous with the same {cols} columns as the section before \
                 it: Word balances each section's columns on its own; the engine ends a \
                 column block only where the count or the gap changes, so the two sections \
                 share one column flow (same page, other column breaks)",
                k + 1
            ),
            3,
        ));
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
        columns: SectionColumns::Frame, // finish_story decides
        frame: StoryColumns {
            count: 1,
            gutter_pt: 0.0,
        },
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
    fn a_column_story_before_another_page_size_keeps_its_frame_columns() {
        use SectionKind::*;
        // Word starts a page for another page size; whether it balances
        // the columns before it was not asked, so nothing changes there.
        let mut f = sec(Continuous, 1, 1440, 1440);
        f.page_width = 8640;
        let mut diags = Vec::new();
        let placed = place_sections(&[sec(NextPage, 2, 1440, 1440), f], &mut diags);
        assert_eq!(placed[0].frame.count, 2);
        assert_eq!(placed[0].columns, SectionColumns::Frame);
        assert_eq!(placed[1].story, 1);
        assert!(diags.is_empty(), "{diags:?}");
    }

    #[test]
    fn a_column_change_joins_only_on_an_engine_with_span_columns() {
        use SectionKind::*;
        let doc = [
            sec(NextPage, 1, 1440, 1440),
            sec(Continuous, 2, 1440, 1440),
            sec(Continuous, 1, 1440, 1440),
        ];
        let mut diags = Vec::new();
        let placed = place_sections(&doc, &mut diags);
        assert_eq!(
            placed
                .iter()
                .map(|p| (p.story, p.columns))
                .collect::<Vec<_>>(),
            vec![
                (0, SectionColumns::Frame),
                (
                    0,
                    SectionColumns::Split {
                        count: 2,
                        inside_pt: 36.0
                    }
                ),
                (0, SectionColumns::Frame),
            ]
        );
        assert_eq!(diags.len(), 1);
        assert_eq!(diags[0].severity, "info");

        let mut diags = Vec::new();
        let old = LowerOptions {
            mid_page_columns: false,
        };
        let placed = place_sections_with(&doc, old, &mut diags);
        assert_eq!(
            placed.iter().map(|p| p.story).collect::<Vec<_>>(),
            vec![0, 1, 2]
        );
        assert_eq!(placed[1].frame.count, 2);
        assert!(placed.iter().all(|p| p.columns == SectionColumns::Frame));
        assert_eq!(diags.iter().filter(|d| d.severity == "warning").count(), 2);
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
