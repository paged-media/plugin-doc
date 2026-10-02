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

//! ADR 029 — Word sizes each LINE by the fonts on it: a line set
//! entirely in 10 pt steps 11.5 pt where the paragraph's 12 pt lines step
//! 13.8 (measured on a real document, `docs/reference/acceptance-real-docx.md` round
//! 13: 21 paragraphs ending in a 10 pt citation were 2.3 pt too tall). The
//! engine gives a line the largest leading among its characters (InDesign's
//! rule, core's `mixed-leading`), so runs whose face differs from the
//! paragraph's first run carry their own leading. A run of nothing but
//! spaces takes its neighbour's: Word does not let it make a line taller.

// `__feat__<id>` names link these tests to the Cockpit feature row.
#![allow(non_snake_case)]

use docx_conformance::mixed_sizes_docx;
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue, StyleCollection};

/// The `characterLeading` a style resolves to through `basedOn`.
fn leading(doc: &LoweredDoc, collection: StyleCollection, style_id: Option<&str>) -> Option<f32> {
    let mut next = style_id.map(str::to_string);
    for _ in 0..32 {
        let style = doc
            .styles
            .iter()
            .find(|s| s.collection == collection && Some(&s.id) == next.as_ref())?;
        if let Some(PropValue::Length(pt)) = style
            .props
            .iter()
            .find(|p| p.path == "characterLeading")
            .map(|p| &p.value)
        {
            return Some(*pt);
        }
        next = style.based_on.clone();
    }
    None
}

#[test]
fn runs_of_another_size_carry_their_own_leading__feat__plugin_doc_word_pagination() {
    let doc = docx_import::import_docx(&mixed_sizes_docx()).expect("import");
    let lowered = docx_lower::lower(&doc);
    let paras: Vec<_> = lowered
        .story
        .blocks
        .iter()
        .map(|b| match b {
            LoweredBlock::Paragraph(p) => p,
            LoweredBlock::Table(_) => panic!("no tables in the fixture"),
        })
        .collect();
    let near = |got: Option<f32>, want: f32| got.is_some_and(|pt| (pt - want).abs() < 0.05);
    let para_leading = |i: usize| {
        leading(
            &lowered,
            StyleCollection::Paragraph,
            paras[i].para_style_id.as_deref(),
        )
    };
    let run_leading = |i: usize, r: usize| {
        leading(
            &lowered,
            StyleCollection::Character,
            paras[i].runs[r].char_style_id.as_deref(),
        )
    };
    // Times New Roman single spacing: 13.8 pt at 12 pt, 11.5 pt at 10 pt
    // (`fixtures/line-spacing.word.json`).
    assert!(
        near(para_leading(0), 13.8),
        "the paragraph takes its first run's pitch: {:?}",
        para_leading(0)
    );
    assert_eq!(run_leading(0, 0), None, "the first run inherits it");
    assert!(
        near(run_leading(0, 1), 11.5),
        "the 10 pt run carries 11.5: {:?}",
        run_leading(0, 1)
    );
    assert!(
        near(run_leading(0, 2), 11.5),
        "the space after it does not make its line taller: {:?}",
        run_leading(0, 2)
    );
    assert_eq!(
        run_leading(0, 3),
        None,
        "the 12 pt run after it inherits again"
    );
    // A paragraph whose runs agree has one leading and no run carries any.
    assert!(near(para_leading(1), 13.8));
    assert_eq!(run_leading(1, 0), None);
    assert_eq!(run_leading(1, 1), None);
}
