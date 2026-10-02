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

//! ADR 029 — the room between two paragraphs is the room WORD leaves.
//! `paragraph_spacing_docx()` is the document Word laid out
//! (`scripts/word-paragraph-spacing-probe.sh`); the baseline step of each
//! pair is `fixtures/paragraph-spacing.word.json`. Word leaves the larger
//! of space after and space before; the engine adds them. So each pair's
//! lowered space after + space before must equal Word's step less one line.

// `__feat__<id>` names link these tests to the Cockpit feature row.
#![allow(non_snake_case)]

use docx_conformance::{paragraph_spacing_docx, SPACING_PAIRS};
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue};

/// `(pair, step_pt)` rows of `paragraph-spacing.word.json` (one object per
/// line; no JSON dependency in this crate).
fn word_steps() -> Vec<(String, f32)> {
    let json = include_str!("../fixtures/paragraph-spacing.word.json");
    let field = |row: &str, key: &str| -> String {
        let at = row.find(&format!("\"{key}\": ")).expect(key) + key.len() + 4;
        row[at..]
            .split([',', ' ', '}'])
            .next()
            .unwrap()
            .trim_matches('"')
            .to_string()
    };
    json.lines()
        .filter(|l| l.trim_start().starts_with("{ \"pair\""))
        .map(|row| (field(row, "pair"), field(row, "step_pt").parse().unwrap()))
        .collect()
}

/// The length a paragraph style resolves `path` to through `basedOn`.
fn resolved(doc: &LoweredDoc, style_id: Option<&str>, path: &str) -> f32 {
    let mut next = style_id.map(str::to_string);
    for _ in 0..32 {
        let Some(style) = doc.styles.iter().find(|s| Some(&s.id) == next.as_ref()) else {
            break;
        };
        if let Some(PropValue::Length(pt)) = style
            .props
            .iter()
            .find(|p| p.path == path)
            .map(|p| &p.value)
        {
            return *pt;
        }
        next = style.based_on.clone();
    }
    0.0
}

fn check(compat: Option<u32>) {
    let doc = docx_import::import_docx(&paragraph_spacing_docx(compat)).expect("import");
    let lowered = docx_lower::lower(&doc);
    let style = |i: usize| match &lowered.story.blocks[i] {
        LoweredBlock::Paragraph(p) => p.para_style_id.clone(),
        LoweredBlock::Table(_) => panic!("no tables in the fixture"),
    };
    let steps = word_steps();
    assert_eq!(steps.len(), SPACING_PAIRS.len());
    let pitch = steps[0].1;
    let mut wrong = Vec::new();
    for (k, (pair, (label, step))) in SPACING_PAIRS.iter().zip(&steps).enumerate() {
        assert_eq!(pair.label, label);
        // Block 0 is the first line; each pair is `a`, `b`, `sep`.
        let (a, b) = (style(1 + 3 * k), style(2 + 3 * k));
        let ours = resolved(&lowered, a.as_deref(), "paragraphSpaceAfter")
            + resolved(&lowered, b.as_deref(), "paragraphSpaceBefore");
        let word = step - pitch;
        if (ours - word).abs() > 0.3 {
            wrong.push(format!(
                "{label} (after {} / before {} twips): {ours:.2} pt between the lines, Word {word:.2}",
                pair.after, pair.before
            ));
        }
    }
    assert!(wrong.is_empty(), "compat {compat:?}\n{}", wrong.join("\n"));
}

#[test]
fn the_room_between_paragraphs_is_words__feat__plugin_doc_word_pagination() {
    check(Some(15));
    check(Some(11));
    check(None);
}
