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

//! Widow control and keepLines (`keeps_docx()`, ADR 029 decision 3) against
//! Word's answer (`fixtures/keeps.word.json`, `scripts/word-keeps-probe.sh`).

#![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature

use docx_conformance::keeps_docx;
use docx_import::import_docx;
use docx_lower::ir::{LoweredDoc, PropValue};

fn lowered() -> LoweredDoc {
    docx_lower::lower(&import_docx(&keeps_docx()).expect("import"))
}

/// Word's pages: the line labels on each, top to bottom.
fn word_pages() -> Vec<Vec<String>> {
    let v: serde_json::Value =
        serde_json::from_str(include_str!("../fixtures/keeps.word.json")).unwrap();
    v["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|p| {
            p.as_array()
                .unwrap()
                .iter()
                .map(|l| l.as_str().unwrap().to_string())
                .collect()
        })
        .collect()
}

fn page_of(line: &str) -> usize {
    word_pages()
        .iter()
        .position(|p| p.iter().any(|l| l == line))
        .unwrap_or_else(|| panic!("{line}"))
        + 1
}

/// The keep properties a paragraph resolves to through its style chain.
fn keeps(ir: &LoweredDoc, label: &str) -> (Option<bool>, Option<bool>, Option<f32>, Option<f32>) {
    let p = ir
        .story
        .paragraphs()
        .into_iter()
        .find(|p| p.runs.iter().any(|r| r.text.starts_with(label)))
        .unwrap()
        .clone();
    let mut next = p.para_style_id.clone();
    let (mut together, mut all, mut first, mut last) = (None, None, None, None);
    while let Some(id) = next {
        let st = ir.styles.iter().find(|s| s.id == id).unwrap();
        for sp in &st.props {
            match (sp.path.as_str(), &sp.value) {
                ("paragraphKeepLinesTogether", PropValue::Bool(b)) => {
                    together.get_or_insert(*b);
                }
                ("paragraphKeepAllLinesTogether", PropValue::Bool(b)) => {
                    all.get_or_insert(*b);
                }
                ("paragraphKeepFirstLines", PropValue::Length(n)) => {
                    first.get_or_insert(*n);
                }
                ("paragraphKeepLastLines", PropValue::Length(n)) => {
                    last.get_or_insert(*n);
                }
                _ => {}
            }
        }
        next = st.based_on.clone();
    }
    (together, all, first, last)
}

/// Word, 36 lines a page: K1 (widow control, one line fits) moves whole to
/// page 2; K2 (widow control off) leaves its first line on page 2; K3
/// (keepLines, two of three fit) moves whole to page 4; K4 (widow control,
/// three of four fit) splits 2 | 2; K5 (no keep element ANYWHERE in the
/// hierarchy) moves whole like K1 — Word's widow control is on by default.
/// The native rules that say the same: KeepLinesTogether At Start / At End
/// 2 / 2 for widow control, All Lines for keepLines, off for `w:val="0"`.
#[test]
fn widow_control_and_keep_lines_lower_to_the_rules_word_follows__feat__plugin_doc_read_path() {
    assert_eq!((page_of("F035"), page_of("K1a"), page_of("K1c")), (1, 2, 2));
    assert_eq!((page_of("K2a"), page_of("K2b")), (2, 3));
    assert_eq!((page_of("F099"), page_of("K3a"), page_of("K3c")), (3, 4, 4));
    assert_eq!((page_of("K4b"), page_of("K4c")), (4, 5));
    assert_eq!((page_of("F162"), page_of("K5a")), (5, 6));

    let ir = lowered();
    let widow = (Some(true), None, Some(2.0), Some(2.0));
    assert_eq!(keeps(&ir, "K1"), widow);
    assert_eq!(keeps(&ir, "K4"), widow);
    assert_eq!(keeps(&ir, "K5"), widow, "Word's default");
    assert_eq!(keeps(&ir, "F001"), widow, "Word's default");
    assert_eq!(keeps(&ir, "K2").0, Some(false));
    let k3 = keeps(&ir, "K3");
    assert_eq!((k3.0, k3.1), (Some(true), Some(true)));
}
