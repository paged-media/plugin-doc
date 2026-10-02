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

//! ADR 029 — a table row taller than its page flows as text. Word splits
//! such a row across pages; a native table row never splits (InDesign's
//! rule), so the whole row is overset and its text gone. Found on a real
//! document (`docs/reference/acceptance-real-docx.md`, round 14): a web page saved as
//! Word keeps its article, 46,409 characters, in one cell of a layout table,
//! and 12 of Word's 52 pages were missing.

// `__feat__<id>` names link these tests to the Cockpit feature row.
#![allow(non_snake_case)]

use docx_conformance::{tall_row_docx, TALL_ROW_PARAGRAPHS};
use docx_export::{overlay_story_content, ParagraphContentIn, RunContentIn, StoryContentIn};
use docx_lower::ir::{LoweredBlock, LoweredDoc};

fn lowered() -> LoweredDoc {
    let doc = docx_import::import_docx(&tall_row_docx()).expect("import");
    docx_lower::lower(&doc)
}

#[test]
fn a_row_taller_than_the_page_flows_as_text__feat__plugin_doc_word_pagination() {
    let lowered = lowered();
    // before, the tall table, after, the small table, end: blocks stay 1:1
    // with Word's body blocks.
    assert_eq!(lowered.story.blocks.len(), 5);
    let LoweredBlock::Table(tall) = &lowered.story.blocks[1] else {
        panic!("block 1 is the tall table");
    };
    assert_eq!(
        tall.flow.len(),
        TALL_ROW_PARAGRAPHS,
        "every cell paragraph flows"
    );
    assert!(
        tall.cells.is_empty() && tall.rows == 0,
        "and no native table is built"
    );
    assert_eq!(tall.flow[0].runs[0].text, "Row text 001");
    let LoweredBlock::Table(small) = &lowered.story.blocks[3] else {
        panic!("block 3 is the small table");
    };
    assert!(small.flow.is_empty(), "a row that fits stays a table row");
    assert_eq!((small.rows, small.cols), (1, 1));
    assert!(
        lowered
            .diagnostics
            .iter()
            .any(|d| d.message.contains("flow as text")),
        "the user is told"
    );
}

#[test]
fn save_back_steps_over_a_flowed_table__feat__plugin_doc_word_pagination() {
    let baseline = lowered();
    let para = |text: &str| ParagraphContentIn {
        paragraph_style: None,
        runs: vec![RunContentIn {
            text: text.to_string(),
            character_style: None,
        }],
    };
    // What the editor reads back: before, the flowed paragraphs, after
    // (edited), end. The small table is not among the paragraphs.
    let mut paragraphs = vec![para("before")];
    paragraphs.extend((1..=TALL_ROW_PARAGRAPHS).map(|n| para(&format!("Row text {n:03}"))));
    paragraphs.push(para("after, edited"));
    paragraphs.push(para("end"));
    let edited = overlay_story_content(
        &baseline,
        &StoryContentIn {
            self_id: "s".into(),
            paragraphs,
        },
    );
    let text = |i: usize| match &edited.story.blocks[i] {
        LoweredBlock::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect::<String>(),
        LoweredBlock::Table(_) => panic!("block {i} is a table"),
    };
    assert_eq!(text(0), "before");
    assert_eq!(
        text(2),
        "after, edited",
        "the edit lands on its own paragraph"
    );
    assert_eq!(text(4), "end");
}
