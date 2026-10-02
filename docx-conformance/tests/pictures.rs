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

//! ADR 029 — a line holding an inline picture is as tall as the picture.
//! Word grows the line; the engine (InDesign's rule) grows a line for its
//! object only under AUTO leading, and every lowered style carries Word's
//! pitch as a fixed one. So a body paragraph with a picture takes auto
//! leading (`characterLeading` 0). Measured on real documents
//! (`docs/acceptance-real-docx.md`, round 12): eight pictures took no room.

// `__feat__<id>` names link these tests to the Cockpit feature row.
#![allow(non_snake_case)]

use docx_conformance::image_docx;
use docx_lower::ir::{LoweredBlock, LoweredDoc, PropValue};

/// The `characterLeading` a paragraph style resolves to through `basedOn`.
fn resolved_leading(doc: &LoweredDoc, style_id: Option<&str>) -> Option<f32> {
    let mut next = style_id.map(str::to_string);
    for _ in 0..32 {
        let style = doc.styles.iter().find(|s| Some(&s.id) == next.as_ref())?;
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
fn a_paragraph_with_a_picture_takes_auto_leading__feat__plugin_doc_word_pagination() {
    let doc = docx_import::import_docx(&image_docx()).expect("import");
    let lowered = docx_lower::lower(&doc);
    let mut with_picture = 0;
    for block in &lowered.story.blocks {
        let LoweredBlock::Paragraph(p) = block else {
            continue;
        };
        let leading = resolved_leading(&lowered, p.para_style_id.as_deref());
        if p.images.is_empty() {
            assert!(
                leading.is_some_and(|pt| pt > 0.0),
                "a paragraph of text keeps Word's pitch as its leading, got {leading:?}"
            );
        } else {
            with_picture += 1;
            assert_eq!(leading, Some(0.0), "auto leading: the line grows");
        }
    }
    assert!(with_picture > 0, "the fixture has a picture paragraph");
}
