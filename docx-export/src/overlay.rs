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

//! The DOC-03 read → edited-`LoweredDoc` mapper. The host's structured read
//! (`host.document.storyContent`, core wire `StoryContent`) hands back the EDITED
//! story as paragraphs → runs → text + applied styles. Because paged.doc created
//! every style via `applyStyle`, a run's `characterStyle` IS the plugin's own
//! `char_style_id` token — so the read overlays cleanly onto the import baseline,
//! yielding the "edited" lowering that `diff` compares against the baseline.
//!
//! Isolation-clean: this mirrors the core wire shape as a plugin-local type
//! (deserialized from the JSON the bundle forwards) — no core dependency.

use docx_lower::ir::{LoweredBlock, LoweredDoc, LoweredRun};
use serde::{Deserialize, Serialize};

/// Plugin-local twin of the core `StoryContent` wire struct (camelCase JSON).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryContentIn {
    pub self_id: String,
    pub paragraphs: Vec<ParagraphContentIn>,
}

/// One read-back paragraph.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphContentIn {
    #[serde(default)]
    pub paragraph_style: Option<String>,
    pub runs: Vec<RunContentIn>,
}

/// One read-back run. Only `text` + `character_style` drive the overlay today
/// (direct-override fields are carried for forward use but not yet applied).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunContentIn {
    pub text: String,
    #[serde(default)]
    pub character_style: Option<String>,
}

/// Overlay the read-back story onto the import `baseline`, producing the EDITED
/// lowering. Structure-preserving: read-back paragraphs map 1:1 onto the
/// baseline's PARAGRAPH blocks in order (table blocks are skipped and left as-is
/// — cell content is not yet round-tripped); a paragraph whose run count no
/// longer matches is left untouched (structural edits are a later increment).
pub fn overlay_story_content(baseline: &LoweredDoc, content: &StoryContentIn) -> LoweredDoc {
    let mut edited = baseline.clone();
    let mut ci = 0usize;
    for block in edited.story.blocks.iter_mut() {
        let LoweredBlock::Paragraph(p) = block else {
            continue; // table — not overlaid
        };
        let Some(cp) = content.paragraphs.get(ci) else {
            break; // read-back ran out of paragraphs
        };
        ci += 1;
        if cp.runs.len() == p.runs.len() {
            for (run, cr) in p.runs.iter_mut().zip(&cp.runs) {
                run.text = cr.text.clone();
                run.char_style_id = cr.character_style.clone();
            }
        } else {
            // Increment 2 — the run list CHANGED (a run was added or removed in
            // the editor). Replace it wholesale; `diff` aligns the two lists by
            // (text, style) identity and emits the insert/delete ops.
            let template = p.runs.first().cloned().unwrap_or_default();
            p.runs = cp
                .runs
                .iter()
                .map(|cr| LoweredRun {
                    text: cr.text.clone(),
                    char_style_id: cr.character_style.clone(),
                    ..template.clone()
                })
                .collect();
        }
    }
    edited
}
