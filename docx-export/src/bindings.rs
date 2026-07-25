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

//! The native↔OOXML provenance map (the base-idea's `bindings.json`), built at
//! import time. It resolves a lowered story `(block, run)` coordinate to the
//! source `<w:p>`/`<w:r>` ordinals the byte-splice patcher locates on.
//!
//! Indexing mirrors the Lowered story exactly: `blocks[i]` corresponds to lowered
//! story block `i` (== `doc.body[i]`, since both `lower()` and `build_bindings`
//! walk `doc.body` in order). Within a paragraph, run indices replay lowering's
//! filter (`!text.is_empty()`) so they line up 1:1 with `LoweredRun`s.

use std::collections::HashMap;

use docx_core::{Block, DocxDocument, RunSource, StyleKind};

/// The provenance map for one document.
#[derive(Debug, Clone, Default)]
pub struct DocxBindings {
    /// One entry per lowered story block, in order.
    pub blocks: Vec<BlockBinding>,
    /// Lowered character-style token → the original Word `styleId`. The differ
    /// uses it to recover a real `<w:rStyle>` when projecting a changed run's
    /// style back (the lowered token is lossy, so this map is the only inverse).
    pub char_token_to_style_id: HashMap<String, String>,
}

/// A lowered story block's provenance.
#[derive(Debug, Clone)]
pub enum BlockBinding {
    /// A body paragraph: its `<w:p>` ordinal + a binding per lowered run.
    Paragraph {
        para_ord: u32,
        runs: Vec<RunBinding>,
    },
    /// A table — non-patchable in the current increment (cell content has no
    /// stable body-level provenance yet).
    Table,
}

/// A lowered run's provenance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RunBinding {
    /// The `run_ord`-th direct `<w:r>` child of the paragraph — patchable.
    Direct { run_ord: u32 },
    /// A run with no stable direct-`<w:r>` provenance (hyperlink/field-flattened,
    /// or a linked run) — not patched in the current increment.
    NonPatchable,
}

impl DocxBindings {
    /// Resolve a lowered `(block, run)` to `(para_ord, run_ord)` in the source
    /// `word/document.xml`, or `None` when the target is non-patchable (a table,
    /// an out-of-range index, or a hyperlink/field/linked run).
    pub fn resolve(&self, block: usize, run: usize) -> Option<(u32, u32)> {
        match self.blocks.get(block)? {
            BlockBinding::Paragraph { para_ord, runs } => match runs.get(run)? {
                RunBinding::Direct { run_ord } => Some((*para_ord, *run_ord)),
                RunBinding::NonPatchable => None,
            },
            BlockBinding::Table => None,
        }
    }
}

/// Build the bindings from the imported model. `Run.source` / `Paragraph.
/// source_para_ord` were stamped by `docx-import`; here we only project them onto
/// lowered-story coordinates, replaying the lowering run filter so indices match.
pub fn build_bindings(doc: &DocxDocument) -> DocxBindings {
    let mut blocks = Vec::with_capacity(doc.body.len());
    for block in &doc.body {
        match block {
            Block::Paragraph(p) => {
                let runs = p
                    .runs
                    .iter()
                    // Lowering drops empty-text runs (docx-lower `lower_paragraph`),
                    // so replay that filter to keep run indices aligned.
                    .filter(|r| !r.text.is_empty())
                    .map(|r| match (r.source, r.hyperlink.is_some()) {
                        // A direct run that is NOT a link is patchable. A link run
                        // (even a direct one, e.g. a HYPERLINK-field result) is
                        // deferred — editing it would desync the field/hyperlink.
                        (Some(RunSource::DirectRun(n)), false) => RunBinding::Direct { run_ord: n },
                        _ => RunBinding::NonPatchable,
                    })
                    .collect();
                blocks.push(BlockBinding::Paragraph {
                    para_ord: p.source_para_ord,
                    runs,
                });
            }
            Block::Table(_) => blocks.push(BlockBinding::Table),
        }
    }
    let char_token_to_style_id = doc
        .styles
        .styles
        .iter()
        .filter(|s| s.kind == StyleKind::Character)
        .map(|s| {
            (
                docx_lower::char_style_token(&s.style_id),
                s.style_id.clone(),
            )
        })
        .collect();
    DocxBindings {
        blocks,
        char_token_to_style_id,
    }
}
