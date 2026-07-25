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

use docx_core::{Block, CellPath, DocxDocument, RunSource, StyleKind, VMerge};

/// The provenance map for one document.
#[derive(Debug, Clone, Default)]
pub struct DocxBindings {
    /// One entry per lowered story block, in order.
    pub blocks: Vec<BlockBinding>,
    /// Lowered character-style token → the original Word `styleId`. The differ
    /// uses it to recover a real `<w:rStyle>` when projecting a changed run's
    /// style back (the lowered token is lossy, so this map is the only inverse).
    pub char_token_to_style_id: HashMap<String, String>,
    /// The same, for PARAGRAPH styles (`<w:pStyle>`).
    pub para_token_to_style_id: HashMap<String, String>,
}

/// A lowered story block's provenance.
#[derive(Debug, Clone)]
pub enum BlockBinding {
    /// A body paragraph: its `<w:p>` ordinal + a binding per lowered run.
    Paragraph {
        para_ord: u32,
        runs: Vec<RunBinding>,
    },
    /// A table: one entry per LOWERED cell, in the same order `docx-lower`
    /// emits them (vMerge-continue cells are absorbed, not emitted).
    Table { cells: Vec<CellBinding> },
}

/// A lowered table cell's provenance: one entry per paragraph in the cell.
#[derive(Debug, Clone, Default)]
pub struct CellBinding {
    pub paragraphs: Vec<CellParaBinding>,
}

/// A cell paragraph: its `w:tbl`/`w:tr`/`w:tc`/`w:p` source path + its runs.
#[derive(Debug, Clone)]
pub struct CellParaBinding {
    pub path: CellPath,
    pub runs: Vec<RunBinding>,
}

/// Which element a patchable run sits inside, when it is not a direct `<w:r>`
/// child of the paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Wrapper {
    /// Inside the n-th `<w:hyperlink>` child (the `r:id` lives on the wrapper).
    Hyperlink(u32),
    /// Inside the n-th `<w:fldSimple>` child (the instruction is a wrapper attr).
    Field(u32),
}

/// A lowered run's provenance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RunBinding {
    /// The `run_ord`-th direct `<w:r>` child of the paragraph — patchable.
    Direct { run_ord: u32 },
    /// The `run_ord`-th `<w:r>` inside a hyperlink/field WRAPPER. Patchable: the
    /// link target / field instruction lives on the wrapper, not the run, so
    /// rewriting the run's `<w:t>`/`<w:rPr>` cannot desync it.
    Wrapped { wrapper: Wrapper, run_ord: u32 },
    /// No stable provenance — not patched.
    NonPatchable,
}

/// A resolved run address in the source document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RunAddr {
    pub para_ord: u32,
    pub run_ord: u32,
    /// `None` ⇒ a direct `<w:r>` child of the `<w:p>`.
    pub wrapper: Option<Wrapper>,
}

impl DocxBindings {
    /// Resolve a lowered `(block, run)` to `(para_ord, run_ord)` in the source
    /// `word/document.xml`, or `None` when the target is non-patchable (a table,
    /// an out-of-range index, or a hyperlink/field/linked run).
    pub fn resolve(&self, block: usize, run: usize) -> Option<RunAddr> {
        match self.blocks.get(block)? {
            BlockBinding::Paragraph { para_ord, runs } => match runs.get(run)? {
                RunBinding::Direct { run_ord } => Some(RunAddr {
                    para_ord: *para_ord,
                    run_ord: *run_ord,
                    wrapper: None,
                }),
                RunBinding::Wrapped { wrapper, run_ord } => Some(RunAddr {
                    para_ord: *para_ord,
                    run_ord: *run_ord,
                    wrapper: Some(*wrapper),
                }),
                RunBinding::NonPatchable => None,
            },
            BlockBinding::Table { .. } => None,
        }
    }

    /// The source `<w:p>` ordinal of a lowered story block, or `None` when the
    /// block is a table (or out of range) — the address paragraph-level
    /// structural ops resolve on.
    pub fn para_ord(&self, block: usize) -> Option<u32> {
        match self.blocks.get(block)? {
            BlockBinding::Paragraph { para_ord, .. } => Some(*para_ord),
            BlockBinding::Table { .. } => None,
        }
    }

    /// Resolve a TABLE-CELL run to its `(CellPath, run_ord)` source address, or
    /// `None` when the target is non-patchable (not a table, out of range, or a
    /// hyperlink/field run).
    pub fn resolve_cell(
        &self,
        block: usize,
        cell: usize,
        para: usize,
        run: usize,
    ) -> Option<(CellPath, u32)> {
        let BlockBinding::Table { cells } = self.blocks.get(block)? else {
            return None;
        };
        let cp = cells.get(cell)?.paragraphs.get(para)?;
        match cp.runs.get(run)? {
            RunBinding::Direct { run_ord } => Some((cp.path, *run_ord)),
            // A wrapped run inside a table cell needs both locator paths at once;
            // not addressed in the current increment.
            RunBinding::Wrapped { .. } | RunBinding::NonPatchable => None,
        }
    }
}

/// One paragraph's run bindings, replaying lowering's run filter so indices line
/// up 1:1 with `LoweredRun`s.
fn run_bindings(p: &docx_core::Paragraph) -> Vec<RunBinding> {
    p.runs
        .iter()
        // Lowering drops empty-text runs (docx-lower `lower_paragraph`), so
        // replay that filter to keep run indices aligned.
        .filter(|r| !r.text.is_empty())
        .map(|r| match r.source {
            // A direct `<w:r>` child — patchable. This INCLUDES a complex field's
            // RESULT run (it carries a hyperlink target, but the URL lives in a
            // separate `instrText` run, so rewriting this run cannot desync it).
            Some(RunSource::DirectRun(n)) => RunBinding::Direct { run_ord: n },
            // Wrapped runs are patchable through the wrapper's own locator path;
            // the link target / field instruction lives on the WRAPPER.
            Some(RunSource::Hyperlink { link_ord, run_ord }) => RunBinding::Wrapped {
                wrapper: Wrapper::Hyperlink(link_ord),
                run_ord,
            },
            Some(RunSource::Field { field_ord, run_ord }) => RunBinding::Wrapped {
                wrapper: Wrapper::Field(field_ord),
                run_ord,
            },
            None => RunBinding::NonPatchable,
        })
        .collect()
}

/// Build the bindings from the imported model. `Run.source` / `Paragraph.
/// source_para_ord` / `Paragraph.source_cell` were stamped by `docx-import`; here
/// we only project them onto lowered-story coordinates, replaying the lowering
/// run filter + table-cell emission order so indices match.
pub fn build_bindings(doc: &DocxDocument) -> DocxBindings {
    let mut blocks = Vec::with_capacity(doc.body.len());
    for block in &doc.body {
        match block {
            Block::Paragraph(p) => {
                blocks.push(BlockBinding::Paragraph {
                    para_ord: p.source_para_ord,
                    runs: run_bindings(p),
                });
            }
            // Replay `docx-lower::lower_table`'s cell emission EXACTLY (a
            // vMerge-continue cell is absorbed into its restart cell above and
            // not emitted) so cell indices line up with `LoweredTable.cells`.
            Block::Table(t) => {
                let mut cells: Vec<CellBinding> = Vec::new();
                for row in &t.rows {
                    for cell in &row.cells {
                        if cell.v_merge == VMerge::Continue {
                            continue;
                        }
                        cells.push(CellBinding {
                            paragraphs: cell
                                .paragraphs
                                .iter()
                                .map(|p| CellParaBinding {
                                    path: p.source_cell.unwrap_or_default(),
                                    runs: run_bindings(p),
                                })
                                .collect(),
                        });
                    }
                }
                blocks.push(BlockBinding::Table { cells });
            }
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
    let para_token_to_style_id = doc
        .styles
        .styles
        .iter()
        .filter(|s| s.kind == StyleKind::Paragraph)
        .map(|s| {
            (
                docx_lower::para_style_token(&s.style_id),
                s.style_id.clone(),
            )
        })
        .collect();
    DocxBindings {
        blocks,
        char_token_to_style_id,
        para_token_to_style_id,
    }
}
