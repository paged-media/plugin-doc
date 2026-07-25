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

//! paged.doc M2 edited save-back — the targeted-patch export engine.
//!
//! `apply_edits` resolves an [`EditSet`] (keyed by lowered story coordinates)
//! through the import-built [`DocxBindings`] to source `<w:p>`/`<w:r>` ordinals,
//! renders the replacement `<w:rPr>`/`<w:t>` fragments, byte-splices only those
//! holes in `word/document.xml`, and writes the patched part back into the
//! retained [`OpcPackage`]. Every other part — and every untouched subtree — is
//! re-emitted byte-identical (the preservation invariant). The ooxmlsdk
//! serializer is never linked (the wasm-budget guard).

mod bindings;
mod diff;
mod edit;
mod overlay;
mod rpr;
mod splice;

pub use bindings::{build_bindings, BlockBinding, DocxBindings, RunBinding};
pub use diff::diff;
pub use edit::{CellRunEdit, EditSet, ParaEdit, RunEdit, StructuralEdit};
pub use overlay::{overlay_story_content, ParagraphContentIn, RunContentIn, StoryContentIn};

use std::collections::BTreeMap;

use paged_ooxml::{OoxmlError, OpcPackage};
use splice::{
    patch_document_xml_cols, ColumnAction, ResolvedCellTarget, ResolvedParaTarget,
    ResolvedRowTarget, ResolvedTarget,
};

/// Apply `edits` to `pkg`'s main document part in place. Non-patchable targets
/// (tables, hyperlink/field runs, out-of-range indices) are silently skipped —
/// the caller surfaces them as diagnostics. A no-op edit set leaves `pkg`
/// untouched. On success the main part is dirtied; `pkg.write()` then yields the
/// saved `.docx` bytes.
pub fn apply_edits(
    pkg: &mut OpcPackage,
    main_part: &str,
    bindings: &DocxBindings,
    edits: &EditSet,
) -> Result<(), OoxmlError> {
    let mut targets: Vec<ResolvedTarget> = Vec::new();
    for e in &edits.runs {
        let Some(addr) = bindings.resolve(e.block, e.run) else {
            continue; // non-patchable — skipped
        };
        if e.new_text.is_none() && e.new_props.is_none() {
            continue;
        }
        let new_rpr = e.new_props.as_ref().map(|props| {
            // `Some(Some(id))` sets a real Word rStyle; anything else omits it
            // (a whole-`rPr` replacement can't preserve an old rStyle, so the
            // differ passes it explicitly).
            let rstyle = e.rstyle.as_ref().and_then(|o| o.as_deref());
            rpr::render_rpr(props, rstyle)
        });
        targets.push(ResolvedTarget {
            para_ord: addr.para_ord,
            run_ord: addr.run_ord,
            wrapper: addr.wrapper,
            new_text: e.new_text.clone(),
            new_rpr,
            delete: false,
            insert_after: Vec::new(),
        });
    }

    // Increment 2 — structural ops. All coordinates address the BASELINE, so
    // ops never shift each other's addresses; the patcher applies them in one
    // pass over the unmodified source.
    let mut paras: BTreeMap<u32, ResolvedParaTarget> = BTreeMap::new();
    let mut rows: BTreeMap<(u32, u32), ResolvedRowTarget> = BTreeMap::new();
    // At most ONE column action per table per save (they shift each other's
    // indices, so a second would need re-resolution against the patched grid).
    let mut columns: BTreeMap<u32, ColumnAction> = BTreeMap::new();
    for s in &edits.structural {
        match s {
            StructuralEdit::DeleteRun { block, run } => {
                let Some(a) = bindings.resolve(*block, *run) else {
                    continue;
                };
                match targets.iter_mut().find(|t| {
                    t.para_ord == a.para_ord && t.run_ord == a.run_ord && t.wrapper == a.wrapper
                }) {
                    Some(t) => t.delete = true,
                    None => {
                        let mut t = ResolvedTarget::edit(a.para_ord, a.run_ord);
                        t.wrapper = a.wrapper;
                        t.delete = true;
                        targets.push(t);
                    }
                }
            }
            StructuralEdit::InsertRun {
                block,
                run,
                text,
                props,
                rstyle,
            } => {
                let frag = rpr::render_run(text, props, rstyle.as_deref());
                match run {
                    // After an existing run.
                    Some(run) => {
                        let Some(a) = bindings.resolve(*block, *run) else {
                            continue;
                        };
                        match targets.iter_mut().find(|t| {
                            t.para_ord == a.para_ord
                                && t.run_ord == a.run_ord
                                && t.wrapper == a.wrapper
                        }) {
                            Some(t) => t.insert_after.push(frag),
                            None => {
                                let mut t = ResolvedTarget::edit(a.para_ord, a.run_ord);
                                t.wrapper = a.wrapper;
                                t.insert_after.push(frag);
                                targets.push(t);
                            }
                        }
                    }
                    // At the paragraph's start.
                    None => {
                        let Some(p) = bindings.para_ord(*block) else {
                            continue;
                        };
                        paras.entry(p).or_default().prepend_runs.push(frag);
                    }
                }
            }
            // Column ops require a UNIFORM grid (no gridSpan) — otherwise
            // whether a new column widens a span or splits it is ambiguous, so
            // the op is skipped rather than corrupting the table.
            StructuralEdit::DeleteColumn { block, col } => {
                let Some(t) = bindings.uniform_table_ord(*block) else {
                    continue;
                };
                columns.insert(t, ColumnAction::Delete { col: *col });
            }
            StructuralEdit::InsertColumn {
                block,
                after_col,
                text,
            } => {
                let Some(t) = bindings.uniform_table_ord(*block) else {
                    continue;
                };
                columns.insert(
                    t,
                    ColumnAction::Insert {
                        after_col: *after_col,
                        text: text.clone(),
                    },
                );
            }
            StructuralEdit::DeleteRow { block, row } => {
                let Some(t) = bindings.table_ord(*block) else {
                    continue;
                };
                rows.entry((t, *row)).or_default().delete = true;
            }
            StructuralEdit::InsertRow {
                block,
                after_row,
                cells,
            } => {
                let Some(t) = bindings.table_ord(*block) else {
                    continue;
                };
                rows.entry((t, *after_row))
                    .or_default()
                    .insert_after
                    .push(rpr::render_table_row(cells));
            }
            StructuralEdit::DeleteParagraph { block } => {
                let Some(p) = bindings.para_ord(*block) else {
                    continue;
                };
                paras.entry(p).or_default().delete = true;
            }
            StructuralEdit::InsertParagraph {
                block,
                text,
                props,
                para_style,
                rstyle,
            } => {
                let Some(p) = bindings.para_ord(*block) else {
                    continue;
                };
                let frag =
                    rpr::render_paragraph(text, props, rstyle.as_deref(), para_style.as_deref());
                paras.entry(p).or_default().insert_after.push(frag);
            }
        }
    }

    // Increment 3 — paragraph `<w:pPr>` edits.
    for pe in &edits.paragraphs {
        let Some(p) = bindings.para_ord(pe.block) else {
            continue; // a table (or out of range) — skipped
        };
        let pstyle = pe.pstyle.as_ref().and_then(|o| o.as_deref());
        paras.entry(p).or_default().new_ppr = Some(rpr::render_ppr(&pe.new_props, pstyle));
    }

    // Table-cell run edits — their own `w:tbl`/`w:tr`/`w:tc`/`w:p` locator path.
    let mut cell_targets: Vec<ResolvedCellTarget> = Vec::new();
    for c in &edits.cells {
        let Some((path, run_ord)) = bindings.resolve_cell(c.block, c.cell, c.para, c.run) else {
            continue; // non-patchable — skipped
        };
        if c.new_text.is_none() && c.new_props.is_none() {
            continue;
        }
        let new_rpr = c.new_props.as_ref().map(|props| {
            let rstyle = c.rstyle.as_ref().and_then(|o| o.as_deref());
            rpr::render_rpr(props, rstyle)
        });
        cell_targets.push(ResolvedCellTarget {
            table_ord: path.table_ord,
            row: path.row,
            cell: path.cell,
            para: path.para,
            run_ord,
            new_text: c.new_text.clone(),
            new_rpr,
        });
    }

    if targets.is_empty()
        && paras.is_empty()
        && cell_targets.is_empty()
        && rows.is_empty()
        && columns.is_empty()
    {
        return Ok(());
    }

    let patched = {
        let src = pkg.require(main_part)?;
        patch_document_xml_cols(src, &targets, &paras, &cell_targets, &rows, &columns)
    };
    pkg.set_part(main_part, patched);
    Ok(())
}
