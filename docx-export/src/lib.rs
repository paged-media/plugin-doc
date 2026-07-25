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
pub use edit::{EditSet, RunEdit};
pub use overlay::{overlay_story_content, ParagraphContentIn, RunContentIn, StoryContentIn};

use paged_ooxml::{OoxmlError, OpcPackage};
use splice::{patch_document_xml, ResolvedTarget};

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
        let Some((para_ord, run_ord)) = bindings.resolve(e.block, e.run) else {
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
            para_ord,
            run_ord,
            new_text: e.new_text.clone(),
            new_rpr,
        });
    }

    if targets.is_empty() {
        return Ok(());
    }

    let patched = {
        let src = pkg.require(main_part)?;
        patch_document_xml(src, &targets)
    };
    pkg.set_part(main_part, patched);
    Ok(())
}
