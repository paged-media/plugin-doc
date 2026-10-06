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

//! The plain-Rust engine core behind the wasm surface.
//!
//! `DocSession` holds a loaded `.docx` (its retained bytes for the preservation
//! invariant, its parsed semantic model, the lowering, and — for M2 edited
//! save-back — the retained OPC package + the native↔OOXML provenance bindings).
//! All real work lives here so it is native-testable; `lib.rs`'s
//! `#[wasm_bindgen]` layer is a pure forwarding shim.

use docx_core::DocxDocument;
use docx_export::{
    apply_edits, build_bindings, diff, overlay_story_content, DocxBindings, EditSet, RunEdit,
    StoryContentIn,
};
use docx_import::import_docx_with_package;
use docx_lower::ir::{LoweredBlock, LoweredDoc};
use docx_lower::{lower_with, LowerOptions};
use paged_ooxml::OpcPackage;
use serde::Deserialize;

/// A loaded Word document session.
pub struct DocSession {
    /// The original `.docx` bytes, retained for the preservation invariant
    /// (zero-edit save-back re-emits them verbatim).
    source: Vec<u8>,
    /// The parsed semantic model.
    model: DocxDocument,
    /// The retained OPC package — the save-back patch target (`set_part` +
    /// `write`, untouched parts re-emitted byte-identical).
    package: OpcPackage,
    /// The main document part name (e.g. `word/document.xml`).
    main_part: String,
    /// The native↔OOXML provenance map for edited save-back.
    bindings: DocxBindings,
    /// What the target engine can lay out (ADR 029: span/split columns for
    /// a mid-page column change). The lowering, the skeleton and the
    /// save-back baseline all use it, so they stay aligned.
    options: LowerOptions,
}

impl DocSession {
    /// Load a `.docx`/`.dotx` package. Returns a human-readable error string on a
    /// hard failure (unreadable container / unparseable main document).
    pub fn load(bytes: &[u8]) -> Result<DocSession, String> {
        let (model, package, main_part) =
            import_docx_with_package(bytes).map_err(|e| e.to_string())?;
        let bindings = build_bindings(&model);
        Ok(DocSession {
            source: bytes.to_vec(),
            model,
            package,
            main_part,
            bindings,
            options: LowerOptions::default(),
        })
    }

    /// The Tier-0 lowering (the IR the host-model turns into mutations).
    pub fn lowered(&self) -> LoweredDoc {
        lower_with(&self.model, self.options)
    }

    /// Lower a mid-page column change to span/split columns (`true`, the
    /// default: protocol 64) or to a page break (`false`: an engine that
    /// refused the properties). Applies to every later lowering, skeleton
    /// and save-back.
    pub fn set_mid_page_columns(&mut self, on: bool) {
        self.options.mid_page_columns = on;
    }

    /// Whether the lowering uses span/split columns anywhere (the bundle
    /// only needs to probe the engine then).
    pub fn uses_mid_page_columns(&self) -> bool {
        let ir = self.lowered();
        ir.styles.iter().any(|s| {
            s.props
                .iter()
                .any(|p| p.path == docx_lower::SPAN_COLUMN_TYPE)
        })
    }

    /// The skeleton's story ids, in story order.
    pub fn skeleton_stories(&self) -> Vec<String> {
        docx_skeleton::story_ids_with(&self.model, self.options)
    }

    /// The lowering as a JSON string (the wasm boundary form).
    pub fn lowered_json(&self) -> String {
        serde_json::to_string(&self.lowered()).unwrap_or_else(|_| "null".to_string())
    }

    /// The parsed semantic model (exposed for tests / debugging).
    pub fn model(&self) -> &DocxDocument {
        &self.model
    }

    /// ADR 029 — the page skeleton for standalone open: IDML bytes (a page,
    /// margin-box frame and empty story per Word section that starts a page;
    /// continuous sections can join the story before) plus the story ids, in
    /// order, for the grow rules and the pour.
    pub fn skeleton(&self, name: &str) -> Result<docx_skeleton::Skeleton, String> {
        docx_skeleton::skeleton_with(self.model(), name, self.options)
    }

    /// Number of top-level body blocks.
    pub fn block_count(&self) -> usize {
        self.model.body.len()
    }

    /// Zero-edit save-back: re-emit the retained source verbatim.
    pub fn save_verbatim(&self) -> Vec<u8> {
        self.source.clone()
    }

    /// M2 edited save-back: patch the given run edits into the retained package
    /// (only the changed `<w:t>`/`<w:rPr>` subtrees rewritten; every other part +
    /// untouched subtree byte-identical) and return the saved `.docx` bytes.
    /// Non-patchable edits are skipped. A cheap clone of the package keeps the
    /// session reusable for a subsequent save.
    ///
    /// DEFERRED (RFI DOC-03): the `EditSet` here is supplied by the caller
    /// (tests today). Wiring it from the LIVE editor needs a structured
    /// whole-document read door — `host.nativeDocument.readModel()` returns
    /// opaque core-native bytes this isolation-clean plugin cannot diff.
    pub fn save_edited(&self, edits: &EditSet) -> Result<(Vec<u8>, Vec<String>), String> {
        let mut package = self.package.clone();
        let skips = apply_edits(&mut package, &self.main_part, &self.bindings, edits)
            .map_err(|e| e.to_string())?;
        let bytes = package.write().map_err(|e| e.to_string())?;
        Ok((bytes, skips))
    }

    /// The object model's writes (the bundle's `run.text` and
    /// `mergeField.value`) as an [`EditSet`] over this document's lowering:
    /// a run edit sets one run's text; a merge-field edit sets the field's
    /// RESULT — its first result run takes the value, the others are emptied —
    /// and never touches the field code, so Word can still merge. Body
    /// paragraphs only: a table cell's runs are refused (save-back of cell
    /// runs goes through `cells`, which the object model does not address yet).
    pub fn object_edits(&self, edits: &ObjectEdits) -> Result<EditSet, String> {
        let ir = self.lowered();
        let paragraph = |block: usize| match ir.story.blocks.get(block) {
            Some(LoweredBlock::Paragraph(p)) => Ok(p),
            Some(LoweredBlock::Table(_)) => Err(format!(
                "block {block} is a table: its cells are not writable here yet"
            )),
            None => Err(format!("no block {block}")),
        };
        let mut set = EditSet::default();
        for e in &edits.runs {
            let p = paragraph(e.block)?;
            if e.run >= p.runs.len() {
                return Err(format!("block {} has no run {}", e.block, e.run));
            }
            set.runs.push(RunEdit::text(e.block, e.run, e.text.clone()));
        }
        for e in &edits.merge_fields {
            let p = paragraph(e.block)?;
            let runs: Vec<usize> = p
                .runs
                .iter()
                .enumerate()
                .filter(|(_, r)| {
                    r.merge_field
                        .as_ref()
                        .is_some_and(|m| m.occurrence == e.occurrence)
                })
                .map(|(i, _)| i)
                .collect();
            let Some((&first, rest)) = runs.split_first() else {
                return Err(format!(
                    "block {} has no merge field #{}",
                    e.block, e.occurrence
                ));
            };
            set.runs
                .push(RunEdit::text(e.block, first, e.value.clone()));
            for &r in rest {
                set.runs.push(RunEdit::text(e.block, r, String::new()));
            }
        }
        Ok(set)
    }

    /// [`Self::object_edits`], saved: the edited `.docx` bytes plus the
    /// patcher's refusal ledger.
    pub fn save_object_edits(&self, edits: &ObjectEdits) -> Result<(Vec<u8>, Vec<String>), String> {
        self.save_edited(&self.object_edits(edits)?)
    }

    /// M2 edited save-back, driven by the DOC-03 structured read: overlay the
    /// host's read-back `StoryContent` onto the import baseline lowering, diff to
    /// an `EditSet`, and save. This is the LIVE end-to-end path (the bundle reads
    /// `host.document.storyContent(storyId)` and forwards it here) — the seam that
    /// makes edited save-back run without a hand-authored `EditSet`.
    pub fn save_edited_from_content(
        &self,
        content: &StoryContentIn,
    ) -> Result<(Vec<u8>, Vec<String>), String> {
        let baseline = self.lowered();
        let edited = overlay_story_content(&baseline, content);
        let edits = diff(&baseline, &edited, &self.bindings);
        self.save_edited(&edits)
    }
}

/// The object model's writes, addressed in lowered story coordinates (the
/// bundle's `run` and `mergeField` objects).
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ObjectEdits {
    #[serde(default)]
    pub runs: Vec<RunTextEdit>,
    #[serde(default)]
    pub merge_fields: Vec<MergeFieldEdit>,
}

/// Set one body run's text.
#[derive(Debug, Clone, Deserialize)]
pub struct RunTextEdit {
    pub block: usize,
    pub run: usize,
    pub text: String,
}

/// Set a merge field's result: the `occurrence`-th merge field of a body
/// paragraph block.
#[derive(Debug, Clone, Deserialize)]
pub struct MergeFieldEdit {
    pub block: usize,
    pub occurrence: u32,
    pub value: String,
}
