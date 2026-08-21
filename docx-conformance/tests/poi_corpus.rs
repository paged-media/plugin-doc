/*
 * This file is part of paged (https://paged.media), the commercial editor
 * for the paged IDML engine.
 *
 * paged is free software: you may redistribute it and/or modify it under the
 * terms of the GNU Affero General Public License, version 3, as published by
 * the Free Software Foundation, OR under the Paged Media Enterprise License
 * (PMEL), a commercial license available from And The Next GmbH. Full
 * copyright and license information is available in LICENSE.md, distributed
 * with this source code.
 *
 *  @copyright  Copyright (c) And The Next GmbH
 *  @license    AGPL-3.0-only OR Paged Media Enterprise License (PMEL)
 */

//! Apache POI's word-processing corpus — twenty years of bug reports.
//!
//! `real_word_corpus.rs` runs the importer over the Envato packs and
//! asserts that every OOXML file IMPORTS, which is right for those: they
//! are well-formed designer output, so a failure is a real defect.
//!
//! **That assertion would be wrong here.** POI's corpus accumulated from
//! bug reports, and a good number of its files are deliberately
//! malformed because they once crashed something. Asserting they all
//! import would be asserting that upstream's regression suite contains
//! no regressions. The honest properties for THIS corpus are weaker and
//! more useful:
//!
//!   * never panic, on any input
//!   * never hang
//!   * fail with a TYPED error rather than a leaked one
//!
//! The import rate is REPORTED, not gated. A drop is something to read,
//! not an automatic failure.
//!
//! The 159 legacy `.doc` files are the reason this set is worth its
//! weight. `paged-ooxml` sniffs container magic and returns a NAMED
//! error — `LegacyBinaryDoc` for CFB/OLE, `RichTextFormat` for
//! RTF-saved-as-.doc — so a user learns what they opened instead of
//! reading "Could not find EOCD". Until now that machinery faced FOUR
//! files. It now faces 159.
//!
//! `docx/poi-converted/` inverts that bargain. Those are the same
//! legacy `.doc` files re-saved as OOXML by desktop Word 16
//! (`corpus/harness/convert-office.sh`, a maintainer tool; CI consumes
//! the committed output). The INPUT was a bug-report corpus, but the
//! OUTPUT is Word's own writer — so the bar there is the
//! `real_word_corpus.rs` bar: **every one must import**, and a refusal
//! is a real defect rather than a curiosity. What they add is twenty
//! years of content odd enough to file a bug about, expressed in the
//! OOXML of the producer this importer exists to read.
//!
//! 128 of the 159 converted, and 128/128 import as of 2026-08-21, so
//! that starts green and is a gate rather than a ratchet. The 31 Word
//! itself would not re-save are the population you would expect: 11
//! `clusterfuzz-testcase-*`, three encrypted, six in the Word 2.0/6.0/95
//! formats modern Word blocks by policy, and eleven POI bug-report files
//! broken enough that Word declines them too. A file Word cannot read is
//! not evidence about this importer either way.
//!
//! Licence: Apache-2.0. These and the sibling `xlsx/poi` set are the
//! first REDISTRIBUTABLE fixtures this project carries — everything else
//! is Envato, which grants use and not redistribution. See
//! `corpus/docx/poi/PROVENANCE.md`. The conversions inherit that licence
//! from their inputs.
//!
//! OPT-IN — the assets live in the private corpus checkout:
//!
//! ```text
//! PAGED_DOCX_CORPUS=1 cargo test -p docx-conformance --test poi_corpus -- --ignored --nocapture
//! ```

use std::path::{Path, PathBuf};

use docx_import::import_docx;

const EXTS: &[&str] = &["docx", "dotx", "docm", "doc"];

/// The corpus `docx/` root, or `None` with a printed reason.
fn docx_root() -> Option<PathBuf> {
    let Some(switch) = std::env::var_os("PAGED_DOCX_CORPUS") else {
        eprintln!(
            "SKIP poi docx lane: PAGED_DOCX_CORPUS unset \
             (set it to 1, or to a corpus root, and run with --ignored)"
        );
        return None;
    };
    let switch = switch.to_string_lossy().into_owned();
    let root = if switch == "1" || switch.is_empty() {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../corpus")
    } else {
        PathBuf::from(switch)
    };
    let dir = root.join("docx");
    if !dir.is_dir() {
        eprintln!("SKIP poi docx lane: {} not readable", dir.display());
        return None;
    }
    Some(dir)
}

/// Recursive walk — `docx-conformance` has no `walkdir`, and pulling a
/// dependency in for one test lane is not worth it.
fn walk(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for e in entries.flatten() {
        let p = e.path();
        if p.is_symlink() {
            continue;
        }
        if p.is_dir() {
            // Skip dot-dirs; `Document fonts/.cache` and friends are
            // machine-local.
            if !p
                .file_name()
                .is_some_and(|n| n.to_string_lossy().starts_with('.'))
            {
                walk(&p, out);
            }
        } else if p.is_file()
            && p.extension()
                .is_some_and(|e| EXTS.contains(&e.to_string_lossy().to_lowercase().as_str()))
        {
            out.push(p);
        }
    }
}

/// Every document under one named source directory, e.g. `poi-converted`.
///
/// The walk is recursive on purpose. The non-recursive `read_dir` this
/// replaced could only ever see `docx/poi`, which is why
/// `docx/poi-converted/` sat in the corpus unwired to any assertion.
fn source_files(source: &str) -> Option<Vec<PathBuf>> {
    let root = docx_root()?;
    let dir = root.join(source);
    let mut out = Vec::new();
    walk(&dir, &mut out);
    out.sort();
    if out.is_empty() {
        eprintln!("SKIP: no documents under {}", dir.display());
        return None;
    }
    Some(out)
}

fn poi_files() -> Option<Vec<PathBuf>> {
    let root = docx_root()?;
    let dir = root.join("poi");
    let Ok(entries) = std::fs::read_dir(&dir) else {
        eprintln!("SKIP poi docx lane: {} not readable", dir.display());
        return None;
    };
    let mut out: Vec<PathBuf> = entries
        .flatten()
        .map(|e| e.path())
        .filter(|p| {
            p.is_file()
                && p.extension().is_some_and(|e| {
                    matches!(
                        e.to_string_lossy().to_lowercase().as_str(),
                        "docx" | "dotx" | "docm" | "doc"
                    )
                })
        })
        .collect();
    out.sort();
    if out.is_empty() {
        eprintln!("SKIP poi docx lane: no documents under {}", dir.display());
        return None;
    }
    Some(out)
}

fn ext_of(p: &Path) -> String {
    p.extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default()
}

#[test]
#[ignore = "poi docx lane: opt-in (PAGED_DOCX_CORPUS=1 + the private corpus mount)"]
fn no_document_in_the_poi_corpus_panics_the_importer() {
    let Some(files) = poi_files() else {
        return;
    };
    println!("poi docx corpus: {} file(s)", files.len());

    let mut by_ext: std::collections::BTreeMap<String, (usize, usize)> = Default::default();
    for path in &files {
        let bytes = std::fs::read(path).expect("read poi fixture");
        // Err is a RESULT here, not a failure. A panic is the failure,
        // and cargo names the file that caused it.
        match import_docx(&bytes) {
            Ok(_) => by_ext.entry(ext_of(path)).or_default().0 += 1,
            Err(_) => by_ext.entry(ext_of(path)).or_default().1 += 1,
        }
    }

    let opened: usize = by_ext.values().map(|(o, _)| o).sum();
    println!("  imported {opened}, refused {}", files.len() - opened);
    for (ext, (ok, err)) in &by_ext {
        println!("    .{ext:<5} {ok:>4} imported  {err:>4} refused");
    }

    // The one hard assertion: a producer population this wide must
    // produce SOMETHING readable. Zero would mean the importer only ever
    // handled the Envato packs' narrow shape.
    assert!(
        opened > 0,
        "not one of {} POI documents imported — the importer has only ever \
         seen hand-built fixtures and Envato templates, so this means it \
         cannot read the wider world at all",
        files.len()
    );
}

#[test]
#[ignore = "poi docx lane: opt-in (PAGED_DOCX_CORPUS=1 + the private corpus mount)"]
fn every_legacy_doc_is_refused_by_name() {
    let Some(files) = poi_files() else {
        return;
    };
    let legacy: Vec<_> = files.iter().filter(|p| ext_of(p) == "doc").collect();
    if legacy.is_empty() {
        eprintln!("SKIP: no .doc in the corpus");
        return;
    }

    // ADR-007, at scale. Some legacy .doc files EMBED a complete OPC
    // package, so a zip reader finds it by EOCD scan, half-imports, and
    // dies later with a nonsense XML error about themeManager.xml. The
    // container sniff exists to stop that, and to say WHICH format the
    // user actually opened.
    let mut opened = Vec::new();
    let mut unnamed = Vec::new();
    for path in &legacy {
        let name = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .into_owned();
        let bytes = std::fs::read(path).expect("read poi doc");
        match import_docx(&bytes) {
            Ok(_) => opened.push(name),
            Err(e) => {
                let msg = format!("{e:?}");
                // RTF-saved-as-.doc is a real population here too, and
                // has its own named error.
                if !(msg.contains("LegacyBinaryDoc") || msg.contains("RichTextFormat")) {
                    unnamed.push(format!("{name}: {msg}"));
                }
            }
        }
    }

    println!("legacy .doc: {} file(s)", legacy.len());
    assert!(
        opened.is_empty(),
        "{} legacy .doc file(s) IMPORTED as OOXML: {:?} — a half-read .doc \
         is worse than a refused one, because the caller gets a document \
         that is not the one they opened",
        opened.len(),
        &opened[..opened.len().min(5)]
    );
    assert!(
        unnamed.is_empty(),
        "{} legacy .doc file(s) were refused with a LEAKED container error \
         instead of a named format — the whole point of the sniff is that a \
         user learns what they opened: {:?}",
        unnamed.len(),
        &unnamed[..unnamed.len().min(5)]
    );
}

#[test]
#[ignore = "poi docx lane: opt-in (PAGED_DOCX_CORPUS=1 + the private corpus mount)"]
fn every_converted_document_imports() {
    let Some(files) = source_files("poi-converted") else {
        return;
    };

    // The inverse of the bargain above. Whatever the input was, desktop
    // Word 16 wrote these — so a refusal here is this importer failing
    // on Word's own output, which is a defect, not the "upstream ships
    // deliberately-broken fixtures" story that makes the POI rate a
    // report rather than a gate.
    //
    // Same bar and same shape as `real_xlsx_corpus.rs::
    // every_converted_workbook_opens`, which gates the Excel half of the
    // same harness.
    let mut failed = Vec::new();
    for path in &files {
        let bytes = std::fs::read(path).expect("read converted fixture");
        if let Err(e) = import_docx(&bytes) {
            failed.push(format!(
                "{}: {e:?}",
                path.file_name().unwrap_or_default().to_string_lossy()
            ));
        }
    }

    println!(
        "poi-converted: {} file(s), {} imported",
        files.len(),
        files.len() - failed.len()
    );
    assert!(
        failed.is_empty(),
        "{} of {} Word-converted document(s) failed to import — Word 16 \
         wrote every one of these, so each is a real importer defect:\n  {}",
        failed.len(),
        files.len(),
        failed.join("\n  ")
    );
}
