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

//! Real Word documents — the ones a human actually authored.
//!
//! Every other test in this crate builds its `.docx` in memory
//! (`docx_conformance::*_docx`), which is the right default: those
//! packages are minimal, precise and carry no binary blobs. But they are
//! also all OUR shape — five or six parts, hand-written XML, no theme,
//! no `settings.xml`, no `numbering.xml`, no `fontTable.xml`, no RSID
//! noise, no `mc:AlternateContent` fallbacks. Until 2026-08-19 this
//! parser had never seen a file Word itself wrote.
//!
//! The corpus campaign extracted 23 genuine Word documents out of the
//! Envato pack zips — 18 OOXML (including 5 `.dotx` templates), 4 true
//! binary CFB/OLE `.doc`, and one that is RTF despite its extension.
//! This lane runs the importer over all of them.
//!
//! OPT-IN, like the sibling corpus lanes (`PAGED_IDML_CORPUS`,
//! `PAGED_PSD_ORACLE`, `PAGED_ABR_CORPUS`): the assets live in the
//! private corpus checkout, which CI does not have.
//!
//! ```text
//! PAGED_DOC_CORPUS=1 cargo test -p docx-conformance --test real_word_corpus -- --ignored --nocapture
//! ```
//!
//! What it asserts is deliberately structural, not fidelity: a real
//! designed document has no "expected" lowering to compare against. It
//! must not panic, must not hang, must either import or fail with a
//! typed error, and when it imports it must produce something non-empty.
//! That is exactly the surface a hand-built fixture cannot cover.

use std::path::PathBuf;

use docx_import::import_docx;

/// Every `assets/word/*` file across the extracted packs, or `None` with
/// a printed reason when this machine has no corpus.
fn corpus_word_files() -> Option<Vec<PathBuf>> {
    let Some(switch) = std::env::var_os("PAGED_DOC_CORPUS") else {
        eprintln!(
            "SKIP doc corpus lane: PAGED_DOC_CORPUS unset \
             (set it to 1, or to a corpus root, and run with --ignored)"
        );
        return None;
    };
    let switch = switch.to_string_lossy().into_owned();
    let root = if switch == "1" || switch.is_empty() {
        // Default sibling layout: ~/paged/corpus next to
        // ~/paged/plugins/plugin-doc.
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../corpus")
    } else {
        PathBuf::from(switch)
    };
    let packs = root.join("envato/packs");
    if !packs.is_dir() {
        eprintln!(
            "SKIP doc corpus lane: {} is not a directory",
            packs.display()
        );
        return None;
    }
    let mut out = Vec::new();
    let Ok(entries) = std::fs::read_dir(&packs) else {
        eprintln!("SKIP doc corpus lane: cannot read {}", packs.display());
        return None;
    };
    for pack in entries.flatten() {
        let word = pack.path().join("assets/word");
        let Ok(files) = std::fs::read_dir(&word) else {
            continue;
        };
        for f in files.flatten() {
            let p = f.path();
            if p.is_file() {
                out.push(p);
            }
        }
    }
    out.sort();
    if out.is_empty() {
        eprintln!(
            "SKIP doc corpus lane: no assets/word/* under {} — run corpus/envato/unpack.sh",
            packs.display()
        );
        return None;
    }
    Some(out)
}

fn ext_of(p: &std::path::Path) -> String {
    p.extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default()
}

#[test]
#[ignore = "doc corpus lane: opt-in (PAGED_DOC_CORPUS=1 + the private corpus mount)"]
fn every_real_word_document_imports_or_fails_cleanly() {
    let Some(files) = corpus_word_files() else {
        return;
    };
    println!("doc corpus: {} file(s)", files.len());

    let mut ooxml_ok = 0usize;
    let mut rejected: Vec<(String, String)> = Vec::new();

    for path in &files {
        let name = path
            .file_name()
            .map(|n| n.to_string_lossy().into_owned())
            .unwrap_or_default();
        let bytes = std::fs::read(path).expect("read corpus word file");
        let ext = ext_of(path);

        match import_docx(&bytes) {
            Ok(doc) => {
                // A real designed document is never empty. This is the
                // whole assertion: the importer didn't just "succeed" by
                // returning a hollow shell.
                assert!(
                    !doc.body.is_empty(),
                    "{name}: imported with ZERO blocks — a real Word document \
                     always has content, so this is a silent parse failure"
                );
                ooxml_ok += 1;
                println!("  ok       {name:<52} {} block(s)", doc.body.len());
            }
            Err(e) => {
                // `.doc` is CFB/OLE, not OOXML — rejecting it is CORRECT.
                // Recording rather than asserting keeps the lane honest
                // about which formats it actually covers.
                rejected.push((name.clone(), format!("{e:?}")));
                println!("  rejected {name:<52} ({ext}) {e:?}");
            }
        }
    }

    println!(
        "doc corpus: {ooxml_ok} imported, {} rejected",
        rejected.len()
    );

    // The lane exists to prove the OOXML path survives real Word output.
    // If nothing imported, either the corpus is wrong or the parser is.
    assert!(
        ooxml_ok > 0,
        "not one real Word document imported — the corpus has {} files and \
         every one was rejected: {rejected:?}",
        files.len()
    );

    // Every rejection must be a legacy format. An OOXML file that the
    // importer refuses is a real defect, not a format boundary.
    for (name, err) in &rejected {
        let ext = ext_of(std::path::Path::new(name));
        assert!(
            ext == "doc",
            "{name} is .{ext} (an OOXML format) but the importer rejected it: {err}"
        );
        // ADR-007: and it must be rejected BY NAME. "Could not find EOCD"
        // is what the bare zip reader says about an RTF that someone saved
        // as .doc — it tells the user nothing about what they opened.
        assert!(
            err.contains("LegacyBinaryDoc") || err.contains("RichTextFormat"),
            "{name} was rejected with a leaked container error instead of a \
             named format: {err} — add a sniff in paged_ooxml::opc"
        );
    }
}

#[test]
#[ignore = "doc corpus lane: opt-in (PAGED_DOC_CORPUS=1 + the private corpus mount)"]
fn dotx_templates_import_like_documents() {
    let Some(files) = corpus_word_files() else {
        return;
    };
    let dotx: Vec<_> = files.iter().filter(|p| ext_of(p) == "dotx").collect();
    if dotx.is_empty() {
        eprintln!("SKIP: no .dotx in the corpus");
        return;
    }
    // A .dotx is a template: same OPC shape as .docx with a different
    // content type. Nothing in the importer should care, and this is the
    // only place that claim is tested against a real template.
    for path in dotx {
        let bytes = std::fs::read(path).expect("read dotx");
        let doc = import_docx(&bytes).unwrap_or_else(|e| {
            panic!(
                "{}: a .dotx template must import exactly like a .docx: {e:?}",
                path.display()
            )
        });
        assert!(
            !doc.body.is_empty(),
            "{}: template imported with zero blocks",
            path.display()
        );
    }
}
