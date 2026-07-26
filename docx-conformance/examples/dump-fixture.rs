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

//! Write one of the in-memory fixture packages out as a real `.docx` file.
//!
//! The conformance fixtures are built in memory precisely so this repo carries
//! no binary blobs — but a HOST-integration test (the editor's DTP journey)
//! needs an actual file to hand to a file input. This dumps one, so that
//! fixture stays generated from the same source of truth rather than becoming
//! an opaque committed blob whose provenance nobody can check.
//!
//!   cargo run -p docx-conformance --example dump-fixture -- memo /tmp/memo.docx
//!
//! Names: memo | tier1 | list | table | image | hyperlink

use std::path::PathBuf;

fn main() {
    let mut args = std::env::args().skip(1);
    let name = args.next().unwrap_or_else(|| usage("missing <fixture>"));
    let out = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| usage("missing <out.docx>"));

    let bytes = match name.as_str() {
        "memo" => docx_conformance::memo_docx(),
        "tier1" => docx_conformance::tier1_docx(),
        "list" => docx_conformance::list_docx(),
        "table" => docx_conformance::table_docx(),
        "image" => docx_conformance::image_docx(),
        "hyperlink" => docx_conformance::hyperlink_docx(),
        other => usage(&format!("unknown fixture {other:?}")),
    };

    std::fs::write(&out, &bytes).unwrap_or_else(|e| panic!("write {}: {e}", out.display()));
    println!("wrote {} ({} bytes)", out.display(), bytes.len());
}

fn usage(why: &str) -> ! {
    eprintln!("dump-fixture: {why}");
    eprintln!("usage: dump-fixture <memo|tier1|list|table|image|hyperlink> <out.docx>");
    std::process::exit(2)
}
