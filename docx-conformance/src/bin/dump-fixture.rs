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

//! Write a conformance fixture to a path, so the wasm-boundary smoke harness
//! (a Node script booting the real docx-js artifact) has a real `.docx`.
//! `cargo run -p docx-conformance --bin dump-fixture -- <path> [memo|tier1|one]`

fn main() {
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "memo.docx".into());
    let which = std::env::args().nth(2).unwrap_or_else(|| "memo".into());
    let bytes = match which.as_str() {
        "tier1" => docx_conformance::tier1_docx(),
        "list" => docx_conformance::list_docx(),
        "image" => docx_conformance::image_docx(),
        "hyperlink" => docx_conformance::hyperlink_docx(),
        "fieldlink" => docx_conformance::field_hyperlink_docx(),
        "footnote" => docx_conformance::footnote_docx(),
        "table" => docx_conformance::table_docx(),
        "one" => docx_conformance::one_paragraph_docx(),
        _ => docx_conformance::memo_docx(),
    };
    std::fs::write(&path, bytes).expect("write fixture");
    eprintln!("wrote {which} -> {path}");
}
