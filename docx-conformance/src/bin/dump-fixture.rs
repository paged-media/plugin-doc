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

//! Write the memo conformance fixture to a path, so the wasm-boundary smoke
//! harness (a Node script booting the real docx-js artifact) has a real `.docx`.
//! `cargo run -p docx-conformance --bin dump-fixture -- <path>`

fn main() {
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "memo.docx".to_string());
    std::fs::write(&path, docx_conformance::memo_docx()).expect("write fixture");
    eprintln!("wrote {path}");
}
