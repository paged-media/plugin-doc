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

//! Write the skeleton IDML of the conformance fixtures to a directory, for
//! asking InDesign whether it is real IDML (ADR 029):
//!
//! ```sh
//! cargo run -p docx-skeleton --example skeleton_idml -- <out-dir>
//! ```

fn main() {
    let out = std::env::args()
        .nth(1)
        .expect("usage: skeleton_idml <out-dir>");
    std::fs::create_dir_all(&out).expect("out dir");
    let fixtures = [
        ("pagination", docx_conformance::pagination_docx()),
        ("continuous", docx_conformance::continuous_docx()),
    ];
    for (stem, docx) in fixtures {
        let doc = docx_import::import_docx(&docx).expect("import");
        let sk = docx_skeleton::skeleton(&doc, &format!("{stem}.docx")).expect("skeleton");
        let path = format!("{out}/{stem}.idml");
        std::fs::write(&path, &sk.idml).expect("write");
        println!(
            "{path} ({} bytes, {} stories)",
            sk.idml.len(),
            sk.section_stories.len()
        );
    }
}
