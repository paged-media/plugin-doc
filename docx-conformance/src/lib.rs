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

//! TEST-ONLY fixture builders: assemble minimal, real `.docx` OPC packages in
//! memory so the conformance suite carries no binary blobs. Each builder zips
//! the required parts (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`,
//! and optionally `word/styles.xml`) plus an `unknown/note.txt` part used to
//! prove the preservation invariant (unknown parts survive a round-trip).

use std::io::Write;

use zip::write::SimpleFileOptions;

/// Zip a set of `(name, bytes)` parts into an OPC package, in the given order.
pub fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut cursor = std::io::Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut cursor);
        let opts =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, bytes) in parts {
            zip.start_file(*name, opts).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    cursor.into_inner()
}

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="48"/></w:rPr>
  </w:style>
</w:styles>"#;

/// thoughts ADR 029 — pagination ground truth: what Word itself does with
/// sections, margins and page breaks, to measure standalone open against
/// (`scripts/word-pagination-probe.sh` exports it from Word as PDF).
///
/// Every paragraph is one line on an exact 12 pt grid (Inter 10 pt,
/// `w:spacing w:line="240" w:lineRule="exact"`, no space before/after), so a
/// page holds `body height / 12` lines and a misplaced break shows as a
/// numbered paragraph on the wrong page. (Word cannot load the installed
/// Inter — a variable font — and silently lays these in Calibri; the
/// exact grid makes that harmless here, see `line_spacing_docx`.)
///
/// - Section 1: US Letter, 1 in margins → 648 pt body = 54 lines. 120
///   paragraphs (`S1 P001`…), so three pages by line count. `S1 P054`, the
///   last line of page 1, carries `w:keepNext`: Word must move it to page 2.
/// - Section 2 (`nextPage`): A5 landscape, 0.5 in margins → ~347 pt body =
///   28 lines. 40 paragraphs (`S2 P001`…).
pub fn pagination_docx() -> Vec<u8> {
    fn para(text: &str, keep_next: bool) -> String {
        format!(
            r#"<w:p><w:pPr>{keep}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#,
            keep = if keep_next { "<w:keepNext/>" } else { "" },
        )
    }
    let mut body = String::new();
    for n in 1..=120 {
        let mut p = para(&format!("S1 P{n:03} of the first section."), n == 54);
        if n == 120 {
            // The section break rides the section's LAST paragraph.
            p = p.replace(
                "<w:pPr>",
                r#"<w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#,
            );
        }
        body.push_str(&p);
    }
    for n in 1..=40 {
        body.push_str(&para(&format!("S2 P{n:03} of the second section."), false));
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="8391" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// One case of [`line_spacing_docx`]: a section of one-line paragraphs in
/// `font` at `half_pts`, all carrying `w:spacing w:line=… w:lineRule=…`.
#[derive(Debug, Clone, Copy)]
pub struct LineSpacingCase {
    /// The paragraph label prefix (`"L1"` → `"L1 P001 …"`).
    pub label: &'static str,
    pub font: &'static str,
    /// `w:sz` (half-points).
    pub half_pts: u32,
    /// `w:spacing/@w:line`.
    pub line: i32,
    /// `w:spacing/@w:lineRule` (`"auto"`, `"exact"`, `"atLeast"`).
    pub rule: &'static str,
}

/// The cases of [`line_spacing_docx`], one section (and so one measured page)
/// each. `fixtures/line-spacing.word.json` records what Word made of them.
pub const LINE_SPACING_CASES: &[LineSpacingCase] = &[
    LineSpacingCase {
        label: "L1",
        font: "Calibri",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L2",
        font: "Calibri",
        half_pts: 20,
        line: 276,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L3",
        font: "Calibri",
        half_pts: 20,
        line: 360,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L4",
        font: "Calibri",
        half_pts: 20,
        line: 480,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L5",
        font: "Calibri",
        half_pts: 20,
        line: 240,
        rule: "exact",
    },
    LineSpacingCase {
        label: "L6",
        font: "Calibri",
        half_pts: 20,
        line: 240,
        rule: "atLeast",
    },
    LineSpacingCase {
        label: "L7",
        font: "Calibri",
        half_pts: 20,
        line: 120,
        rule: "atLeast",
    },
    // Cross-checks: the single-line height must follow the FACE (faces whose
    // hhea, typo and win metrics disagree in different directions) and
    // scale with the SIZE, so the rule is not fitted to one font at one size.
    LineSpacingCase {
        label: "L8",
        font: "Calibri",
        half_pts: 28,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L9",
        font: "Aptos",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L10",
        font: "Aptos",
        half_pts: 20,
        line: 240,
        rule: "atLeast",
    },
    LineSpacingCase {
        label: "L11",
        font: "Arial",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L12",
        font: "Arial",
        half_pts: 20,
        line: 360,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L13",
        font: "Times New Roman",
        half_pts: 24,
        line: 240,
        rule: "auto",
    },
    // atLeast ABOVE the face's single line: the value wins.
    LineSpacingCase {
        label: "L14",
        font: "Arial",
        half_pts: 20,
        line: 240,
        rule: "atLeast",
    },
    // Single (auto 240) at 10 pt for every face in docx-lower's line-height
    // table, so each entry there is a number Word produced, not a guess.
    LineSpacingCase {
        label: "L15",
        font: "Calibri Light",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L16",
        font: "Cambria",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L17",
        font: "Candara",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L18",
        font: "Consolas",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L19",
        font: "Constantia",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L20",
        font: "Corbel",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L21",
        font: "Georgia",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L22",
        font: "Verdana",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L23",
        font: "Tahoma",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L24",
        font: "Trebuchet MS",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L25",
        font: "Courier New",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L26",
        font: "Garamond",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L27",
        font: "Century Gothic",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L28",
        font: "Book Antiqua",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L29",
        font: "Palatino Linotype",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L30",
        font: "Helvetica",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L31",
        font: "Arial Narrow",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L32",
        font: "Gill Sans MT",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L33",
        font: "Franklin Gothic Book",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L34",
        font: "Lucida Sans Unicode",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L35",
        font: "Comic Sans MS",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L36",
        font: "Inter",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    // The faces Word sets list bullets in: a bullet makes its line as tall
    // as ITS font's line.
    LineSpacingCase {
        label: "L37",
        font: "Symbol",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
    LineSpacingCase {
        label: "L38",
        font: "Wingdings",
        half_pts: 20,
        line: 240,
        rule: "auto",
    },
];

/// ADR 029 — line-spacing ground truth: what Word does with
/// `w:spacing/@w:line` under each `@w:lineRule` (`scripts/word-line-spacing-probe.sh`
/// exports it from Word as PDF and measures it).
///
/// One US-Letter section (1 in margins → 648 pt body) per
/// [`LINE_SPACING_CASES`] entry, each holding 80 one-line paragraphs with no
/// space before/after — more than any case fits on a page, so page 1 of every
/// section is full and its lines-per-page and baseline pitch reveal Word's
/// line height for that case.
pub fn line_spacing_docx() -> Vec<u8> {
    const SECT: &str = r#"<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#;
    let mut body = String::new();
    for (ci, c) in LINE_SPACING_CASES.iter().enumerate() {
        let last_case = ci + 1 == LINE_SPACING_CASES.len();
        for n in 1..=80 {
            // The section break rides each section's LAST paragraph, except
            // the document's last section, which is the body-level sectPr.
            let sect = if n == 80 && !last_case { SECT } else { "" };
            body.push_str(&format!(
                r#"<w:p><w:pPr>{sect}<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{rule}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{label} P{n:03} line {line} {rule}.</w:t></w:r></w:p>"#,
                line = c.line,
                rule = c.rule,
                font = c.font,
                sz = c.half_pts,
                label = c.label,
            ));
        }
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}{SECT}</w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// ADR 029 — a table row taller than the page. `before`, then a 1 x 1
/// table whose cell holds `TALL_ROW_PARAGRAPHS` one-line paragraphs (about
/// 2,400 pt of Times New Roman 12 on a 648 pt body), then `after`; and a
/// small 1 x 1 table, which stays a table. Word splits the tall row across
/// pages; a native table row never splits.
pub fn tall_row_docx() -> Vec<u8> {
    const RPR: &str = r#"<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>"#;
    let para = |text: &str| {
        format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#
        )
    };
    let table = |cell: &str| {
        format!(
            r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl>"#
        )
    };
    let tall: String = (1..=TALL_ROW_PARAGRAPHS)
        .map(|n| para(&format!("Row text {n:03}")))
        .collect();
    let body = format!(
        "{}{}{}{}{}",
        para("before"),
        table(&tall),
        para("after"),
        table(&para("small table")),
        para("end"),
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// The paragraphs in [`tall_row_docx`]'s tall cell.
pub const TALL_ROW_PARAGRAPHS: usize = 180;

/// ADR 029 — a paragraph whose runs differ in size: Times New Roman 12 pt
/// text, then a run at 10 pt, then 12 pt again; and a paragraph all at
/// 12 pt. Word sizes each line by the fonts on it, so the first paragraph's
/// runs must carry their own leadings.
pub fn mixed_sizes_docx() -> Vec<u8> {
    let run = |half_pts: u32, text: &str| {
        format!(
            r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="{half_pts}"/><w:szCs w:val="{half_pts}"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
        )
    };
    let ppr =
        r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>"#;
    let body = format!(
        "<w:p>{ppr}{}{}{}{}</w:p><w:p>{ppr}{}{}</w:p>",
        run(24, "Twelve point text, "),
        run(20, "[a citation in ten point]"),
        run(24, " "),
        run(24, "and twelve again."),
        run(24, "All of this paragraph "),
        run(24, "is twelve point."),
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// One [`paragraph_spacing_docx`] pair: `after` on the first paragraph,
/// `before` on the second (twips), and whether each asks for
/// `w:contextualSpacing`.
pub struct SpacingPair {
    pub label: &'static str,
    pub after: u32,
    pub before: u32,
    pub contextual: (bool, bool),
}

/// The pairs [`paragraph_spacing_docx`] measures, in document order.
pub const SPACING_PAIRS: &[SpacingPair] = &[
    SpacingPair {
        label: "S01",
        after: 0,
        before: 0,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S02",
        after: 120,
        before: 0,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S03",
        after: 0,
        before: 120,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S04",
        after: 120,
        before: 120,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S05",
        after: 240,
        before: 120,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S06",
        after: 120,
        before: 240,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S07",
        after: 200,
        before: 480,
        contextual: (false, false),
    },
    SpacingPair {
        label: "S08",
        after: 120,
        before: 120,
        contextual: (true, true),
    },
    SpacingPair {
        label: "S09",
        after: 120,
        before: 120,
        contextual: (true, false),
    },
    SpacingPair {
        label: "S10",
        after: 120,
        before: 120,
        contextual: (false, true),
    },
];

/// ADR 029 — paragraph-spacing ground truth: how much room Word leaves
/// between a paragraph with space AFTER and one with space BEFORE
/// (`scripts/word-paragraph-spacing-probe.sh` has Word export it as PDF and
/// measures the baseline distances).
///
/// Times New Roman 12 pt, single spacing, US Letter with 1 in margins. Each
/// [`SPACING_PAIRS`] entry is two one-line paragraphs, `<label>a` and
/// `<label>b`, followed by a `sep` paragraph with no spacing. Then a page
/// break and `TOP`, a paragraph with 24 pt space before that opens page 2
/// (does the space survive at the top of a page?), and `TOP2`, the same on a
/// page it reaches by `w:pageBreakBefore`.
///
/// `compat` is `settings.xml`'s `compatibilityMode` (`None` writes no
/// settings part): Word's layout rules differ by mode.
pub fn paragraph_spacing_docx(compat: Option<u32>) -> Vec<u8> {
    const RPR: &str = r#"<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>"#;
    let para = |ppr: &str, before: u32, after: u32, ctx: bool, inner: &str| {
        let ctx = if ctx { "<w:contextualSpacing/>" } else { "" };
        format!(
            r#"<w:p><w:pPr>{ppr}<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/>{ctx}</w:pPr>{inner}</w:p>"#
        )
    };
    let run = |text: &str| format!(r#"<w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r>"#);
    let mut body = String::new();
    body.push_str(&para("", 0, 0, false, &run("first line")));
    for c in SPACING_PAIRS {
        body.push_str(&para(
            "",
            0,
            c.after,
            c.contextual.0,
            &run(&format!("{}a after {}", c.label, c.after)),
        ));
        body.push_str(&para(
            "",
            c.before,
            0,
            c.contextual.1,
            &run(&format!("{}b before {}", c.label, c.before)),
        ));
        body.push_str(&para("", 0, 0, false, &run("sep")));
    }
    // An empty paragraph holding only a page break, then TOP.
    body.push_str(&para(
        "",
        0,
        0,
        false,
        &format!(r#"<w:r>{RPR}<w:br w:type="page"/></w:r>"#),
    ));
    body.push_str(&para("", 480, 0, false, &run("TOP before 480")));
    body.push_str(&para("", 0, 0, false, &run("TOPnext")));
    body.push_str(&para(
        "<w:pageBreakBefore/>",
        480,
        0,
        false,
        &run("TOP2 before 480"),
    ));
    body.push_str(&para("", 0, 0, false, &run("TOP2next")));
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>"#
    );
    let Some(mode) = compat else {
        return zip_parts(&[
            ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
            ("word/document.xml", document.as_bytes()),
            ("word/styles.xml", STYLES.as_bytes()),
        ]);
    };
    let content_types = CONTENT_TYPES.replace(
        "</Types>",
        r#"<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"#,
    );
    let rels = DOC_RELS.replace(
        "</Relationships>",
        r#"<Relationship Id="rSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#,
    );
    let settings = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
        ("word/settings.xml", settings.as_bytes()),
    ])
}

/// One page of [`breaks_docx`]: 5 in × 2.667 in, 0.5 in margins → a
/// 288 pt × 120 pt body = exactly ten 12 pt lines (two 126 pt columns in a
/// two-column section).
const BREAKS_PAGE: &str = r#"<w:pgSz w:w="7200" w:h="3840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>"#;

/// `w:sectPr` for a [`breaks_docx`] section starting as `kind`, in `cols`
/// columns.
fn breaks_sect(kind: &str, cols: u32) -> String {
    format!(
        r#"<w:sectPr><w:type w:val="{kind}"/>{BREAKS_PAGE}<w:cols w:num="{cols}" w:space="720"/></w:sectPr>"#
    )
}

/// The run properties every [`breaks_docx`] run carries.
const BREAKS_RPR: &str =
    r#"<w:rPr><w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/><w:sz w:val="20"/></w:rPr>"#;

/// A [`breaks_docx`] paragraph: `ppr` (extra `w:pPr` children that precede
/// `w:spacing` in schema order, e.g. `<w:pageBreakBefore/>`), the inner run
/// XML, and an optional section break (`w:sectPr`, which follows `w:spacing`).
fn breaks_para(ppr: &str, inner: &str, sect: &str) -> String {
    format!(
        r#"<w:p><w:pPr>{ppr}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>{sect}</w:pPr>{inner}</w:p>"#
    )
}

/// A plain text run.
fn breaks_run(text: &str) -> String {
    format!(r#"<w:r>{BREAKS_RPR}<w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

/// ADR 028/029 — break ground truth: what Word does with every way a `.docx`
/// says "start over" (`scripts/word-breaks-probe.sh` has Word export it as
/// PDF; `fixtures/breaks.word.json` records the page map).
///
/// Every page is the same small size with a ten-line body (Inter 10 pt on an
/// exact 12 pt grid, as [`pagination_docx`]), and every paragraph is one
/// line whose first word is its label, so the page map reads straight off the
/// PDF. The sections:
///
/// - **A** (one column): `A05` has `w:pageBreakBefore` mid-page; `A15` has it
///   too but is the 11th line after `A05`, so it already opens a page; an
///   EMPTY paragraph holding only `<w:br w:type="page"/>` follows `A17`; `A19`
///   ENDS with a page break; `A21a`/`A21b` are one paragraph (one run) with a
///   page break BETWEEN them.
/// - **B** (`nextPage`, two columns): after `B03` an empty paragraph holds
///   only `<w:br w:type="column"/>`; `B08a`/`B08b` are one paragraph with a
///   column break between them.
/// - **C** (`oddPage`) is planned to start where the next page is already odd,
///   **D** (`oddPage`) where it is even (Word adds a blank page), **E**
///   (`evenPage`) where it is already even, **F** (`evenPage`) where it is odd
///   (a blank page again). Whether the plan held is Word's to say: the
///   fixture records what Word did.
pub fn breaks_docx() -> Vec<u8> {
    let page_br = format!(r#"<w:r>{BREAKS_RPR}<w:br w:type="page"/></w:r>"#);
    let col_br = format!(r#"<w:r>{BREAKS_RPR}<w:br w:type="column"/></w:r>"#);
    let pbb = "<w:pageBreakBefore/>";
    let mut body = String::new();
    let plain = |label: &str| breaks_para("", &breaks_run(label), "");

    // Section A.
    for n in 1..=4 {
        body.push_str(&plain(&format!("A{n:02} text")));
    }
    body.push_str(&breaks_para(pbb, &breaks_run("A05 pageBreakBefore"), ""));
    for n in 6..=14 {
        body.push_str(&plain(&format!("A{n:02} text")));
    }
    body.push_str(&breaks_para(
        pbb,
        &breaks_run("A15 pageBreakBefore top"),
        "",
    ));
    body.push_str(&plain("A16 text"));
    body.push_str(&plain("A17 text"));
    body.push_str(&breaks_para("", &page_br, ""));
    body.push_str(&plain("A18 after empty break"));
    body.push_str(&breaks_para(
        "",
        &format!("{}{page_br}", breaks_run("A19 ends in break")),
        "",
    ));
    body.push_str(&plain("A20 after end break"));
    body.push_str(&breaks_para(
        "",
        &format!(
            r#"<w:r>{BREAKS_RPR}<w:t xml:space="preserve">A21a before</w:t><w:br w:type="page"/><w:t xml:space="preserve">A21b after mid break</w:t></w:r>"#
        ),
        "",
    ));
    body.push_str(&plain("A22 text"));
    body.push_str(&breaks_para(
        "",
        &breaks_run("A23 last of A"),
        &breaks_sect("nextPage", 1),
    ));

    // Section B, two columns.
    for n in 1..=3 {
        body.push_str(&plain(&format!("B{n:02} col")));
    }
    body.push_str(&breaks_para("", &col_br, ""));
    for n in 4..=7 {
        body.push_str(&plain(&format!("B{n:02} col")));
    }
    body.push_str(&breaks_para(
        "",
        &format!(
            r#"<w:r>{BREAKS_RPR}<w:t xml:space="preserve">B08a col</w:t><w:br w:type="column"/><w:t xml:space="preserve">B08b col</w:t></w:r>"#
        ),
        "",
    ));
    body.push_str(&breaks_para(
        "",
        &breaks_run("B09 last of B"),
        &breaks_sect("nextPage", 2),
    ));

    // Sections C–F: three lines each; the section kind is on its own sectPr,
    // which rides its LAST paragraph (F's is the body-level sectPr).
    for (label, kind) in [
        ("C", "oddPage"),
        ("D", "oddPage"),
        ("E", "evenPage"),
        ("F", "evenPage"),
    ] {
        for n in 1..=3 {
            let text = format!("{label}{n:02} {kind}");
            let sect = if n == 3 && label != "F" {
                breaks_sect(kind, 1)
            } else {
                String::new()
            };
            body.push_str(&breaks_para("", &breaks_run(&text), &sect));
        }
    }
    let last = breaks_sect("evenPage", 1);
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}{last}</w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// One page of [`line_breaks_docx`]: 5 in × 4⅓ in, 0.5 in margins → a
/// 288 pt × 240 pt body = twenty 12 pt lines.
const LINE_BREAKS_PAGE: &str = r#"<w:pgSz w:w="7200" w:h="6240"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>"#;

/// The cases of [`line_breaks_docx`], in document order: a label, the
/// paragraph's exact line pitch in twips, and its run XML (`""` for a blank
/// line). `{R}` stands for the fixture's run properties.
pub const LINE_BREAK_CASES: &[(&str, u32, &str)] = &[
    ("L01", 240, r#"<w:r>{R}<w:t>L01 one line</w:t></w:r>"#),
    // A plain `<w:br/>` mid-paragraph, inside one run.
    (
        "L02",
        240,
        r#"<w:r>{R}<w:t xml:space="preserve">L02a before</w:t><w:br/><w:t xml:space="preserve">L02b after</w:t></w:r>"#,
    ),
    // Two in a row: an empty line between.
    (
        "L03",
        240,
        r#"<w:r>{R}<w:t xml:space="preserve">L03a before</w:t><w:br/><w:br/><w:t xml:space="preserve">L03c after two</w:t></w:r>"#,
    ),
    // At the paragraph's end.
    (
        "L04",
        240,
        r#"<w:r>{R}<w:t xml:space="preserve">L04a ends in br</w:t><w:br/></w:r>"#,
    ),
    // `textWrapping`, in a run of its own between two runs.
    (
        "L05",
        240,
        r#"<w:r>{R}<w:t xml:space="preserve">L05a wrap</w:t></w:r><w:r>{R}<w:br w:type="textWrapping"/></w:r><w:r>{R}<w:t xml:space="preserve">L05b own run</w:t></w:r>"#,
    ),
    // `w:cr`.
    (
        "L06",
        240,
        r#"<w:r>{R}<w:t xml:space="preserve">L06a cr</w:t><w:cr/><w:t xml:space="preserve">L06b after cr</w:t></w:r>"#,
    ),
    // One blank line on the 12 pt pitch.
    ("B01", 240, ""),
    ("L07", 240, r#"<w:r>{R}<w:t>L07 after blank</w:t></w:r>"#),
    // Two blank lines on a 24 pt pitch: two lines each, IF they are styled.
    ("B02", 480, ""),
    ("B03", 480, ""),
    (
        "L08",
        240,
        r#"<w:r>{R}<w:t>L08 after tall blanks</w:t></w:r>"#,
    ),
    // Page 2: two blank lines with DIFFERENT pitches (24 pt, then 12 pt).
    // The engine styles blank lines at one offset together (the last wins).
    ("L09", 240, r#"<w:r>{R}<w:t>L09 page two</w:t></w:r>"#),
    ("B04", 480, ""),
    ("B05", 240, ""),
    (
        "L10",
        240,
        r#"<w:r>{R}<w:t>L10 after mixed blanks</w:t></w:r>"#,
    ),
];

/// Plain line breaks and blank lines, as Word lays them out (core
/// `ab383b1`: U+2028 is a line break inside a paragraph; `65cf615`: a caret
/// styles a blank line). `scripts/word-line-breaks-probe.sh` has Word export
/// it as PDF; `fixtures/line-breaks.word.json` records each line's page and
/// position. Inter 10 pt (Word lays it in Calibri) on an exact pitch per
/// paragraph ([`LINE_BREAK_CASES`]), every visible line labelled by its
/// first word; `L09` has `w:pageBreakBefore`, so page 2 starts with it.
pub fn line_breaks_docx() -> Vec<u8> {
    let rpr = r#"<w:rPr><w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/><w:sz w:val="20"/></w:rPr>"#;
    let mut body = String::new();
    for (label, line, runs) in LINE_BREAK_CASES {
        let pbb = if *label == "L09" {
            "<w:pageBreakBefore/>"
        } else {
            ""
        };
        body.push_str(&format!(
            r#"<w:p><w:pPr>{pbb}<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="exact"/>{rpr}</w:pPr>{}</w:p>"#,
            runs.replace("{R}", rpr)
        ));
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr>{LINE_BREAKS_PAGE}<w:cols w:space="720"/></w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// The cases of [`run_specials_docx`], in document order: a label and the
/// paragraph's run XML. `{R}` stands for the fixture's run properties. Every
/// paragraph has left tab stops at 1 in and 2 in (72 pt / 144 pt from the
/// left margin), so a tab is visible as where the text after it starts.
///
/// The run children the import turns into a character of the text
/// (`<w:tab/>` and `<w:ptab>` → `\t`, `<w:noBreakHyphen/>` → U+2011,
/// `<w:softHyphen/>` → U+00AD, a plain `<w:br/>` → U+2028, a `<w:sym>` →
/// its Unicode equivalent), which an edited save writes back as the
/// elements they were.
pub const RUN_SPECIAL_CASES: &[(&str, &str)] = &[
    // A tab between two words of one run.
    (
        "T01",
        r#"<w:r>{R}<w:t>T01</w:t><w:tab/><w:t>value</w:t></w:r>"#,
    ),
    // Two tabs in a row.
    (
        "T02",
        r#"<w:r>{R}<w:t>T02a</w:t><w:tab/><w:tab/><w:t>T02b</w:t></w:r>"#,
    ),
    // A tab before the run's first text.
    (
        "T03",
        r#"<w:r>{R}<w:tab/><w:t xml:space="preserve">T03 leading</w:t></w:r>"#,
    ),
    // A non-breaking hyphen.
    (
        "N01",
        r#"<w:r>{R}<w:t>N01 well</w:t><w:noBreakHyphen/><w:t>known</w:t></w:r>"#,
    ),
    // All three kinds in one run, the break with an attribute to keep.
    (
        "M01",
        r#"<w:r>{R}<w:t>M01a</w:t><w:tab/><w:t>b</w:t><w:br w:clear="all"/><w:t>M01c</w:t><w:noBreakHyphen/><w:t>d</w:t></w:r>"#,
    ),
    // A literal tab character inside the `<w:t>` (not Word's own form).
    (
        "X01",
        "<w:r>{R}<w:t xml:space=\"preserve\">X01\tliteral</w:t></w:r>",
    ),
    // Run content the import does not carry as text.
    (
        "S01",
        r#"<w:r>{R}<w:t>S01 before</w:t><w:sym w:font="Symbol" w:char="F0B7"/><w:t>after</w:t></w:r>"#,
    ),
    (
        "H01",
        r#"<w:r>{R}<w:t>H01 hyphen</w:t><w:softHyphen/><w:t>ation</w:t></w:r>"#,
    ),
    (
        "P01",
        r#"<w:r>{R}<w:t>P01 left</w:t><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:t>right</w:t></w:r>"#,
    ),
];

/// Tabs, non-breaking hyphens and the run content around them
/// ([`RUN_SPECIAL_CASES`]): one paragraph per case on the
/// [`line_breaks_docx`] page, Arial 10 pt on a 12 pt pitch, with left tab
/// stops at 72 pt and 144 pt. The save-back tests edit it; Word opens the
/// edited result (`scripts/word-run-specials-probe.sh`).
pub fn run_specials_docx() -> Vec<u8> {
    let rpr = r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/></w:rPr>"#;
    let mut body = String::new();
    for (_, runs) in RUN_SPECIAL_CASES {
        body.push_str(&format!(
            r#"<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="left" w:pos="2880"/></w:tabs><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>{rpr}</w:pPr>{}</w:p>"#,
            runs.replace("{R}", rpr)
        ));
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr>{LINE_BREAKS_PAGE}<w:cols w:space="720"/></w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// The cases of [`symbols_docx`], in document order: a label, the
/// paragraph's extra `w:pPr` children, and its run content (`{R}` is the
/// fixture's run properties). After them comes a one-row table and a
/// paragraph `AFTER` (see [`symbols_docx`]).
pub const SYMBOL_CASES: &[(&str, &str, &str)] = &[
    // Symbol: the bullet, alpha, Delta, Omega.
    (
        "S01",
        "",
        r#"<w:r>{R}<w:t>S01 a</w:t><w:sym w:font="Symbol" w:char="F0B7"/><w:t xml:space="preserve">b </w:t><w:sym w:font="Symbol" w:char="F061"/><w:sym w:font="Symbol" w:char="F044"/><w:sym w:font="Symbol" w:char="F057"/></w:r>"#,
    ),
    // Wingdings: check, ballot x, small square, arrowhead, smiley.
    (
        "S02",
        "",
        r#"<w:r>{R}<w:t xml:space="preserve">S02 </w:t><w:sym w:font="Wingdings" w:char="F0FC"/><w:t xml:space="preserve"> </w:t><w:sym w:font="Wingdings" w:char="F0FB"/><w:t xml:space="preserve"> </w:t><w:sym w:font="Wingdings" w:char="F0A7"/><w:t xml:space="preserve"> </w:t><w:sym w:font="Wingdings" w:char="F0D8"/><w:t xml:space="preserve"> </w:t><w:sym w:font="Wingdings" w:char="F04A"/></w:r>"#,
    ),
    // Wingdings' Windows logo: no Unicode equivalent.
    (
        "S03",
        "",
        r#"<w:r>{R}<w:t>S03 x</w:t><w:sym w:font="Wingdings" w:char="F0FF"/><w:t>y</w:t></w:r>"#,
    ),
    // A long word with optional hyphens in a 90 pt measure: Word breaks it
    // at one and shows a hyphen there.
    (
        "H01",
        r#"<w:ind w:right="3960"/>"#,
        r#"<w:r>{R}<w:t>H01 Donau</w:t><w:softHyphen/><w:t>dampf</w:t><w:softHyphen/><w:t>schiff</w:t><w:softHyphen/><w:t>fahrts</w:t><w:softHyphen/><w:t>gesell</w:t><w:softHyphen/><w:t>schaft</w:t></w:r>"#,
    ),
    // The same word with NO optional hyphens: what Word does without them
    // (its automatic hyphenation is off by default).
    (
        "H02",
        r#"<w:ind w:right="3960"/>"#,
        r#"<w:r>{R}<w:t>H02 Donaudampfschifffahrtsgesellschaft</w:t></w:r>"#,
    ),
    // An optional hyphen where the line does not break: nothing shows.
    (
        "H03",
        "",
        r#"<w:r>{R}<w:t>H03 un</w:t><w:softHyphen/><w:t>broken</w:t></w:r>"#,
    ),
    // Absolute-position tabs: right of the margins, centre of the margins,
    // right of the indents (1 in each side), right of the margins with a
    // dot leader.
    (
        "P01",
        "",
        r#"<w:r>{R}<w:t>P01 left</w:t><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:t>right</w:t></w:r>"#,
    ),
    (
        "P02",
        "",
        r#"<w:r>{R}<w:t>P02</w:t><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/><w:t>centre</w:t></w:r>"#,
    ),
    (
        "P03",
        r#"<w:ind w:left="1440" w:right="1440"/>"#,
        r#"<w:r>{R}<w:t>P03</w:t><w:ptab w:relativeTo="indent" w:alignment="right" w:leader="none"/><w:t>indent</w:t></w:r>"#,
    ),
    (
        "P04",
        "",
        r#"<w:r>{R}<w:t>P04</w:t><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/><w:t>dots</w:t></w:r>"#,
    ),
];

/// Symbols, optional hyphens, absolute-position tabs and empty table cells
/// ([`SYMBOL_CASES`]) on a 5 in × 7 in page (288 pt of text width from a
/// 36 pt margin), Arial 10 pt on an exact 12 pt pitch. Then a
/// one-row table of three 100 pt cells: `C01`, an EMPTY paragraph with an
/// exact 48 pt line, and an empty paragraph with no properties; then a
/// paragraph `AFTER`, whose position shows how tall Word made the row. Word
/// saves it as PDF in `scripts/word-symbols-probe.sh`
/// (`fixtures/symbols.word.json`).
pub fn symbols_docx() -> Vec<u8> {
    let rpr = r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/></w:rPr>"#;
    let spacing = r#"<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>"#;
    let mut body = String::new();
    for (_, ppr, runs) in SYMBOL_CASES {
        body.push_str(&format!(
            r#"<w:p><w:pPr>{spacing}{ppr}{rpr}</w:pPr>{}</w:p>"#,
            runs.replace("{R}", rpr)
        ));
    }
    let cell = |ppr: &str, runs: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>{ppr}{runs}</w:p></w:tc>"#
        )
    };
    let row = [
        cell(
            &format!("<w:pPr>{spacing}</w:pPr>"),
            &format!("<w:r>{rpr}<w:t>C01</w:t></w:r>"),
        ),
        cell(
            r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="960" w:lineRule="exact"/></w:pPr>"#,
            "",
        ),
        cell("", ""),
    ]
    .concat();
    body.push_str(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{row}</w:tr></w:tbl><w:p><w:pPr>{spacing}{rpr}</w:pPr><w:r>{rpr}<w:t>AFTER</w:t></w:r></w:p>"#
    ));
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr>{SYMBOLS_PAGE}<w:cols w:space="720"/></w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// The page of [`symbols_docx`]: 5 in × 7 in, 0.5 in margins → a 288 pt wide
/// body, tall enough for every case on one page.
const SYMBOLS_PAGE: &str = r#"<w:pgSz w:w="7200" w:h="10080"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>"#;

/// One section of [`continuous_docx`]: its label prefix, how it starts, its
/// paragraph count, column count, page size and margins (twips).
#[derive(Debug, Clone, Copy)]
pub struct ContinuousCase {
    /// Paragraph label prefix (`"A2"` → `"A2-01 …"`).
    pub label: &'static str,
    /// `w:type/@w:val`.
    pub kind: &'static str,
    pub lines: u32,
    pub columns: u32,
    /// `w:pgSz` `(w, h)`.
    pub page: (i32, i32),
    /// `w:pgMar` `(top, right, bottom, left)`.
    pub margins: (i32, i32, i32, i32),
}

/// The base page of [`continuous_docx`]: 5 in × 4⅓ in.
pub const CONTINUOUS_PAGE: (i32, i32) = (7200, 6240);
/// The base margins: 0.5 in all round → a 288 pt × 240 pt body = twenty
/// 12 pt lines (two 126 pt columns with the 36 pt gap).
pub const CONTINUOUS_MARGINS: (i32, i32, i32, i32) = (720, 720, 720, 720);

/// The sections of [`continuous_docx`], in order.
pub const CONTINUOUS_CASES: &[ContinuousCase] = {
    const P: (i32, i32) = CONTINUOUS_PAGE;
    const M: (i32, i32, i32, i32) = CONTINUOUS_MARGINS;
    const fn c(
        label: &'static str,
        kind: &'static str,
        lines: u32,
        columns: u32,
        page: (i32, i32),
        margins: (i32, i32, i32, i32),
    ) -> ContinuousCase {
        ContinuousCase {
            label,
            kind,
            lines,
            columns,
            page,
            margins,
        }
    }
    &[
        // (a) same geometry, same columns: an invisible boundary. A3 runs
        //     past the page.
        c("A1", "nextPage", 4, 1, P, M),
        c("A2", "continuous", 4, 1, P, M),
        c("A3", "continuous", 16, 1, P, M),
        // (b) one column → two → one (the newsletter), then two columns
        //     that a nextPage section follows.
        c("B1", "nextPage", 3, 1, P, M),
        c("B2", "continuous", 9, 2, P, M),
        c("B3", "continuous", 3, 1, P, M),
        c("B4", "continuous", 6, 2, P, M),
        // (d) a nextColumn section in two columns.
        c("D1", "nextPage", 3, 2, P, M),
        c("D2", "nextColumn", 3, 2, P, M),
        // (c1) continuous with other LEFT/RIGHT margins, then back.
        c("C1", "nextPage", 3, 1, P, M),
        c("C2", "continuous", 3, 1, P, (720, 1440, 720, 2160)),
        c("C3", "continuous", 3, 1, P, M),
        // (c2) continuous with other TOP/BOTTOM margins, running past the
        //      page.
        c("E1", "nextPage", 3, 1, P, M),
        c("E2", "continuous", 25, 1, P, (1440, 720, 1440, 720)),
        // (c3) continuous with another PAGE SIZE.
        c("F1", "nextPage", 3, 1, P, M),
        c("F2", "continuous", 3, 1, (8640, 6240), M),
        c("G1", "nextPage", 3, 1, P, M),
    ]
};

/// ADR 029 — continuous-section ground truth: where Word puts a section that
/// starts `continuous` (and one `nextColumn`), in every case that decides the
/// lowering (`scripts/word-continuous-probe.sh`; `fixtures/continuous.word.json`
/// records the answer). The sections are [`CONTINUOUS_CASES`]; every
/// paragraph is one line labelled `<section>-NN` (Inter 10 pt on an exact
/// 12 pt grid, as [`breaks_docx`]), so each line's page, column and position
/// read straight off the PDF.
pub fn continuous_docx() -> Vec<u8> {
    let mut body = String::new();
    let sect = |c: &ContinuousCase| {
        let (w, h) = c.page;
        let (t, r, b, l) = c.margins;
        format!(
            r#"<w:sectPr><w:type w:val="{kind}"/><w:pgSz w:w="{w}" w:h="{h}"/><w:pgMar w:top="{t}" w:right="{r}" w:bottom="{b}" w:left="{l}" w:header="360" w:footer="360" w:gutter="0"/><w:cols w:num="{cols}" w:space="720"/></w:sectPr>"#,
            kind = c.kind,
            cols = c.columns,
        )
    };
    let last = CONTINUOUS_CASES.len() - 1;
    for (k, case) in CONTINUOUS_CASES.iter().enumerate() {
        for n in 1..=case.lines {
            // A section's sectPr rides its LAST paragraph; the final
            // section's is the body-level one.
            let s = if n == case.lines && k != last {
                sect(case)
            } else {
                String::new()
            };
            body.push_str(&breaks_para(
                "",
                &breaks_run(&format!("{}-{n:02} {}", case.label, case.kind)),
                &s,
            ));
        }
    }
    let tail = sect(&CONTINUOUS_CASES[last]);
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}{tail}</w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// One section of [`columns_docx`]: its label prefix, how it starts, its
/// paragraph count and its `w:cols` (count, `w:space`, and for unequal
/// columns each `w:col` as `(w, space)` twips). Page and margins are
/// [`CONTINUOUS_PAGE`] / [`CONTINUOUS_MARGINS`] throughout.
#[derive(Debug, Clone, Copy)]
pub struct ColumnsCase {
    pub label: &'static str,
    pub kind: &'static str,
    pub lines: u32,
    pub columns: u32,
    /// `w:cols/@w:space` (twips).
    pub space: i32,
    /// Unequal columns (`w:equalWidth="0"`): each column's `(w, space)`.
    pub widths: &'static [(i32, i32)],
}

/// The sections of [`columns_docx`], in order.
pub const COLUMNS_CASES: &[ColumnsCase] = {
    const fn c(label: &'static str, kind: &'static str, lines: u32, columns: u32) -> ColumnsCase {
        ColumnsCase {
            label,
            kind,
            lines,
            columns,
            space: 720,
            widths: &[],
        }
    }
    &[
        // (h) 1 → 2 → 3 → 1: two column counts change into each other.
        c("H1", "nextPage", 2, 1),
        c("H2", "continuous", 6, 2),
        c("H3", "continuous", 7, 3),
        c("H4", "continuous", 2, 1),
        // (i) a page-starting two-column section, a one-column one, and two
        //     columns again before a nextPage section.
        c("I1", "nextPage", 7, 2),
        c("I2", "continuous", 2, 1),
        c("I3", "continuous", 5, 2),
        // (j) two columns with a 0.25 in gap.
        c("J1", "nextPage", 2, 1),
        ColumnsCase {
            space: 360,
            ..c("J2", "continuous", 6, 2)
        },
        c("J3", "continuous", 2, 1),
        // (k) unequal columns: 2.5 in, 0.25 in gap, 1.25 in.
        c("K1", "nextPage", 2, 1),
        ColumnsCase {
            space: 360,
            widths: &[(3600, 360), (1800, 0)],
            ..c("K2", "continuous", 6, 2)
        },
        c("K3", "continuous", 2, 1),
        // (l) two columns running past the page, then one column.
        c("L1", "nextPage", 3, 1),
        c("L2", "continuous", 50, 2),
        c("L3", "continuous", 2, 1),
        // (m) two columns running past the page before a nextPage section.
        c("M1", "nextPage", 3, 1),
        c("M2", "continuous", 40, 2),
        // (n) 3 → 2 directly.
        c("N1", "nextPage", 2, 1),
        c("N2", "continuous", 7, 3),
        c("N3", "continuous", 5, 2),
        c("N4", "continuous", 1, 1),
        // (p) two columns into two columns with another gap.
        c("P1", "nextPage", 2, 1),
        c("P2", "continuous", 4, 2),
        ColumnsCase {
            space: 360,
            ..c("P3", "continuous", 4, 2)
        },
        c("P4", "continuous", 1, 1),
        // (q) two columns into two columns with the SAME gap.
        c("Q1", "nextPage", 2, 1),
        c("Q2", "continuous", 5, 2),
        c("Q3", "continuous", 5, 2),
        c("Q4", "continuous", 1, 1),
        // (r) the same, the second before a nextPage section.
        c("R1", "nextPage", 2, 1),
        c("R2", "continuous", 5, 2),
        c("R3", "continuous", 5, 2),
        // (o) the document ends in two columns.
        c("O1", "nextPage", 2, 1),
        c("O2", "continuous", 5, 2),
    ]
};

/// ADR 029 — where Word puts a `continuous` section that CHANGES THE
/// COLUMNS mid-page, in the cases [`continuous_docx`] leaves open: one count
/// into another, a page-starting multi-column section, another gap, unequal
/// columns, a section running past the page (balanced on its last page or
/// not), and a document that ends in columns
/// (`scripts/word-columns-probe.sh`; `fixtures/columns.word.json`). Every
/// paragraph is one line holding only its label (`<section>-NN`), short
/// enough for a 72 pt sub-column.
pub fn columns_docx() -> Vec<u8> {
    let (w, h) = CONTINUOUS_PAGE;
    let (t, r, b, l) = CONTINUOUS_MARGINS;
    let sect = |c: &ColumnsCase| {
        let cols = if c.widths.is_empty() {
            format!(r#"<w:cols w:num="{}" w:space="{}"/>"#, c.columns, c.space)
        } else {
            let each: String = c
                .widths
                .iter()
                .map(|(cw, sp)| format!(r#"<w:col w:w="{cw}" w:space="{sp}"/>"#))
                .collect();
            format!(
                r#"<w:cols w:num="{}" w:space="{}" w:equalWidth="0">{each}</w:cols>"#,
                c.columns, c.space
            )
        };
        format!(
            r#"<w:sectPr><w:type w:val="{kind}"/><w:pgSz w:w="{w}" w:h="{h}"/><w:pgMar w:top="{t}" w:right="{r}" w:bottom="{b}" w:left="{l}" w:header="360" w:footer="360" w:gutter="0"/>{cols}</w:sectPr>"#,
            kind = c.kind,
        )
    };
    let mut body = String::new();
    let last = COLUMNS_CASES.len() - 1;
    for (k, case) in COLUMNS_CASES.iter().enumerate() {
        for n in 1..=case.lines {
            let s = if n == case.lines && k != last {
                sect(case)
            } else {
                String::new()
            };
            body.push_str(&breaks_para(
                "",
                &breaks_run(&format!("{}-{n:02}", case.label)),
                &s,
            ));
        }
    }
    let tail = sect(&COLUMNS_CASES[last]);
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}{tail}</w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// A document with a Normal paragraph, a centered Heading1 paragraph, and a
/// paragraph mixing a plain run with a bold red run — enough to exercise style
/// application, direct-format synthesis, and swatch minting. Also carries an
/// unknown part to prove preservation.
pub fn memo_docx() -> Vec<u8> {
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Plain body text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>A Centered Heading</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">Mix of normal and </w:t></w:r>
      <w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>bold red</w:t></w:r>
      <w:r><w:t xml:space="preserve"> text.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
        // An unknown part the model does not touch — must round-trip verbatim.
        ("customXml/unknown.txt", b"paged preserves unknown parts"),
    ])
}

/// A Tier-1a document: `docDefaults` (Calibri 11pt), a paragraph with tab stops +
/// keepNext, and runs exercising underline on/none.
pub fn tier1_docx() -> Vec<u8> {
    let styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:keepNext/>
        <w:tabs>
          <w:tab w:val="left" w:pos="720"/>
          <w:tab w:val="right" w:pos="4320"/>
          <w:tab w:val="clear" w:pos="1440"/>
        </w:tabs>
      </w:pPr>
      <w:r><w:t>Name</w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underlined</w:t></w:r>
      <w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>plain</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", styles.as_bytes()),
    ])
}

/// A document with a bullet list (numId 1) and a numbered list (numId 2, decimal),
/// plus the `numbering.xml` part they resolve through.
pub fn list_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"#;
    let numbering = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First bullet</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second bullet</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Step one</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Step two</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/numbering.xml", numbering.as_bytes()),
    ])
}

/// A document with a 2-column table exercising `gridSpan` (a spanning header) and
/// `vMerge` (a vertically merged cell).
pub fn table_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before the table.</w:t></w:r></w:p>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Title spanning</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Right top</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>Right bottom</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>After the table.</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

/// A document with an inline image (`w:drawing` → `wp:inline` → picture blip)
/// plus the `word/media/image1.png` media part it embeds.
pub fn image_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;
    // extent 914400 x 685800 EMU = 72 x 54 pt. All namespaces on the root so the
    // nested wp:/a:/pic:/r: content resolves.
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Text before image.</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="685800"/>
            <wp:docPr id="1" name="Picture 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rId100"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="685800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
    // The media bytes are opaque to the importer (it resolves + carries them);
    // a PNG signature makes the fixture realistic without a full encoder.
    let png = b"\x89PNG\r\n\x1a\n-fake-image-bytes-for-conformance-";
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/media/image1.png", png),
    ])
}

/// A document with an external hyperlink (`w:hyperlink r:id=…` → an
/// `TargetMode="External"` relationship).
pub fn hyperlink_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://paged.media/" TargetMode="External"/>
</Relationships>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Visit </w:t></w:r>
      <w:hyperlink r:id="rId50"><w:r><w:t>Paged Media</w:t></w:r></w:hyperlink>
      <w:r><w:t xml:space="preserve"> today.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

/// A document whose hyperlinks are expressed as FIELDS (not `w:hyperlink`): one
/// complex `fldChar begin/instrText/separate/result/end` field and one simple
/// `w:fldSimple`, both carrying a `HYPERLINK "url"` instruction. No rels part —
/// the URL is inline in the field code.
pub fn field_hyperlink_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    // The complex field splits its instruction across two instrText runs (as Word
    // often does) to exercise instruction accumulation.
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Go </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> HYPERLINK &quot;https://example.com/</w:instrText></w:r>
      <w:r><w:instrText xml:space="preserve">complex&quot; </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>complex link</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:fldSimple w:instr="HYPERLINK &quot;https://example.com/simple&quot;"><w:r><w:t>simple link</w:t></w:r></w:fldSimple>
      <w:r><w:t xml:space="preserve"> done.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

/// A document with a footnote: an in-text `w:footnoteReference` plus a
/// `word/footnotes.xml` carrying the note body (and Word's two separator
/// pseudo-notes, which must be skipped).
pub fn footnote_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Body with a note</w:t></w:r>
      <w:r><w:footnoteReference w:id="2"/></w:r>
      <w:r><w:t xml:space="preserve"> and more.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
    let footnotes = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>The note body.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/footnotes.xml", footnotes.as_bytes()),
    ])
}

/// A plain 3-row x 2-col table (no merges) — the clean case for testing ROW
/// alignment, where deleting a middle row is unambiguous.
pub fn simple_table_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let mut rows = String::new();
    for r in 0..3 {
        // Each row carries a DISTINCT, UNMODELLED property (`w:trHeight`). paged
        // does not read it, so it is the marker that proves WHICH `<w:tr>` node
        // survived an edit — text alone cannot show that.
        rows.push_str(&format!(
            "<w:tr><w:trPr><w:trHeight w:val=\"{}\"/></w:trPr>",
            100 + r
        ));
        for c in 0..2 {
            rows.push_str(&format!(
                "<w:tc><w:p><w:r><w:t>R{r}C{c}</w:t></w:r></w:p></w:tc>"
            ));
        }
        rows.push_str("</w:tr>");
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>{rows}</w:tbl>
  </w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

/// Blank lines inside table cells (core `65cf615`: a caret, `cell`
/// addressed, styles them). One row, three cells: `A`, a 24 pt blank line,
/// a 12 pt blank line, `B` (two blank lines at ONE cell offset with
/// different pitches — diagnosed); `C`, a 24 pt blank line, `D`, and a
/// trailing 24 pt blank line; and an empty cell.
pub fn cell_blank_lines_docx() -> Vec<u8> {
    let p = |line: u32, text: &str| {
        let runs = if text.is_empty() {
            String::new()
        } else {
            format!("<w:r><w:t>{text}</w:t></w:r>")
        };
        format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="exact"/></w:pPr>{runs}</w:p>"#
        )
    };
    let cell = |paras: &[(u32, &str)]| {
        let body: String = paras.iter().map(|(l, t)| p(*l, t)).collect();
        format!("<w:tc>{body}</w:tc>")
    };
    let row = [
        cell(&[(240, "A"), (480, ""), (240, ""), (240, "B")]),
        cell(&[(240, "C"), (480, ""), (240, "D"), (480, "")]),
        cell(&[(240, "")]),
    ]
    .concat();
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{row}</w:tr></w:tbl>{}<w:sectPr>{LINE_BREAKS_PAGE}</w:sectPr></w:body>
</w:document>"#,
        p(240, "after")
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}

/// An outer table whose FIRST cell contains a NESTED table. Exercises the
/// locator's nesting handling: a nested `<w:tr>` must not be counted against the
/// outer table's rows, or an edit aimed at an outer row lands inside the nested
/// table. (A `<w:tc>` holding a table must still end with a `<w:p>`.)
pub fn nested_table_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr><w:tc>
        <w:tbl>
          <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
          <w:tr><w:tc><w:p><w:r><w:t>INNER</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
        <w:p><w:r><w:t>outer r0</w:t></w:r></w:p>
      </w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>outer r1</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

/// A document with INTERNAL (bookmark) links in both Word forms: a
/// `<w:hyperlink w:anchor="...">` and a `HYPERLINK \l "bm"` field. Neither has an
/// external URL, so both must stay styled-only — never a native clickable link.
pub fn internal_anchor_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">See </w:t></w:r>
      <w:hyperlink w:anchor="chapter2"><w:r><w:t>chapter two</w:t></w:r></w:hyperlink>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:fldSimple w:instr="HYPERLINK \l &quot;bookmark1&quot;"><w:r><w:t>the bookmark</w:t></w:r></w:fldSimple>
      <w:r><w:t xml:space="preserve">.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

/// A document with a header, a footer, and non-HYPERLINK fields (`PAGE`, `DATE`)
/// — the constructs paged.doc preserves but cannot yet place or recompute.
pub fn header_field_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Printed on </w:t></w:r>
      <w:fldSimple w:instr="DATE \@ &quot;d MMMM yyyy&quot;"><w:r><w:t>23 July 2026</w:t></w:r></w:fldSimple>
      <w:r><w:t xml:space="preserve">, page </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE  \* MERGEFORMAT </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>7</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rH1"/>
      <w:footerReference w:type="default" r:id="rF1"/>
      <w:pgSz w:w="11906" w:h="16838"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
    let header = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Quarterly Report</w:t></w:r></w:p></w:hdr>"#;
    let footer = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:ftr>"#;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/header1.xml", header.as_bytes()),
        ("word/footer1.xml", footer.as_bytes()),
    ])
}

/// The smallest well-formed document: one paragraph, one run, no styles part.
pub fn one_paragraph_docx() -> Vec<u8> {
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello, world.</w:t></w:r></w:p></w:body>
</w:document>"#;
    let root_rels_only_doc = ROOT_RELS;
    zip_parts(&[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#
                .as_bytes(),
        ),
        ("_rels/.rels", root_rels_only_doc.as_bytes()),
        ("word/document.xml", document.as_bytes()),
    ])
}

const W_NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A plain one-run paragraph (Word's default formatting), optionally ending
/// its section with `sect` (a `w:sectPr`).
fn plain_para(text: &str, sect: &str) -> String {
    let ppr = if sect.is_empty() {
        String::new()
    } else {
        format!("<w:pPr>{sect}</w:pPr>")
    };
    format!(r#"<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#)
}

/// A paragraph that holds only a page break.
const PAGE_BREAK_PARA: &str = r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#;

/// The header/footer parts of [`headers_docx`]: (rel id, part, footer?, label).
pub const HEADER_PARTS: &[(&str, &str, bool, &str)] = &[
    ("rH1d", "header1.xml", false, "H1-default"),
    ("rH1f", "header2.xml", false, "H1-first"),
    ("rH1e", "header3.xml", false, "H1-even"),
    ("rF1d", "footer1.xml", true, "F1-default"),
    ("rH3d", "header4.xml", false, "H3-default"),
];

/// thoughts ADR 033 / RFI DOC-05 — which header and footer Word shows on
/// each page of a three-section document (`scripts/word-headers-probe.sh`
/// has Word save it as PDF; `fixtures/headers.word.json` records, per page,
/// the header and footer text).
///
/// `settings.xml` has `w:evenAndOddHeaders`. Every section is three pages
/// (two page breaks), each page's one body line names it (`S2 page 1`), and
/// every header/footer is its label plus a `PAGE` field.
///
/// - **S1** (`titlePg`): header default / first / even, footer default
///   only (so its first and even footers are blank, unless Word says
///   otherwise).
/// - **S2**: NO references and no `titlePg`, page numbering restarting at 1
///   (`w:pgNumType w:start="1"`): its headers are S1's, and its first page
///   is number 1, odd, though it is the fourth sheet.
/// - **S3** (`titlePg`): its own default header only: the first and even
///   headers and every footer come from S1.
pub fn headers_docx() -> Vec<u8> {
    let mut content_types = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
"#,
    );
    let mut rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
"#,
    );
    let mut parts: Vec<(String, String)> = Vec::new();
    for (id, part, footer, label) in HEADER_PARTS {
        let (kind, root) = if *footer {
            ("footer", "ftr")
        } else {
            ("header", "hdr")
        };
        content_types.push_str(&format!(
            r#"  <Override PartName="/word/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>
"#
        ));
        rels.push_str(&format!(
            r#"  <Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{part}"/>
"#
        ));
        parts.push((
            format!("word/{part}"),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:{root} {W_NS}><w:p><w:r><w:t xml:space="preserve">{label} p</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:{root}>"#
            ),
        ));
    }
    content_types.push_str("</Types>");
    rels.push_str("</Relationships>");
    let settings = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W_NS}><w:evenAndOddHeaders/></w:settings>"#
    );
    let page = r#"<w:pgSz w:w="7200" w:h="3840"/><w:pgMar w:top="1080" w:right="720" w:bottom="1080" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>"#;
    let s1 = format!(
        r#"<w:sectPr><w:headerReference w:type="default" r:id="rH1d"/><w:headerReference w:type="first" r:id="rH1f"/><w:headerReference w:type="even" r:id="rH1e"/><w:footerReference w:type="default" r:id="rF1d"/>{page}<w:titlePg/></w:sectPr>"#
    );
    let s2 = format!(
        r#"<w:sectPr><w:type w:val="nextPage"/>{page}<w:pgNumType w:start="1"/></w:sectPr>"#
    );
    let s3 = format!(
        r#"<w:sectPr><w:headerReference w:type="default" r:id="rH3d"/><w:type w:val="nextPage"/>{page}<w:titlePg/></w:sectPr>"#
    );
    let mut body = String::new();
    for (k, sect) in [(1, s1.as_str()), (2, s2.as_str())] {
        body.push_str(&plain_para(&format!("S{k} page 1"), ""));
        body.push_str(PAGE_BREAK_PARA);
        body.push_str(&plain_para(&format!("S{k} page 2"), ""));
        body.push_str(PAGE_BREAK_PARA);
        body.push_str(&plain_para(&format!("S{k} page 3"), sect));
    }
    body.push_str(&plain_para("S3 page 1", ""));
    body.push_str(PAGE_BREAK_PARA);
    body.push_str(&plain_para("S3 page 2", ""));
    body.push_str(PAGE_BREAK_PARA);
    body.push_str(&plain_para("S3 page 3", ""));
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {W_NS}><w:body>{body}{s3}</w:body></w:document>"#
    );
    let mut all: Vec<(&str, &[u8])> = vec![
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/settings.xml", settings.as_bytes()),
    ];
    for (name, xml) in &parts {
        all.push((name.as_str(), xml.as_bytes()));
    }
    zip_parts(&all)
}

/// The sections of [`footnote_numbering_docx`]: (label, its own
/// `w:footnotePr` children, pages, footnotes per page).
pub const FOOTNOTE_NUMBERING_SECTIONS: &[(&str, &str, u32, u32)] = &[
    ("S1", r#"<w:numStart w:val="3"/>"#, 1, 2),
    (
        "S2",
        r#"<w:numFmt w:val="upperLetter"/><w:numRestart w:val="eachSect"/>"#,
        1,
        2,
    ),
    ("S3", "", 1, 2),
    ("S4", r#"<w:numRestart w:val="eachPage"/>"#, 2, 2),
];

/// thoughts ADR 034 / RFI DOC-06 — footnote numbering ground truth
/// (`scripts/word-footnotes-probe.sh`; `fixtures/footnote-numbering.word.json`
/// records each note's mark as Word prints it).
///
/// `settings.xml` says `w:footnotePr/w:numFmt="lowerRoman"`; each section of
/// [`FOOTNOTE_NUMBERING_SECTIONS`] has its own `w:footnotePr` (or none), and
/// each page carries its footnotes, named `S2-p1-n2`. The questions: does a
/// section's `w:footnotePr` override the document's field by field or
/// whole; does a section without one inherit the previous section's; where
/// does an `eachSect` / `eachPage` restart start.
pub fn footnote_numbering_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"#;
    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"#;
    let settings = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W_NS}><w:footnotePr><w:numFmt w:val="lowerRoman"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>"#
    );
    let page = r#"<w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>"#;
    let mut notes = String::from(
        r#"<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#,
    );
    let mut body = String::new();
    let mut id = 1;
    let n = FOOTNOTE_NUMBERING_SECTIONS.len();
    for (k, (label, pr, pages, per_page)) in FOOTNOTE_NUMBERING_SECTIONS.iter().enumerate() {
        let fpr = if pr.is_empty() {
            String::new()
        } else {
            format!("<w:footnotePr>{pr}</w:footnotePr>")
        };
        let sect = format!(r#"<w:sectPr>{fpr}<w:type w:val="nextPage"/>{page}</w:sectPr>"#);
        for pg in 1..=*pages {
            let mut runs =
                format!(r#"<w:r><w:t xml:space="preserve">{label} page {pg}</w:t></w:r>"#);
            for note in 1..=*per_page {
                let name = format!("{label}-p{pg}-n{note}");
                runs.push_str(&format!(
                    r#"<w:r><w:t xml:space="preserve"> {name}</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{id}"/></w:r>"#
                ));
                notes.push_str(&format!(
                    r#"<w:footnote w:id="{id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note {name}</w:t></w:r></w:p></w:footnote>"#
                ));
                id += 1;
            }
            let last_page = pg == *pages;
            let ppr = if last_page && k + 1 < n {
                format!("<w:pPr>{sect}</w:pPr>")
            } else {
                String::new()
            };
            body.push_str(&format!("<w:p>{ppr}{runs}</w:p>"));
            if !last_page {
                body.push_str(PAGE_BREAK_PARA);
            }
        }
        if k + 1 == n {
            body.push_str(&sect);
        }
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {W_NS}><w:body>{body}</w:body></w:document>"#
    );
    let footnotes = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes {W_NS}>{notes}</w:footnotes>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/settings.xml", settings.as_bytes()),
        ("word/footnotes.xml", footnotes.as_bytes()),
    ])
}

/// A `w:drawing` holding a picture of `cx` × `cy` EMU, as `wp:inline`
/// (`anchor` empty) or as a `wp:anchor` whose attributes are `anchor` and
/// whose position/wrap children are `children`.
fn picture_drawing(anchor: &str, children: &str, cx: i64, cy: i64) -> String {
    let graphic = format!(
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>"#
    );
    if anchor.is_empty() {
        format!(
            r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="Picture"/>{graphic}</wp:inline></w:drawing>"#
        )
    } else {
        format!(
            r#"<w:drawing><wp:anchor {anchor}>{children}<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{{WRAP}}<wp:docPr id="2" name="Float"/>{graphic}</wp:anchor></w:drawing>"#
        )
    }
}

/// The floating drawings of [`floats_docx`], one paragraph each: (label,
/// `wp:anchor` attributes, position children, wrap element).
pub const FLOAT_CASES: &[(&str, &str, &str, &str)] = &[
    (
        "F1 square beside its paragraph",
        r#"distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1""#,
        r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>152400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
        r#"<wp:wrapSquare wrapText="bothSides"/>"#,
    ),
    (
        "F2 top and bottom, centred on the margin",
        r#"distT="25400" distB="50800" distL="0" distR="0" simplePos="0" relativeHeight="251660288" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="0""#,
        r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>"#,
        r#"<wp:wrapTopAndBottom/>"#,
    ),
    (
        "F3 behind the text",
        r#"distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251661312" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1""#,
        r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV>"#,
        r#"<wp:wrapNone/>"#,
    ),
    (
        "F4 simple position",
        r#"distT="0" distB="0" distL="0" distR="0" simplePos="1" relativeHeight="251662336" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1""#,
        r#"<wp:simplePos x="1270000" y="2540000"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
        r#"<wp:wrapTight wrapText="right" distL="38100" distR="38100"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>"#,
    ),
];

/// thoughts ADR 035 / RFI DOC-07 — floating drawings (`wp:anchor`): one
/// paragraph per [`FLOAT_CASES`] entry, each with its float as the
/// paragraph's first run (72 × 54 pt); then a paragraph whose ONE run holds
/// an inline picture and a floating one; then a table cell holding a float;
/// then a paragraph holding a drawing that is not a picture (a chart
/// reference).
pub fn floats_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;
    let (cx, cy) = (914400, 685800);
    let float = |attrs: &str, children: &str, wrap: &str| {
        picture_drawing(attrs, children, cx, cy).replace("{WRAP}", wrap)
    };
    let mut body = String::new();
    for (label, attrs, children, wrap) in FLOAT_CASES {
        body.push_str(&format!(
            r#"<w:p><w:r>{}</w:r><w:r><w:t xml:space="preserve">{label}: the paragraph the float is anchored in.</w:t></w:r></w:p>"#,
            float(attrs, children, wrap)
        ));
    }
    let (_, a1, c1, w1) = FLOAT_CASES[0];
    body.push_str(&format!(
        r#"<w:p><w:r>{}{}</w:r><w:r><w:t xml:space="preserve">F5 two drawings in one run.</w:t></w:r></w:p>"#,
        picture_drawing("", "", cx, cy),
        float(a1, c1, w1)
    ));
    body.push_str(&format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r>{}</w:r><w:r><w:t xml:space="preserve">F6 in a cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        float(a1, c1, w1)
    ));
    body.push_str(
        r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="685800"/><wp:docPr id="9" name="Chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve">F7 a chart.</w:t></w:r></w:p>"#,
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    );
    let png = b"\x89PNG\r\n\x1a\n-fake-image-bytes-for-conformance-";
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/media/image1.png", png),
    ])
}

/// A black 96 × 96 grey PNG (a real one: Word refuses a fake image with a
/// repair prompt; opaque, so its box shows in Word's PDF; and not 1 × 1,
/// which Word draws at two thirds of its VML size).
const TINY_PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x60, 0x00, 0x00, 0x00, 0x60, 0x08, 0x00, 0x00, 0x00, 0x00, 0xc7, 0xf3, 0x28,
    0xe4, 0x00, 0x00, 0x00, 0x20, 0x49, 0x44, 0x41, 0x54, 0x78, 0xda, 0xed, 0xc1, 0x81, 0x00, 0x00,
    0x00, 0x00, 0xc3, 0xa0, 0xf9, 0x53, 0x5f, 0xe0, 0x08, 0x55, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x7c, 0x03, 0x24, 0x60, 0x00, 0x01, 0x7c, 0xec, 0x86, 0xc6, 0x00, 0x00, 0x00,
    0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
];

/// ADR 029 acceptance against real Word documents (`docs/acceptance-real-docx.md`):
/// the constructs the corpus documents exposed, one paragraph each, on a
/// 5 in × 7 in page with 0.5 in margins (a 288 pt text width). Word's answer:
/// `scripts/word-real-docx-probe.sh` → `fixtures/real-docx.word.json`.
///
/// - `T01` — a TOC entry: its style (`Toc2`) sets a left stop at 0.5 in and a
///   RIGHT stop with a dot leader at 300 pt, past the 288 pt margin; the
///   paragraph adds a left stop at 1 in, which the entry's text runs past.
///   Word merges the paragraph's stops with the style's and puts the page
///   number, after the dot leader, at the margin.
/// - `B01` — a paragraph style based on a CHARACTER style (`FootRef`), with a
///   bulleted `w:numPr` of its own; the document defaults set no size. Word
///   ignores the cross-type `basedOn` (10 pt, the bare defaults) and bullets
///   the paragraph through its style.
/// - `N01` — a paragraph naming no style: Word lays it in the default
///   paragraph style (Normal, 12 pt), not on the bare defaults.
/// - `V01` — a legacy VML inline picture 100 pt square; `A01` after it shows
///   how much room Word gave the picture's line.
/// - `F01` — a FLOATING VML picture (100 × 50 pt, `position:absolute`) whose
///   text wraps top and bottom; `A02` after it shows Word giving it a band.
pub fn real_docx_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;
    let styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="240"/></w:pPr>
    <w:rPr><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Toc2">
    <w:name w:val="toc 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:leader="dot" w:pos="6000"/></w:tabs></w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="FootRef">
    <w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="BulletInd">
    <w:name w:val="Bullet Indented"/>
    <w:basedOn w:val="FootRef"/>
    <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:after="240"/></w:pPr>
  </w:style>
</w:styles>"#;
    let numbering = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#;
    let document = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Toc2"/><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr><w:r><w:t>T01</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>An entry past the inch</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>7</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="BulletInd"/></w:pPr><w:r><w:t>B01 bulleted through its style</w:t></w:r></w:p>
    <w:p><w:r><w:t>N01 names no style</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">V01 </w:t></w:r><w:r><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="Pic1" o:spid="_x0000_i1025" type="#_x0000_t75" style="width:100pt;height:100pt;visibility:visible;mso-wrap-style:square"><v:imagedata r:id="rId3" o:title=""/></v:shape></w:pict></w:r></w:p>
    <w:p><w:r><w:t>A01 after the picture</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">F01 </w:t></w:r><w:r><w:pict><v:shape id="Pic2" o:spid="_x0000_s1026" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:100pt;height:50pt;z-index:251658240;mso-position-horizontal-relative:text;mso-position-vertical-relative:text"><v:imagedata r:id="rId3" o:title=""/><w10:wrap type="topAndBottom"/></v:shape></w:pict></w:r></w:p>
    <w:p><w:r><w:t>A02 after the floating picture</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="7200" w:h="10080"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>"##;
    zip_parts(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", styles.as_bytes()),
        ("word/numbering.xml", numbering.as_bytes()),
        ("word/media/image1.png", TINY_PNG),
    ])
}

/// The paragraphs of [`keeps_docx`], in order: a label, its `w:pPr` keep
/// element, and its line count (a multi-line paragraph's lines are
/// `w:br` line breaks, `K1a` / `K1b` …). `F` rows are one-line fillers
/// (`F001` …), as many as the number says.
pub const KEEPS_CASES: &[(&str, &str, u32)] = &[
    ("F", "", 35),
    // One line fits at the page foot: widow control moves it all.
    ("K1", "<w:widowControl/>", 3),
    ("F", "", 32),
    // The same, widow control off: the first line stays.
    ("K2", r#"<w:widowControl w:val="0"/>"#, 3),
    ("F", "", 32),
    // Two of three fit: keepLines moves it all.
    ("K3", "<w:keepLines/>", 3),
    ("F", "", 30),
    // Three of four fit: widow control pulls a second line over (2 | 2).
    ("K4", "<w:widowControl/>", 4),
    ("F", "", 33),
    // No keep element anywhere in the hierarchy: what Word does by default.
    ("K5", "", 3),
    ("F", "", 2),
];

/// ADR 029 decision 3 against Word: `w:widowControl` and `w:keepLines`
/// ([`KEEPS_CASES`]) on 5 in × 7 in pages with 0.5 in margins, Arial 10 pt
/// on an exact 12 pt pitch, so a page holds 36 lines. Word's answer:
/// `scripts/word-keeps-probe.sh` → `fixtures/keeps.word.json`.
pub fn keeps_docx() -> Vec<u8> {
    let rpr = r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/></w:rPr>"#;
    let spacing = r#"<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>"#;
    let mut body = String::new();
    let mut filler = 0;
    for (label, keep, lines) in KEEPS_CASES {
        if *label == "F" {
            for _ in 0..*lines {
                filler += 1;
                body.push_str(&format!(
                    r#"<w:p><w:pPr>{spacing}{rpr}</w:pPr><w:r>{rpr}<w:t>F{filler:03}</w:t></w:r></w:p>"#
                ));
            }
            continue;
        }
        let runs: Vec<String> = (0..*lines)
            .map(|k| format!("<w:t>{label}{}</w:t>", (b'a' + k as u8) as char))
            .collect();
        body.push_str(&format!(
            r#"<w:p><w:pPr>{keep}{spacing}{rpr}</w:pPr><w:r>{rpr}{}</w:r></w:p>"#,
            runs.join("<w:br/>")
        ));
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr>{SYMBOLS_PAGE}</w:sectPr></w:body>
</w:document>"#
    );
    zip_parts(&[
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
    ])
}
