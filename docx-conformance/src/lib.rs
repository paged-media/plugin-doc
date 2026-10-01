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
/// (`<w:tab/>` → `\t`, `<w:noBreakHyphen/>` → U+2011, a plain `<w:br/>` →
/// U+2028), and the ones it does not carry at all (`<w:sym>`,
/// `<w:softHyphen/>`, `<w:ptab>`), which an edited save must refuse.
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
