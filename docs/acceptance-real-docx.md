# ADR 029 acceptance: real Word documents

**Status (2026-10-02): NOT met.** ADR 029 decision 6 asks that "a real 50+ page
`.docx` from the corpus opens with pagination matching Word's own PDF export within
a stated tolerance: page count equal, page-start paragraphs equal on at least 95% of
pages, deviations itemised". Three real documents were measured. None meets it: the
best page-start match is 8%. This record gives the numbers, every deviation, what
was fixed in plugin-doc on the way (each fix checked against Word), and what is
left, in priority order. Most of what is left is engine work (core), not lowering.

## Documents

Chosen by content (page count and constructs read from the files, not their names),
from the private corpus, addressed by path:

| id | corpus path (sha256 prefix) | Word's pages | what is in it |
|---|---|---:|---|
| parentinvguid | `docx/poi-converted/parentinvguid.doc.docx` (`33d8945230e45d11`) | 61 | US Letter, 3 sections, a 5-page table of contents (TOC field, dot-leader right tabs), headings with keepNext, 82 list paragraphs (some bulleted through their style), 5 footnotes, footers, 6 page breaks, a 5-page table, 2 VML pictures. Times New Roman 12 pt. |
| bug59058 | `docx/poi/bug59058.docx` (`89d0abb8eed1a58f`) | 52 | A web page saved as Word: six tables (three 39-row data tables, 3 Word pages each), 18 DrawingML pictures, form controls and hidden (`w:vanish`) text, then two scientific articles. Times, Times New Roman, Calibri, MinionPro, Verdana. |
| Bug50936_3 | `docx/poi-converted/Bug50936_3.doc.docx` (`bdd90e4f945db6c6`) | 50 | French internship report, A4: TOC, multi-level numbered headings, 249 list paragraphs, 4 tables, 8 VML pictures (3 inline, 5 floating with top-and-bottom wrap), headers with a table, 74 explicit widow-control paragraphs. Arial 11 pt. |

## Method

- **Word's answer.** Microsoft Word 16 (macOS) saved each file as PDF
  (`scripts/word-pagination-probe.sh`, staged in Word's own container). Per page,
  the body lines inside the section's margin box (`scripts/real-docx/word_map.py`,
  `pdftotext -bbox-layout`; headers and footers fall outside it).
- **Ours.** The editor on the PUBLISHED engine `@paged-media/canvas-wasm@0.64.0`
  (editor `main` 04e4f80 in a worktree, `@paged-media/doc` linked to this repo's
  bundle), File▸Open of the `.docx`, page growth settled; then per paragraph of
  every story the pages its lines land on (`paragraphBounds` walk +
  `selectionGeometry`, `scripts/real-docx/measure.spec.ts`).
- **The metric.** A page's START PARAGRAPH is the body paragraph holding its first
  line (one continuing from the previous page counts; empty paragraphs are invisible
  on both sides). Word's paragraphs are found in its PDF text by aligning our
  paragraph texts in order (`scripts/real-docx/compare.py`; 624/626, 173/178 and
  813/813 non-empty paragraphs aligned). A page matches when both start in the same
  paragraph. `Δ ¶` is how many paragraphs ours is ahead (+) or behind (−).
- **Two measurements per document.** *As opened*: what a user gets, the editor's
  default face (Inter) for every family. *Word's faces*: Word's own font files
  (from Word.app; local only, never committed) registered before the open, then the
  document saved and loaded again, because the engine builds its font table once,
  at load, from the families the stories reference then, and the skeleton's stories
  are empty (see P0-1 below). Without the reload, registering fonts changes
  nothing; measured: identical geometry with and without Times New Roman.
- **Fonts Word used** (`pdffonts` on its PDFs): Times New Roman, Arial, Calibri,
  Verdana, Century, Courier New, Symbol, Wingdings, Lucida Sans Unicode, Cambria
  Math, Times. Word had no MinionPro, MyriadPro, Helvetica or agencyr
  (bug59058 asks for them) and substituted silently. The *Word's faces*
  measurement registers Times New Roman, Arial, Calibri, Verdana, Tahoma, Century,
  Lucida Sans, Symbol, Wingdings, Courier New and Arial Unicode MS; not Times (the
  macOS face bug59058 uses in 1,211 runs), which stays on the default face.

Reproduce: `bash scripts/real-docx-acceptance.sh <editor-worktree> <out-dir> <docx…>`.
It is a script, not an editor spec: the documents are the private corpus and the
fonts are this machine's, so CI could run neither.

## Results

| document | measurement | pages (ours / Word) | page-start match | within ±1 ¶ | median \|Δ ¶\| |
|---|---|---:|---:|---:|---:|
| parentinvguid | before, as opened | 65 / 61 | 2 / 61 = 3.3% | 2 | 28.5 |
| parentinvguid | before, Word's faces | 56 / 61 | 5 / 61 = 8.2% | 11 | 7 |
| parentinvguid | **after, as opened** | 75 / 61 | 2 / 61 = 3.3% | 2 | 72 |
| parentinvguid | **after, Word's faces** | 64 / 61 | 4 / 61 = 6.6% | 10 | 22.5 |
| bug59058 | before (both) | 34 / 52 | 0 / 52 = 0% | 0 | 128 |
| bug59058 | **after (both)** | 36 / 52 | 0 / 52 = 0% | 0 | 118 |
| Bug50936_3 | before, as opened | 47 / 50 | 2 / 50 = 4.0% | 3 | 28 |
| Bug50936_3 | before, Word's faces | 42 / 50 | 2 / 50 = 4.0% | 2 | 58 |
| Bug50936_3 | **after, as opened** | 46 / 50 | 4 / 50 = 8.0% | 6 | 13 |
| Bug50936_3 | **after, Word's faces** | 45 / 50 | 4 / 50 = 8.0% | 5 | 48.5 |

"Before" is plugin-doc `070ae48`; "after" is this change on top of `3b6015e`
(measured again after that rebase: the same numbers). bug59058 measures the
same in both modes: its tables, pictures and hidden text dominate.

**The fixes do not all move the aggregate the right way, and that is reported as
measured.** Each fix below reproduces Word's own answer on a fixture, and the
editor reproduces the keeps fixture's page map exactly. But the engine has errors
in the other direction (P0-2, P0-3, P0-4), which the lowering errors used to
cancel: with the tab, style and VML fixes and before widow control, parentinvguid
measured 60 / 61 pages (8.2%); widow control, which Word applies by default,
brought out an engine keep anomaly (P0-4) and took it to 64.

## What was fixed (plugin-doc)

Each against Word's answer: `docx-conformance/fixtures/real-docx.word.json`
(`real_docx_docx()`, `scripts/word-real-docx-probe.sh`, tests
`docx-conformance/tests/real_docx.rs`) and `fixtures/keeps.word.json`
(`keeps_docx()`, `scripts/word-keeps-probe.sh`, `tests/keeps.rs`).

1. **Tab alignments are IDML's names.** The lowering emitted Word's names
   (`right`, `center`, `decimal`); the engine's composer reads only `RightAlign`,
   `CenterAlign`, `CharacterAlign`, so every right tab from Word was a left tab
   (every TOC page number wrapped to its own line). The unit tests asserted the
   Word names and passed.
2. **Tab stops merge down the style chain.** Word adds a paragraph's stops to its
   style's (a `clear` stop removes one); the native tab list replaces the style's,
   so the lowering carries the merged set. Before, a TOC paragraph with one direct
   stop lost its style's right stop.
3. **Dot / hyphen / underscore / middle-dot leaders** are carried.
4. **A right/centre/decimal stop past the right margin moves to the margin.** Word
   sets the text after it out in the margin on the same line (T01: the page number
   ends at the 300 pt stop, margin 288); a native frame would move it to a new line.
   It keeps Word's line, a few points left; an info diagnostic says so.
5. **No automatic hyphenation** unless `w:autoHyphenation` (Word's default is off,
   the engine's on). A paragraph with an optional hyphen keeps it on, which the
   composer needs to break there.
6. **The default paragraph style.** A paragraph naming no style is in Word's
   default paragraph style (Normal), not on the bare document defaults: before,
   every unstyled paragraph lost Normal's size and spacing (12 pt became 10 pt).
7. **The document defaults carry Word's 10 pt** where `docDefaults` set no size
   (the engine's default is 12 pt).
8. **A `basedOn` naming a style of another type is ignored**, as Word does (a
   paragraph style "based on" the footnote-reference character style); before, the
   style hung off a parent the host never created and fell back to the engine's
   root style (the wrong font, size and spacing). An info diagnostic names it.
9. **Lists from the paragraph style** (`w:numPr` in a style; Word applies it), and
   **the level's own indents** (`w:lvl/w:pPr/w:ind`) instead of 18 pt per level,
   below the paragraph's and its style's own indents (Word's precedence).
10. **Widow control and keepLines.** `w:widowControl` was never imported, and
    `w:keepLines` lowered to InDesign's widow rule. Word measured: widow control is
    ON where nothing sets it (K5), off for `w:val="0"`; keepLines keeps all lines.
    They lower to KeepLinesTogether At Start / At End 2 / 2, and All Lines. The
    editor lays the fixture on Word's six pages.
11. **Legacy VML pictures** (`w:pict` / `v:imagedata`) become images: inline ones,
    and floating ones with a top-and-bottom wrap (Word gives them a band of their
    own). Other VML (shapes, text boxes) is reported as not placed, not dropped
    silently.
12. A table whose cells the engine refuses, or one it does not create, is now
    reported (the outcome was ignored).
13. `word-pagination-probe.sh` addresses the document by name and waits for the
    PDF to stop growing: with "active document", a 61-page file still opening made
    Word save the PREVIOUS document under the next file's name.

## Deviations, by cause

The per-page tables follow. Causes, from inspecting Word's and our pages side by
side (likely causes, not each proven):

- **parentinvguid.** Pages 1–3 match. 4–6 (TOC): one entry off per page (line
  breaks of two-line entries; TOC hyperlinks drawn blue and underlined, Word draws
  them plain). 7–11: Word's pages carry footnotes (pages 7, 9, 10) that we do not
  place, so ours hold more (P0-5). 12–34: drift of +1…+15 paragraphs: line
  breaking (P0-3) plus the footnotes. 19–20 and 23–24: a page with three lines, or
  a lone heading, then empty: the engine's keep anomaly (P0-4). 40–42: Word's
  5-page Appendix B table, ours shorter (P0-6). 43–64: the templates of
  Appendices D/E: blank lines of different styles at one offset (one style wins,
  already diagnosed), tables, line breaking.
- **bug59058.** No page matches. Its three 39-row tables take 3 Word pages each and
  a third of a page in ours (P0-6); its 18 pictures take no room in our flow (P0-2);
  its hidden form text (`w:vanish`) shows in ours and not in Word's (P1-2); Word's
  first pages hold no body paragraph the alignment can name (tables only).
- **Bug50936_3.** Pages 1–2 and 6–7 match. 3–5 (TOC): −3…−6. 8–13: the org chart
  and two floating pictures (Word pages 12, 26, 30) take no room in ours (P0-2),
  multi-level heading numbers ("3.10") come out as "1." (P1-3), tables with
  multi-paragraph cells overlap their rows (P0-6); from page 13 on, ours is 40–130
  paragraphs ahead.

## What is left, in priority order

**P0, engine (core RFIs; plugin-doc cannot fix them):**

1. *The font table is built once, at load* (`paged-canvas` `CanvasModel::font_table`,
   "the registry only changes at loadDocument boundaries"). Families a mutation
   introduces after the load (every family of a poured Word document; the
   skeleton's stories are empty) never resolve: everything lays out in the
   default face, whatever fonts are registered. Measured: setting a story range's
   family to Inter or to a registered Times New Roman leaves its geometry
   identical; a size change moves it. Rebuild the table when a mutation brings in a
   new (family, style) key.
2. *An inline anchored frame does not take room in its line.* `insertAnchoredFrame`
   pictures are drawn hanging above their line and the text runs on underneath
   (measured on `real_docx_docx()` V01/F01: the 100 pt picture overlaps the two
   paragraphs above, A01 follows V01 at one line's pitch). Word grows the line to
   the picture. Every picture of a Word document costs its height in pagination.
3. *Line breaking.* The engine composes paragraphs with Knuth–Plass; Word breaks
   first-fit, line by line. On the same faces and measure they break the same
   paragraph differently (often a word earlier in ours), so paragraphs gain or
   lose a line and the drift adds up. An InDesign "single-line composer" mode (the
   lowering would select it for Word documents) would give Word's breaks.
4. *A keep anomaly on generated pages.* With widow control on (Word's default), the
   engine leaves pages with three lines of a continuing paragraph, or a lone
   keep-with-next heading, and empty space below (parentinvguid ours 19 and 23).
   Neither is a keep rule's answer.
5. *Footnotes.* No native construct: Word's pages lose the footnote area's lines,
   ours do not (already diagnosed; ADR 028).
6. *Table rows.* Rows ignore their cell paragraphs' space after (the lowering
   emits it), so 39-row tables come out a third of Word's height; rows with several
   paragraphs in a cell overlap. The lowering does not carry cell margins yet (P1-1).

**P1, plugin-doc:**

1. Cell margins (`w:tblCellMar`, `w:tcMar`; Word's default 5.4 pt left/right)
   onto `cellInset*`, which the wire can set.
2. Hidden text (`w:vanish`): shown today. Needs a native hidden condition on a
   range (the wire can toggle a condition, not apply one: an RFI first).
3. Multi-level numbering (`%1.%2` level text, start values): the marker is the
   level's own number only.
4. TOC hyperlink runs styled blue + underlined where Word draws them plain.
5. Floating VML shapes and text boxes (org charts), and floating DrawingML with
   other wraps: ADR 029 leaves floating drawings undecided.

**P2:** headers/footers placed on the page (no native header story yet; they do
not move the body, so pagination does not depend on them); blank lines of
different styles at one offset (diagnosed; a wire change).

#### parentinvguid — per page (after, Word's faces)

| page | Word's page starts in | ours | Δ ¶ |
|---:|---|---|---:|
| 1 | Parental Involvement: | Parental Involvement: | = |
| 2 | TABLE OF CONTENTS | TABLE OF CONTENTS | = |
| 3 | B-7. What are an SEA’s responsibilities for revi… | B-7. What are an SEA’s responsibilities for revi… | = |
| 4 | C-14. Do the parental involvement requirements o… | C-13. What funds must an LEA reserve for parenta… | -1 |
| 5 | D-3. What information do the parents’ “right-to-… | D-2. What notification and dissemination require… | -1 |
| 6 | Coordination with Other Programs and Community I… | E-9. May a school or all schools within a distri… | -1 |
| 7 | PARENTAL INVOLVEMENT | PARENTAL INVOLVEMENT | = |
| 8 | Three decades of research provide convincing evi… | This guidance is divided into five major section… | +1 |
| 9 | A. GENERAL INFORMATION | The purpose of this guidance is to assist SEAs, … | -2 |
| 10 | A-5. What does the research show about how famil… | Pass their classes, earn credits, and be promote… | +3 |
| 11 | A-9. What is meant by providing information to p… | This means that, whenever practicable, written t… | +1 |
| 12 | A-11. What Federal civil rights provisions are a… | In implementing parental involvement programs, a… | +1 |
| 13 | A-14. What other resources and research are avai… | Education News Parents Can Use, a television ser… | +2 |
| 14 | B. RESPONSIBILITIES OF STATES | In addition, each SEA must assure that it will p… | +4 |
| 15 | State report cards must include information rela… | A central requirement of the NCLB Act is that SE… | +2 |
| 16 | Throughout the school improvement process, the a… | An SEA must promptly notify the parents of each … | +3 |
| 17 | Yes.  An LEA may receive funds under Title I, Pa… | C-3. What specific information must an LEA’s wri… | +4 |
| 18 | Identifying barriers to greater participation by… | C-5. What other information related to parents m… | +6 |
| 19 | Similar to State report cards, LEA report cards … | be active participants in assisting their childr… | +4 |
| 20 | the methods of instruction used in the program i… | be active participants in assisting their childr… | -7 |
| 21 | In addition, if a language instruction education… | detailing the option that parents have a right t… | -6 |
| 22 | SEAs must adopt written procedures, consistent w… | Yes.  Under the equitable participation provisio… | -3 |
| 23 | An LEA reserves one and a half percent ($90,000)… | Yes.  LEAs with a Title I, Part A allocation of … | -6 |
| 24 | C-17. On what basis may an LEA distribute to sch… | Yes.  LEAs with a Title I, Part A allocation of … | -27 |
| 25 | In addition, an LEA must review and publicize th… | LEA’s total Title I allocation     $6,000,000 | -21 |
| 26 | what the school is doing to address the problem … | C-19. If an LEA reserves more than the required … | -20 |
| 27 | If a Title I school is identified for improvemen… | No.  The LEA may retain for district-wide parent… | -29 |
| 28 | Each school must develop, jointly with parents o… | may include the professional qualifications of t… | -28 |
| 29 | D-5. What meetings must schools hold to inform p… | An LEA is responsible for ensuring that technica… | -25 |
| 30 | Parent involvement is very important in a school… | An opportunity to participate in the development… | -27 |
| 31 | D-10. What information must a school provide to … | Schools served under Title I, Part A must involv… | -29 |
| 32 | The parental involvement requirements of section… | The planning, review, and improvement of the sch… | -42 |
| 33 | E-6. What school staff training must schools and… | The purpose of a schoolwide program is to improv… | -39 |
| 34 | E-9. May a school or all schools within a distri… | The importance of communication between teachers… | -37 |
| 35 | The Department encourages schools and LEAs to de… | It is the responsibility of schools and LEAs to … | -37 |
| 36 | Appendix A:  Definitions | It is the responsibility of schools and LEAs to … | -38 |
| 37 | The consistent academic failure of a school that… | E-5. Is volunteering in a child’s classroom an a… | -45 |
| 38 | The carrying out of other activities, such as th… | The Department strongly encourages parents to at… | -54 |
| 39 | Relies on measurements or observational methods … | The Department encourages schools and LEAs to de… | -53 |
| 40 | Appendix B:  Key Title I, Part A Parental Notice… | Results in continuous and substantial academic i… | -63 |
| 41 | Appendix B:  Key Title I, Part A Parental Notice… | Is consistent with State law. [Section 200.42(a)… | -45 |
| 42 | Appendix B:  Key Title I, Part A Parental Notice… | In the case of a school identified for school im… | -29 |
| 43 | Appendix B:  Key Title I, Part A Parental Notice… | Ensures that experimental studies are presented … | -14 |
| 44 | Appendix B:  Key Title I, Part A Parental Notice… | (table/picture only) | n/a |
| 45 | *This table includes key Title I, Part A statuto… | (table/picture only) | n/a |
| 46 | Appendix C:  Research Based Resources | (table/picture only) | n/a |
| 47 | The authors analyzed 41 studies that evaluated K… | *This table includes key Title I, Part A statuto… | -10 |
| 48 | This booklet contains a short summary of what sc… | The following resources represent a sample of th… | -12 |
| 49 | *Voorhis, V. and Frances, L. (2001) Interactive … | The following resources represent a sample of th… | -22 |
| 50 | Reaching and Involving Diverse Parents | The authors analyzed 41 studies that evaluated K… | -25 |
| 51 | Appendix D:  District Wide Parental Involvement … | This brochure was published by the Partnership f… | -23 |
| 52 | Consistent with section 1118, the school distric… | This article describes the results of a study on… | -27 |
| 53 | PART II. DESCRIPTION OF HOW DISTRICT WILL IMPLEM… | This study offers some areas for consideration b… | -30 |
| 54 | 5. The _name of school district_ will take the f… | NOTE:  In support of strengthening student acade… | -43 |
| 55 | D.  The school district will, to the extent feas… | In carrying out the Title I, Part A parental inv… | -58 |
| 56 | developing appropriate roles for community-based… | [NOTE:  The District wide Parental Involvement P… | -64 |
| 57 | Appendix E:  School-Parent Compact | The ___name of school district___________ will b… | -60 |
| 58 | School Responsibilities | E.  The school district will take the following … | -55 |
| 59 | Staying informed about my child’s education and … | This policy was adopted by the __name of school … | -66 |
| 60 | Hold an annual meeting to inform parents of the … | NOTE:   Each school receiving funds under Title … | -76 |
| 61 | School    Parent(s)   Student | Hold parent-teacher conferences (at least annual… | -83 |

#### bug59058 — per page (after, Word's faces)

| page | Word's page starts in | ours | Δ ¶ |
|---:|---|---|---:|
| 1 | (no body paragraph: Number of New Cases and Deaths…) | Top of Form | n/a |
| 2 | (no body paragraph: Number of New Cases and Deaths…) | (table/picture only) | n/a |
| 3 | (no body paragraph: Number of New Cases and Deaths…) | (table/picture only) | n/a |
| 4 | (no body paragraph: Figure 1.5: Estimated Breast C…) | (table/picture only) | n/a |
| 5 | (no body paragraph: Idaho) | Top of Form | n/a |
| 6 | (no body paragraph: EPIDEMIOLOGY) | Polychlorinated biphenyls (PCBs) are synthetic c… | n/a |
| 7 | (no body paragraph: Ste. B, Berkeley, CA 94709, US…) | For approximately 50 years (beginning in 1929), … | n/a |
| 8 | (no body paragraph: verify cases, comparing ﬁxed v…) | PCBs appear to have a number of toxic qualities … | n/a |
| 9 | (no body paragraph: randomly assigned the order of…) | Although there are some differences in the speci… | n/a |
| 10 | (no body paragraph: the PCB score because the scor…) | Given a reasonable basis for a conclusion of gen… | n/a |
| 11 | (no body paragraph: and risk of breast cancer (dat…) | For the evaluation of PCB exposure and NHL, the … | n/a |
| 12 | (no body paragraph: result (Table 4), nor did adju…) | For the purposes of the present study, a causal … | n/a |
| 13 | (no body paragraph: (1.11, 7.09)) | In order to assess the utility of PCB congener l… | n/a |
| 14 | (no body paragraph: explanations for this ﬁnding, …) | A test for homogeneity was also conducted to ass… | n/a |
| 15 | (no body paragraph: positive values that represent…) | In order to explicate the correlation between ch… | n/a |
| 16 | (no body paragraph: distribution of PCB congeners …) | Consistency of the relationship is seen with the… | n/a |
| 17 | (no body paragraph: congener mixtures and individu…) | The weight-adjusted odds ratio (ORMH) for all 10… | n/a |
| 18 | (no body paragraph: during pregnancy, and in adult…) | Figure 3: Ecological correlations between PCB ac… | n/a |
| 19 | (no body paragraph: Long Island. II. Organochlorin…) | Given this conclusion, the issue of the >2.0 rel… | n/a |
| 20 | (no body paragraph: Services, Public Health Ser- v…) | Taken together, all of these factors indicate th… | n/a |
| 21 | (no body paragraph: Number of New Cases and Deaths…) | National Toxicology Program (NTP), Report on Car… | n/a |
| 22 | (no body paragraph: Number of New Cases and Deaths…) | L. P. Hanrahan, C. Falk, H. A. Anderson et al., … | n/a |
| 23 | (no body paragraph: Number of New Cases and Deaths…) | G. Maifredi, F. Donato, M. Magoni et al., “Polyc… | n/a |
| 24 | Journal of Environmental and Public Health Volum… | M. Merhi, H. Raynal, E. Cahuzac, F. Vinson, J. P… | +103 |
| 25 | Polychlorinated biphenyls (PCBs) are synthetic c… | E. N. Marieb, Human Anatomy and Physiology, Addi… | +95 |
| 26 | PCBs are deemed to be probable carcinogens by th… | S. A. Kafafi, H. Y. Afeefy, A. H. Ali, H. K. Sai… | +99 |
| 27 | The underlying physiologic mechanisms for the de… | K. Hardell, M. Carlberg, L. Hardell et al., “Con… | +104 |
| 28 | A few authors have reported subgroups of NHL tha… | A. M. Ruder, M. J. Hein, N. Nilsen et al., “Mort… | +105 |
| 29 | For the evaluation of PCB exposure and NHL, the … | R. van Reekum, D. L. Streiner, and D. K. Conn, “… | +114 |
| 30 | For the purposes of the present study, a causal … | F. Vinson, M. Merhi, I. Baldi, H. Raynal, and L.… | +118 |
| 31 | In order to assess the utility of PCB congener l… | Béatrice Lauby-Secretan , Dana Loomis , Yann Gro… | +156 |
| 32 | In order to explicate the correlation between ch… | Béatrice Lauby-Secretan , Dana Loomis , Yann Gro… | +151 |
| 33 | In examining the plausibility of a causal relati… | PCBs are a class of aromatic compounds comprisin… | +147 |
| 34 | The weight-adjusted odds ratio (ORMH) for all 10… | PCBs can compromise the immune surveillance mech… | +141 |
| 35 | Figure 3: Ecological correlations between PCB ac… | Overall, all PCBs can induce formation of reacti… | +137 |
| 36 | Given this conclusion, the issue of the >2.0 rel… | 3 Loomis D, Browning SR, Schenck AP, et al. Canc… | +138 |
| 37 | Disclosure | (no page) | n/a |
| 38 | B. L. Johnson, H. E. Hicks, W. Cibulas et al., P… | (no page) | n/a |
| 39 | R. Recio-Vega, V. Velazco-Rodriguez, G. Ocampo-G… | (no page) | n/a |
| 40 | M. Merhi, H. Raynal, E. Cahuzac, F. Vinson, J. P… | (no page) | n/a |
| 41 | N. Tijet, P. C. Boutros, I. D. Moffat, A. B. Oke… | (no page) | n/a |
| 42 | K. A. Bertrand, D. Spiegelman, J. C. Aster et al… | (no page) | n/a |
| 43 | M. D. Freeman and S. S. Kohles, “An evaluation o… | (no page) | n/a |
| 44 | M. D. Freeman, C. J. Centeno, and S. S. Kohles, … | (no page) | n/a |
| 45 | D. Baris, L. W. Kwak, N. Rothman et al., “Blood … | (no page) | n/a |
| 46 | D. Baris, L. W. Kwak, N. Rothman et al., “Blood … | (no page) | n/a |
| 47 | D. Baris, L. W. Kwak, N. Rothman et al., “Blood … | (no page) | n/a |
| 48 | New Cases, Deaths and 5-Year Relative Survival V… | (no page) | n/a |
| 49 | Béatrice Lauby-Secretan , Dana Loomis , Yann Gro… | (no page) | n/a |
| 50 | Individual PCBs activate numerous receptors, inc… | (no page) | n/a |
| 51 | The carcinogenicity of PCBs in animals has been … | (no page) | n/a |
| 52 | 3 Loomis D, Browning SR, Schenck AP, et al. Canc… | (no page) | n/a |

#### Bug50936_3 — per page (after, Word's faces)

| page | Word's page starts in | ours | Δ ¶ |
|---:|---|---|---:|
| 1 | Développement d’un moteur de recherche d’informa… | Développement d’un moteur de recherche d’informa… | = |
| 2 | (Rapport de stage) | (Rapport de stage) | = |
| 3 | 7. Réalisation des filtres pour Windex 23 | 6.4 Moyens mis à disposition 21 | -3 |
| 4 | 12.3 Structure des documents Microsoft Word 45 | 9.3 Orientations techniques souhaitées 39 | -6 |
| 5 | 14. Annexes 50 | 13. Matériel 48 | -1 |
| 6 | My project was a wonderful success and I am glad… | My project was a wonderful success and I am glad… | = |
| 7 | Ce rapport relate, à travers l’enquête entrepris… | Ce rapport relate, à travers l’enquête entrepris… | = |
| 8 | De plus, nous le répétons encore ici, elle fonct… | Activités | +12 |
| 9 | La puissance de l'indexation : | Création des catégories de recherche. | +8 |
| 10 | Un article au sujet de Windex 2 est paru dans le… | Cela permettra à Windex de ne plus se contenter … | +9 |
| 11 | Application de filtres sur l'image sélectionnée | Si ce document doit être scanné, IndexGED permet… | +15 |
| 12 | IIS : Indexateur interne à Windows NT (pour l'In… | Jimagepro (Version client-serveur de Jimage). | +18 |
| 13 | Jimage | Multimédia Solutions est une entreprise qui aime… | +53 |
| 14 | A ce jour, l’entreprise n’a fait aucun emprunt b… | Diplômée d'un BTS secrétariat (1982), puis du CE… | +27 |
| 15 | La prévision des achats se fait principalement e… | Marc Mendez a suivi une formation à l’université… | +40 |
| 16 | Si Hélène est la fonction vitale de Multimédia S… | Le premier objectif de ce projet est la réalisat… | +43 |
| 17 | Enfin, Marc a un avis sur tout, ce qui est vérit… | Besoin : En effet, l'indexeur de Windex, pour l'… | +39 |
| 18 | Service : L'idée est alors d'obtenir un programm… | Besoin : En effet, c'est l'efficacité de la rech… | +41 |
| 19 | Une fois Windex Server en état de fonctionnement… | sa fiabilité et sa robustesse, Java est très réc… | +45 |
| 20 | Service : Une procédure de sauvegarde devra alor… | L'aspect de l'interface Internet se doit de rete… | +51 |
| 21 | De plus, Windex doit tenir compte des évolutions… | MozzleStd, un logiciel de vérification de dispon… | +56 |
| 22 | L'aspect de l'interface Internet se doit de rete… | L'une des demandes principales des clients de Mu… | +65 |
| 23 | Aussi, deux parties distinctes découpe mon proje… | Word 5.1 pour Macintosh | +46 |
| 24 | On peut d’ailleurs voir l’exemple de ce rapport … | Début Novembre s’apprêtait à sortir une nouvelle… | +49 |
| 25 | Voici donc les différentes classes et les versio… | La lecture de ces différents codes sur 9, 10, 11… | +48 |
| 26 | Depuis le 6 novembre, une nouvelle version de Wi… | Il s'est de plus trouvé qu'au même moment, un cl… | +64 |
| 27 | La réalisation de ce filtre a été moins globale … | Lors de la réalisation du cahier des charges, il… | +65 |
| 28 | Le filtre Ascii85Decode : | Comment on le met en place sur le serveur ? | +38 |
| 29 | En effet, après une brève analyse, il m'est appa… | Orange : de 1000 à 2000 sites. | +48 |
| 30 | La rubrique | L’interface Web est l’élément visible de l’icebe… | +49 |
| 31 | Violet : plus de 10000 sites. | Formulaire d'ajout d'URL | +87 |
| 32 | Le formulaire de recherche. | Un Agent est alors charger de vérifier l’arrivée… | +85 |
| 33 | L’accueil du site Web Naevis. | Il doit pouvoir gérer la base de données des Url… | +102 |
| 34 | On les transmet ensuite à Windex qui renverra le… | Une sauvergarde générale de tous mes travaux a é… | +133 |
| 35 | Voyons un peu le résumé de tout ça. | Une version de démonstration est disponible sur … | +107 |
| 36 | Un Agent est alors charger de vérifier l’arrivée… | Durant mon stage, Bruno Lacombe terminait la réa… | +119 |
| 37 | Il est bien sûr très intéressant d'obtenir des s… | 404 : Introuvable | +96 |
| 38 | Un programme particulier pourra être mis en plac… | DHTML : Dynamic HyperText Markup Language. Le DH… | +89 |
| 39 | Un programme particulier pourra être mis en plac… | JDK : Java Development Kit. Environnement de dév… | +115 |
| 40 | Il nous faudra éventuellement réaliser une petit… | Les navigateurs : | +117 |
| 41 | Cependant, par dessus tout, je voudrais remercie… | Opéra : http://www.opera.com/ | +109 |
| 42 | CSS : Cascading Style Sheet. Modèle de feuille d… | Altavista : Moteur d'indexation automatique de l… | +107 |
| 43 | JAR : Java ARchive. Archive compressé incluant t… | Windex est une marque déposée à l'INPI par Multi… | +135 |
| 44 | W3C : World Wide Web Consortium. Organisme qui c… | Un Macintosh présent chez des partenaires de l'e… | +122 |
| 45 | W3C : World Wide Web Consortium. Organisme qui c… | (table/picture only) | n/a |
| 46 | http://www.rasip.fer.hr/research/compress/algori… | (no page) | n/a |
| 47 | http://www.multimania.com/patderam/docu.htm | (no page) | n/a |
| 48 | Windex est une marque déposée à l'INPI par Multi… | (no page) | n/a |
| 49 | Voici la structure du réseau informatique de Mul… | (no page) | n/a |
| 50 | Voici la structure du réseau informatique de Mul… | (no page) | n/a |

## Round 2 — 2026-10-02, engine at core `protocol-65` (`9180b85`)

Re-measured after the engine fixes the first round asked for: fonts registered after
load (`133f19b`), inline pictures as characters of their line (`17d3d3d`), the
single-line composer (`fc2df8a`, set through `paragraphComposer`), widow control on
growing chains (`dd3bfe7`), table row heights (`2e3c998`, `bd2563a`). Same three
documents, same Word PDFs, `scripts/real-docx-acceptance.sh` (now three modes).

| Document | Word | Ours (none / word / word-reload) | Exact page start | Within ±1 paragraph | Median drift (paragraphs) |
|---|---|---|---|---|---|
| parentinvguid | 61 | 64 / 59 / 59 | 8.2% / 6.6% / 6.6% | 6 / 10 / 10 pages | 12.5 / 6 / 6 |
| bug59058 | 52 | 35 / 36 / 36 | 0% | 0 | 131 |
| Bug50936_3 | 50 | 45 / 45 / 45 | 8% | 5 pages | 43 / 45.5 / 45.5 |

**Still not met.** What the round shows:
- `word` and `word-reload` are identical: late fonts work, the reload workaround is gone.
- parentinvguid's page count closed in (75/64 → 64/59 against Word's 61) and its first
  nine pages start within 0–2 paragraphs of Word's (0, 0, 0, −1, −1, −1, 0, +1, −2); the
  slips then accumulate. Exact page-start match punishes any one-line difference on every
  later page, so the number does not move until line breaking and line heights are exact.
- bug59058 is unchanged at 36 of 52 pages: its tables are still far shorter than Word's,
  and the comparison cannot align its page starts at all.
- Bug50936_3 drifts from page 3 (the TOC) and from page 8 on.

Next, in order of pages affected: bug59058's table heights (cell margins, row heights,
pictures in cells — which the pour never places); the TOC pages (tab leaders and right
tabs against Word's line positions); then per-line agreement with Word on body text
(font metrics of the faces Word actually used, justification, hyphenation off), footnotes
(ADR 034) and floating drawings (ADR 035).

## Round 3 — 2026-10-02: line by line, and side by side

Engine at core `protocol-65` rebased on main `1663863`. Same three documents, `word`
mode, plus `REAL_TIMES_DIR` (the macOS "Times" faces extracted from `Times.ttc`, which
bug59058 asks for 1,211 times and round 2 never registered). The page-start numbers did
not move (6.6% / 0% / 8%), so this round measured LINES per paragraph against Word's PDF
(`linecmp.py` in the session scratch; worth promoting to `scripts/real-docx/`) and
compared pages side by side.

| Document | Word lines | Ours (round 2 → 3) | Paragraphs with equal line count |
|---|---|---|---|
| parentinvguid | 2257 | 2333 → 2335 | 230 of 624 |
| bug59058 | 1082 | 1192 → 1092 | 40 of 173 |
| Bug50936_3 | 1541 | 2002 → 2002 | 261 of 813 |

What the side-by-side pages show (each verified on a rendered page, not inferred):
1. **A missing font is the largest single error.** bug59058's body is set in "Times";
   unregistered it fell to Inter and every paragraph gained ~22% lines. With the face
   registered its line total is within 1% of Word's. The editor needs the document's
   faces, or a metric-compatible substitution table; the lowering reports unmeasured
   fonts but nothing tells the user a face is missing.
2. **Kerning and ligatures**: Word sets text without either; the engine's defaults are
   InDesign's. Now lowered on the base style (`characterKerningMethod = None`,
   `characterLigatures = false`; needs core `1663863`). Small effect, as measured.
3. **Break after a hard hyphen**: Word breaks `two-` / `way`; the engine never breaks at
   an existing hyphen. Engine change (the composer already segments words at soft
   hyphens); needs InDesign's rule first.
4. **Headers and footers take body room.** Bug50936_3's header is taller than its top
   margin and pushes the body down; its footer band does the same. Nothing is placed
   (ADR 033), so each of our pages holds more than Word's.
5. **Footnotes take body room** (parentinvguid page 9: Word's body ends at A-4 above a
   footnote, ours runs on to A-5). ADR 034.
6. **Table rows are too short.** InDesign's cell ends at its last baseline; Word gives
   every line its full line box plus paragraph spacing and cell margins, and can centre
   vertically. Cell margins (`w:tblCellMar` / `w:tcMar`) are not imported at all and no
   cell inset is set. Lowering fix: cell insets from the margins, the first paragraph's
   space before, and the last line's descent + the last paragraph's space after.
7. **A picture alone in a paragraph lands on a neighbouring line** and covers its text
   (parentinvguid's cover seal). An empty paragraph cannot be addressed by offset, so the
   anchor falls to the next paragraph. Floating pictures are set inline (ADR 035).
8. **Runs of blank lines keep the default line height**: the engine refuses a caret that
   stands for several empty paragraphs with different styles (13 in a row in bug59058).
9. **Heading numbers**: Word's multi-level `3.5.1`; ours restarts at `1.`.

`cargo run -p docx-conformance --example dump-lowered -- <file.docx>` prints a document's
lowering as JSON (styles with their chains, blocks, sections) for this kind of check.

### After round 3: table geometry (2026-10-02)

Two lowering fixes from the side-by-side pages, measured on the same engine:

| bug59058 | Pages (Word 52) | Median page-start drift |
|---|---|---|
| round 3 | 33 | 163 paragraphs |
| + cell margins, line box and vertical alignment as insets (`35d8fbc`) | 39 | 110 |
| + repeating header rows (`w:tblHeader`) | 40 | 106 |

Page 1 of its first table now ends on the same row as Word's. The other two documents do
not move (their gaps are headers/footers, footnotes, pictures and heading numbers).
