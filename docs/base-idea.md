# paged.doc — base idea (Word / DOCX as a first-class content type)

**Status: DRAFT concept — not ratified. Records a direction for the empty `plugin-doc/`
forward slot, for review before any code.** · Date: 2026-07-23 · Effort to record: L.

**Sources (platform truth this design binds to):**
`README.md` / `CLAUDE.md` (repo map; plugins add content types on the shared SDK);
[ADR-021](../adr/021-paged-native-document-model-idml-as-format.md) (Paged-native model; `.paged`
is a *container of native parts* + one core-owned *composition* part; each content engine stores its
own native format; IDML byte-perfect as a native part) and its addendum (multi-surface renderer/
compositor; the "native document model" is the *composition* model, small and clean);
[ADR-022](../adr/022-idml-relocates-to-plugin-publish.md) (format adapters live in plugin repos — `idml-import`/`idml-export` in `plugin-publish`);
[ADR-017](../adr/017-importer-exporter-door-shape.md) (K-2 importer door: resolve-by-extension **before** the default IDML load; `import(bytes)` loads into the plugin's own engine and **replaces the load target**; the clean "open into a fresh host document" door — `host.document.open(bytes)` — **does not exist yet**);
[ADR-020](../adr/020-paged-web-native-engine-defer-frame-threading.md) (the renderer-neutral `FlowId`/region-chain/overset seam "shared with IDML stories"; W-frag rungs 1–2 proven for web);
[ADR-007](../adr/007-carry-through-rendering-honesty.md) (carry-through + render honesty);
`core/file-format.md` (the `.paged` ZIP/OPC container; unknown parts round-trip untouched — "the single most important guarantee"; the three-role spec/source/derived storage model);
`plugin-sheets/base-idea.md` (the closest engine-plugin precedent: own Rust crates, XLSX **preservation invariant**, dual compiled-`<Table>`/in-frame-C-1 surfaces, K-1 editing);
`plugin-web/base-idea.md` + `web-render/src/{lib,flow}.rs` (foreign flowing-document model; Blitz→C-1; fragmentation as the open hard axis);
`plugin-sdk/packages/plugin-api/src/{host,editor,manifest}.ts` (the plugin contract);
`core/crates/paged-model`, `paged-text`, `paged-scene`, `paged-flow`, `paged-composition`, `paged-compose`, `paged-mutate` (the native text + flow + compose stack);
[ooxmlsdk](https://github.com/KaiserY/ooxmlsdk) (Rust Open XML SDK — docx/xlsx/pptx, code-generated typed part trees; v0.5.1, 2026-04; MIT OR Apache-2.0).

---

## 1. The decision, in one paragraph

`paged.doc` (npm/bundle id `media.paged.doc`) makes Microsoft Word **`.docx`** a first-class Paged
content type, in **two symmetric modes**: **(a) embedded** — a Word document placed inside a host
Paged layout as a content object in a frame (or frame chain); and **(b) standalone** — open a whole
multi-page `.docx` so that *it is* the canvas and edit it page-by-page with the full editor. The
central bet: **because WordprocessingML is, structurally, the same "styled paragraphs of runs flowing
across pages" model the engine was built for, `paged.doc` lowers DOCX onto the engine's *native* text
+ flow stack (`paged-model` `Story`/`Paragraph`/`CharacterRun` → `paged-text` Knuth–Plass → `paged-flow`
region chains), not onto a forked in-plugin layout engine.** DOCX is parsed and re-serialized through a
new **shared `paged-ooxml` foundation** (OPC container + typed part DOM over the `ooxmlsdk` crate + a
preservation invariant), reused by `paged.sheet` and the future `paged.slide` because all three are the
same ECMA-376 family. The two open platform doors this needs — `host.document.open(bytes)` and a
whole-document export read — are filed as RFI items; native lowering **avoids** the two gaps that would
otherwise bite (foreign-flow fragmentation *W-frag*, and the missing caret door *C-9*).

---

## 2. Two modes, and the generalization to every content plugin

The user requirement is explicit and it generalizes: **"the same should be possible for all other
plugins — either their content is embedded, or they can open that content type alone."** This is not a
`paged.doc` feature; it is a platform pattern `paged.doc` should be designed to *instantiate cleanly*,
so sheet/slide/web/image/data inherit it.

| Mode | What it is | Platform mechanism (today) | The gap |
|---|---|---|---|
| **Embedded** | A `.docx` lives in a **frame** of a host `.paged` document, alongside IDML text, images, sheets, etc. | `objectType` + `editContext` (K-1) claim a frame by its metadata; content is either **lowered to native** content in that frame's story/frame-chain, or painted via **C-1 `SceneLayer`**. The doc's bytes persist as a **native part** via `host.parts` (`paged/media.paged.doc/<id>/…`). | None blocking — this is the sheets/web pattern. |
| **Standalone** | Open a `.docx` and **it becomes the document**: a full composition of pages you edit natively. | K-2 importer (ADR-017) resolves `.docx` by extension **before** the IDML loader and routes bytes to `paged.doc.import()`. | **`host.document.open(bytes)` does not exist.** ADR-017 ships *registration + routing* but explicitly defers the "open into a fresh host document" door. **This is the one door `paged.doc` must land** (RFI **DOC-01**, §9). |

**The generalized contract (proposed platform pattern — RFI DOC-05).** Every content plugin declares,
in its manifest, up to two faces of the same engine:

- an **embed contribution** — `objectType` + `editContext` (+ `sceneLayer` and/or native lowering),
  the content-in-a-frame face; and
- an optional **document contribution** — `importer`/`exporter` + the new
  `capabilities.document.openNative` door, the content-as-the-whole-canvas face.

Both faces are the *same* engine and the *same* native part on disk; they differ only in **placement**
— whether the content occupies one region of a host composition or **is** the composition. ADR-021's
addendum already frames the document as a *multi-surface compositor* over content-type engines, so
"embed vs. open" is a placement choice the composition model expresses (a surface/region in a host
composition, vs. the root composition). `paged.doc` is the cleanest first proof of this symmetry
because — unlike a spreadsheet grid or a web frame — a Word document is *natively a paginated document*,
so the standalone mode is the content type's most natural form.

---

## 3. Where `paged.doc` sits — the three precedents it inherits from

1. **`plugin-sheets` — the engine-plugin shape.** Own Rust crates compiling to **one** wasm module; a
   thin TS `*-host-model` doing pure `Lowered → Mutation[]` translation + a `*-bundle` (manifest +
   `activate`); a **preservation invariant** ("Paged never destroys a workbook" — unknown parts kept
   byte-identical, understood-but-untouched parts re-emitted from original bytes); and **two surfaces**
   — compiled to native content (a real `<Table>` via `insertTable`, with live multi-frame pagination
   across the host frame chain, `sheet-lower/paginate.rs` + `frameChain()`), *plus* an in-frame C-1 grid
   with K-1 modal editing. **`paged.doc` copies this shape almost verbatim** — with the crucial
   difference that its "compiled to native content" is not a table, it is **native stories + flow**,
   which the engine already paginates.

2. **`plugin-web` — the foreign flowing-document precedent (and the cautionary tale).** `paged.web`
   proves the platform can host a *foreign document model* and render it print-correctly by lowering an
   in-plugin engine (Blitz/Stylo/Taffy/Parley) to a **C-1 `SceneLayer`**. Its unsolved-hard axis is
   **fragmenting one flow across a linked frame chain** (W-frag; rungs 1–2 proven, a forked engine held
   as last resort). **`paged.doc` deliberately does *not* repeat this** — see §4: Word's model already
   *is* the engine's native model, so there is nothing to fork and no foreign fragmentation to solve.

3. **`plugin-publish` / `idml-import` + `idml-export` — the format-adapter packaging pattern.** ADR-022
   put IDML/PDF I/O in a plugin repo as `*-import`/`*-export` crates; `idml-export` does **verbatim
   carry-through** (untouched entries byte-identical + streaming attribute patches). **`paged.doc`'s
   `docx-import`/`docx-export` are the direct analog**, and the carry-through discipline is exactly the
   DOCX round-trip-fidelity model.

**The native target already exists.** `paged-model` owns `Story { paragraphs }`,
`Paragraph { paragraph_style, justification, indents, space_before/after, tab_list, bullets… }`,
`CharacterRun { text, character_style, font, point_size, fill_color, tracking, underline, baseline_shift… }`,
plus `ParagraphStyleDef`/`CharacterStyleDef`/`ObjectStyleDef`/`TableStyleDef`/`StyleSheet`. `paged-text`
does rustybuzz shaping, **Knuth–Plass** line breaking, TeX hyphenation, and frame-bound composition
with justification and drop caps. `paged-flow` is the **content-agnostic region-chain flow protocol**
(the `FlowId` seam ADR-020/021 bless as the one core-resident neutral concept). This is a
WordprocessingML-shaped stack that predates the Word use case.

---

## 4. The core bet — lower DOCX to native, don't fork an engine

### 4.1 Why native lowering is right *for Word specifically*

The platform has two proven strategies for foreign content: **lower-to-native** (sheets compiling to a
native `<Table>`; the engine owns layout) and **foreign-engine-into-C-1** (web running Blitz; the
plugin owns layout). The right choice is dictated by **how far the foreign model is from the engine's
native model:**

- **Web is far.** Arbitrary CSS box/flow/float/grid layout has no native equivalent → Blitz is
  justified, and fragmentation is genuinely novel work.
- **A spreadsheet is orthogonal.** Grid + formula semantics aren't a layout the text engine has → its
  own engine, with a native `<Table>` as the *lowered* publication form.
- **Word is near — nearest of all.** WordprocessingML is styled paragraphs and runs, paragraph/character
  styles with inheritance, sections with page geometry, tables, floating/inline images, lists, tabs,
  footnotes — **the same concept set the engine implements for IDML.** The distance from `w:p`/`w:r`/
  `w:pPr`/`w:rPr`/`w:sectPr` to `Paragraph`/`CharacterRun`/`ParagraphStyleDef`/section geometry is a
  *mapping*, not a *layout engine*.

**Therefore DOCX lowers onto the native stack**, and the engine's existing Knuth–Plass pagination flows
it across pages. This is the sheets "compiled surface," applied to the content type the engine is *most*
native to.

### 4.2 What native lowering buys — including two platform gaps it sidesteps

- **Real editing, real caret, for free.** Standalone Word editing uses the **native text tools** on
  native stories — the same caret, selection, reflow, styles panel, and undo the editor already has.
  This **sidesteps C-9** (the missing `host.text.caret()` door): a plugin `editContext` has no way to
  read the user's caret offset, but native text frames don't need one because the *editor* owns the
  caret. A foreign-engine approach would hit C-9 head-on.
- **Pagination is the engine's, so there is no W-frag.** Flowing a story across an auto-generated page/
  region chain is what `paged-flow` + composition already do for IDML stories. The hard "fragment a
  *foreign* flow across frames" problem only exists when the flow is computed by a foreign engine (web).
  Lowered to native, DOCX body text uses the **native** flow — already fragmenting across frames/pages.
- **One text/shaping stack, print-correct by construction.** No second font/shaping stack (ADR-020's
  objection to the web fork), full colour management, CPU-testable fidelity lane — all inherited.
- **Cross-content semantics for free.** Native stories participate in find/replace, tagged-PDF/
  accessibility export, and whole-document single-format export (ADR-021's "derived semantic view")
  without `paged.doc` doing anything special.

### 4.3 The honest cost, and how it's paid

The price of native lowering is that **edits happen on the native model, so saving back to `.docx`
requires projecting native → WordprocessingML and preserving everything untouched.** This is the same
bargain `plugin-sheets` already makes and the same discipline `idml-export` already runs:

- **Preservation invariant (DOC — mirror of the sheet invariant).** The original `.docx` OPC package is
  retained (as the content-engine's native part). Unknown parts and unknown subtrees of known parts are
  kept **byte-identical**; understood-but-untouched parts re-emit their **original bytes** (lazy-
  verbatim); a zero-edit round-trip is byte-identical modulo normalization noise. Only parts the user
  actually changed are re-serialized from the model.
- **Provenance-carrying lowering.** Each native paragraph/run/table/style produced from DOCX carries a
  back-reference to its source OOXML node (an `x-paged:media.paged.doc` binding envelope, IDML-carried
  in-session via `setMetadata`), so save-back is a **targeted patch** of the parts that changed, not a
  from-scratch re-serialization of the whole document.
- **Honest diagnostics (ADR-007).** Native-only edits with no clean WordprocessingML expression are
  reported, not silently dropped or silently mangled; features `paged.doc` cannot yet round-trip are
  surfaced as diagnostics on open (§7 fidelity tiers), never faked.

### 4.4 The alternative, documented honestly

| | **Y — lower to native** *(recommended)* | **X — foreign engine → C-1** (the web pattern) |
|---|---|---|
| Layout owner | Core `paged-text` (Knuth–Plass, existing) | An in-plugin Word layout engine (to be built/forked) |
| Multi-page | Native `paged-flow` pagination — **no W-frag** | Requires foreign-flow fragmentation — **W-frag** |
| Caret / editing | **Native text tools**, native caret — **no C-9** | Plugin `editContext` — **needs C-9** (missing) |
| Text/shaping stack | One (core rustybuzz) | Two (plugin's + core's) — ADR-020's objection |
| Round-trip | Native↔OOXML mapping + preservation (the cost) | Foreign model is the source → round-trip is easier |
| Word-exact line breaking | Convergence work (Word ≠ Knuth–Plass; §7 Tier 4) | Whatever the forked engine does |
| Fidelity/effort | High reuse; effort is the **mapping + save-back** | Low reuse; effort is an **engine + fragmentation** |

X's *only* structural advantage is round-trip ease (the foreign model stays the source of truth). Y pays
for that with the mapping/save-back layer but wins on editing quality, reuse, and by dodging two open
platform gaps. **The recommendation is Y**, with X retained as the pre-authorized escalation *iff* a
measured fidelity requirement proves the native mapping cannot express Word layout faithfully enough —
the same "revisit only on that evidence" posture ADR-011/020 hold for web. The most likely real-world
middle path is **hybrid**: lower the 95% that maps cleanly to native, and fall back to a **fixed,
non-editable faithful render** (a placed PDF/PNG derived, §7) for pages containing features the mapping
doesn't yet reach — never a fake, always either native-editable or an honest flat rendering.

---

## 5. The OOXML foundation — `paged-ooxml` over `ooxmlsdk`, shared by doc / sheet / slide

### 5.1 Why a shared foundation

`.docx`, `.xlsx`, and `.pptx` are the **same ECMA-376 / ISO-IEC 29500 family**: the same **OPC**
(Open Packaging Conventions) ZIP container, the same `[Content_Types].xml` + `_rels/` relationship
graph, the same DrawingML for shapes/images/charts, the same theme/font/color primitives, the same
`mc:AlternateContent` markup-compatibility rules. Building three independent OOXML parsers is three
copies of the hard, boring, bug-prone part. **One shared foundation is the right factoring** — and it is
exactly what the user proposed by naming `ooxmlsdk`.

### 5.2 `ooxmlsdk` assessment (as of 2026-07)

| Dimension | Finding | Implication for us |
|---|---|---|
| **Coverage** | docx **and** xlsx **and** pptx; Office 2007→M365 namespaces incl. DrawingML/chart extensions | Serves all three target plugins — the whole point |
| **Approach** | **Code-generated typed part trees** from the OOXML schemas; package-level APIs (`main_document_part`, `workbook_part`, `presentation_part`); open/create/save packages | Gives us a typed DOM to map from/to — no hand-rolled XML |
| **Round-trip** | open + save packages; **byte-exact fidelity not explicitly documented** | **Do not rely on it for preservation** — keep our own OPC verbatim carry-through (§5.4) |
| **Serde** | **No serde**; custom generated ser/de | Fine (it's faster for XML); just an integration note |
| **"No Particle"** | XML children are flattened into **enum vectors**, not named struct fields (`OneSequence` supported) | Ergonomic cost: we **wrap** it in our own semantic model, we don't expose it raw to plugin logic |
| **MC processing** | `OpenSettings` / full markup-compatibility modes **not exposed** | We handle `mc:AlternateContent` fallback selection ourselves at the `paged-ooxml` layer |
| **wasm** | **No wasm claim anywhere** | **Gating spike (RFI DOC-02): confirm `ooxmlsdk` compiles to `wasm32-unknown-unknown` and runs in-browser.** Everything in Paged is wasm; this is a go/no-go for adopting it |
| **Maturity / bus factor** | v0.5.1 (2026-04); single maintainer, ~54 stars | **Vendor and pin it** (the DuckDB-WASM / Blitz posture): be ready to carry patches; treat upstream as a starting point, not a dependency we can't fork |
| **License** | **MIT OR Apache-2.0** | Compatible with an MPL-2.0 shared crate and AGPL plugins; no friction |

**Verdict: adopt `ooxmlsdk` as the low-level typed-DOM + (de)serialization engine, behind our own
`paged-ooxml` wrapper, gated on the wasm spike (DOC-02) and treated as vendored/pinned.** If the wasm
spike fails and can't be cheaply fixed, the fallback is a thinner in-house OPC+quick-xml layer for the
subset of parts we touch — more work, same architecture.

### 5.3 The layering (what is shared vs. what stays per-plugin)

The shared crate owns the **format-mechanical** layer; each plugin keeps its **semantic** model. This
respects the platform's isolation-superset rule (a plugin depends only on `@paged-media/plugin-api`,
`plugin-sdk`, and **published** contracts — never on another plugin), so `paged-ooxml` ships as its own
**published** crate/package, not a cross-plugin source dep.

```
          ┌──────────────────────────────────────────────┐
          │  paged-ooxml  (NEW, shared, published)        │   MPL-2.0 OR PMEL
          │  · OPC container: ZIP, [Content_Types].xml,   │
          │    _rels graph, part naming                   │
          │  · typed part DOM via ooxmlsdk (vendored)     │
          │  · mc:AlternateContent fallback selection     │
          │  · DrawingML / theme / font / color shared    │
          │  · PRESERVATION: verbatim carry-through of    │
          │    untouched + unknown parts (§5.4)           │
          └───────────────┬───────────────┬──────────────┘
                          │               │               │
        ┌─────────────────┘      ┌────────┘      ┌─────────┘
   docx-* (paged.doc)      xlsx via sheet-*   pptx-* (paged.slide, future)
   WordprocessingML →      SpreadsheetML →    PresentationML →
   native stories/flow     cell/formula model native slides/shapes
```

- **Shared (`paged-ooxml`):** everything ECMA-376-common — OPC packaging, relationships, content-types,
  DrawingML graphics, themes, the shared `sharedStrings`/`styles` *plumbing*, MC processing, and the
  **preservation/carry-through** machinery. This is where `ooxmlsdk` lives, wrapped.
- **Per-plugin (semantic):** WordprocessingML→native mapping is `paged.doc`'s; the cell/formula model is
  `paged.sheet`'s; PresentationML→slides is `paged.slide`'s. **No shared semantic model** — that's where
  the formats genuinely differ.

### 5.4 The `sheet-xlsx` consolidation — evolutionary, not big-bang

`plugin-sheets` already ships `sheet-xlsx` with its own preservation invariant and 1000+ tests. **Do not
rewrite it.** The consolidation path:

1. `paged-ooxml` is built **for the two new consumers first** (`paged.doc` now, `paged.slide` later),
   where there is no incumbent to disturb. It earns its shape on greenfield.
2. It is designed so `sheet-xlsx`'s **OPC + carry-through + DrawingML** responsibilities *could* be
   lifted onto it later, leaving `sheet-xlsx` as the thin SpreadsheetML-semantic layer — but only **when
   that migration is cheap and green**, opportunistically, never as a stop-the-world rewrite (the exact
   evolutionary discipline ADR-021 mandates for the model migration).
3. Until then, `sheet-xlsx` and `paged-ooxml` coexist. The win is immediate for doc/slide; the sheets
   consolidation is a later, optional dividend.

---

## 6. Architecture — crates, packages, and the wasm boundary

Mirrors the `plugin-sheets` layout (all Rust semantics; thin TS glue; one wasm module per bundle).

**Rust crates (a Cargo workspace in `plugin-doc/`):**

| Crate | Purpose |
|---|---|
| `docx-core` | Frozen types: the `paged.doc` view of a WordprocessingML document (sections, body, styles, numbering, headers/footers, notes) — a thin semantic layer over `paged-ooxml`'s typed DOM |
| `docx-import` | `paged-ooxml` package → `docx-core` (parse + MC resolution + style/numbering resolution) |
| `docx-lower` | **The heart.** Pure `docx-core` → native lowering: `Story`/`Paragraph`/`CharacterRun` + style defs + tables + placed images + a `paged-flow` region-chain + section→page-geometry. Includes `paginate.rs`-style page generation for standalone mode. Provenance back-refs attached here |
| `docx-export` | native (edited) + retained OOXML → `.docx` **save-back**: targeted patch of changed parts + verbatim carry-through of the rest (the preservation invariant) |
| `docx-render` *(optional, hybrid fallback only)* | Fixed faithful render of unsupported pages to PDF/PNG derived, when native lowering can't yet express a construct (§4.4) |
| `docx-js` | The single `wasm-bindgen` surface — all crates compile to **one** wasm module |
| `docx-conformance` | Test-only: corpus + fidelity harness (§8) |

**TS packages (`plugin-doc/packages/`):**

| Package | Purpose |
|---|---|
| `@paged-media/doc-host-model` | **Pure** `LoweredContent → Mutation[]` translation (no DOM) — the same role as `sheet-host-model`; turns the lowering into `insertText`/`applyStyle`/`createParagraphStyle`/`insertTable`/`placeImage`/`linkFrames` mutations |
| `@paged-media/doc-bundle` | `manifest.json` (`media.paged.doc`) + `activate(host)`: the importer/exporter, the `objectType` + `editContext`, the styles/outline panels, engine boot |

**Manifest sketch (`media.paged.doc`):**

```jsonc
{ "id": "media.paged.doc", "name": "paged.doc", "apiVersion": "^0.x",
  "capabilities": {
    "document": { "read": "broad", "write": "scoped",
                  "readNative": true, "openNative": true },   // openNative = standalone open (DOC-01)
    "rendering": ["sceneLayer"],                               // for the hybrid fixed-render fallback
    "editContext": ["wordDocument"],                           // embedded modal edit (or native, §8.1)
    "wasm": ["ooxml"], "assets": ["fonts"], "clipboard": "text" },
  "contributes": {
    "importers": [{ "id": "media.paged.doc.import", "extensions": [".docx", ".dotx"] }],
    "exporters": [{ "id": "media.paged.doc.export", "extension": ".docx" }],
    "editContexts": [{ "type": "wordDocument", "entry": "doubleClick" }],
    "objectTypes": [{ "type": "wordDocument", "bakedFallback": "group" }],
    "partTypes": [
      { "type": "docx", "role": "source", "format": "opc", "linkable": true },
      { "type": "docLowered", "role": "derived" }
    ] } }
```

**On-disk (three-role model, per `file-format.md`):**

```
paged/media.paged.doc/<id>/source.docx     source  (the retained OPC package — byte-preserved)
paged/media.paged.doc/<id>/bindings.json   spec    (native↔OOXML provenance map for save-back)
paged/media.paged.doc/<id>/preview.pdf     derived (IDML-visible flatten / hybrid fallback)
```

For **standalone** mode the same part set is the whole document's; the core **composition** part
(`document.pgd`, ADR-021) references the lowered flow across the generated pages.

---

## 7. The WordprocessingML → native mapping, in fidelity tiers

The mapping is the bulk of the work; staging it in tiers lets `paged.doc` ship value early and degrade
honestly (ADR-007) on what isn't mapped yet. Each tier is "native-editable"; anything past the current
tier falls to the hybrid fixed-render fallback (§4.4) with a diagnostic, never a silent drop.

- **Tier 0 — structure & direct formatting.** OPC parse; `styles.xml` (`docDefaults`, `basedOn`
  inheritance) → `ParagraphStyleDef`/`CharacterStyleDef`; `document.xml` body → `Story` of
  `Paragraph`/`CharacterRun`; character formatting (bold/italic/underline/strike/color/size/family/
  caps/tracking/baseline) → run attrs; paragraph formatting (alignment, indents, spacing, tabs,
  keep-with-next) → `Paragraph`. **Ships a readable, styled, single-section document.**
- **Tier 1 — flow & sections.** `sectPr` (page size, margins, columns) → section/page geometry;
  page/column breaks; **pagination across generated pages** (native `paged-flow`); `numbering.xml`
  (lists, multi-level) → native numbering; theme (`theme1.xml`) fonts/colors. **Ships a real multi-page
  document — the standalone case's minimum bar.**
- **Tier 2 — the structures.** Tables (`w:tbl` — grid, merged cells via `gridSpan`/`vMerge`, borders,
  shading, cell margins) → native `insertTable`; inline **and** floating images (DrawingML `wp:inline`/
  `wp:anchor`, text wrap) → placed images with wrap; hyperlinks; **headers/footers** (per section,
  first/even/odd) → master/region content. **Ships a document most business `.docx` files render
  faithfully.**
- **Tier 3 — the apparatus.** Footnotes/endnotes; fields (`PAGE`, `NUMPAGES`, `REF`, `TOC`,
  cross-references) → native fields where they exist, computed values otherwise; comments; **tracked
  changes** (`w:ins`/`w:del`/revisions) → native change-tracking if/when present, else accepted-view +
  preserved; text boxes/shapes (DrawingML); content controls (SDT); OMML math.
- **Tier 4 — fidelity convergence.** **Word's line-breaking and justification differ from Knuth–Plass**;
  achieving Word-faithful reflow (not just plausible reflow) is a calibration problem analogous to the
  InDesign pixel-diff gate — likely a `paged-text` compatibility mode, gated on a Word-reference corpus
  (§8). Font substitution/metrics parity; compatibility settings (`compatSetting`).

The lowering is not all-or-nothing per document: a document at Tier 2 with one Tier-3 construct renders
Tier-2-native with that construct shown via the fixed-render fallback and a diagnostic — the same
"never faked" honesty `plugin-web` holds.

---

## 8. The two modes in detail

### 8.1 Embedded — a Word document inside a host layout

- **Entry:** an insert command / drag-drop of a `.docx` onto a page creates a `wordDocument` object — a
  frame carrying `x-paged:media.paged.doc` metadata and the `source.docx` native part (`host.parts`).
- **Render:** `docx-lower` produces native content poured into the frame's story (and, if the content
  overflows, a `paged.doc`-managed **frame chain** via `linkFrames`, so an embedded Word doc can itself
  span several frames of the host layout). Constructs past the current tier → C-1 fixed render.
- **Edit:** double-click enters editing. **Preferred: native text editing** on the poured story (native
  caret/tools — no C-9), scoped by the `editContext` breadcrumb; the `editContext` exists mainly to
  bound the scope, own Cmd-Z coalescing (ADR-012), and offer doc-specific panels (styles, outline). A
  pure-C-1/K-1 modal surface (like the sheets grid) is the fallback for constructs edited as a unit.
- **Persist / round-trip:** edits update the native model and, through the provenance bindings, patch
  `source.docx` on save; untouched parts carry through verbatim.

### 8.2 Standalone — open a `.docx` as the whole canvas

- **Entry:** File→Open / drag-drop a `.docx`. The importer registry (ADR-017) resolves `.docx` **before**
  the IDML loader and routes bytes to `paged.doc.import()`.
- **The one new door — `host.document.open(bytes)` (RFI DOC-01).** ADR-017 ships routing but **not** the
  "open into a fresh host document" door. `paged.doc` needs it: `import()` must **create a new Paged
  document** whose root composition is the lowered Word flow across generated pages — not mutate the
  currently-open document. This is the platform's headline dependency for the standalone requirement, and
  it is the *same* door `plugin-image` needs for "PSD → new document" (ADR-017 names I-05), so it is a
  general platform door, not a `paged.doc` special case.
- **Render + edit:** the whole editor operates on the native composition — every native tool, panel,
  script, and the real caret. This is why native lowering matters most here: **standalone editing quality
  = native editing quality.**
- **Save:** `docx-export` writes back to `.docx` (targeted patch + carry-through). "Export to PDF/print"
  is the existing native path. The document's canonical on-disk form can be either the `.docx` itself
  (round-tripped) or a `.paged` container holding the `source.docx` native part — a UX choice (§9).

### 8.3 Whole-document export needs a second door (RFI DOC-03)

ADR-021's consequences note the exporter door "carries no whole-document read": today an `exporter()`
gets no handle to the entire document model. Faithful `.docx` **save-back of a standalone document**
needs a **whole-document-read** surface (or the exporter must read the retained native part + the live
mutation log). This is the second platform door `paged.doc` surfaces — shared with any
"export the whole composition to a foreign format" case (IDML export included).

---

## 9. Platform gaps this design opens (for the RFI register)

| ID | Gap | Why `paged.doc` needs it | Shared with |
|---|---|---|---|
| **DOC-01** | `host.document.open(bytes)` — open into a **fresh** host document | Standalone mode (§8.2); ADR-017 deferred it explicitly | plugin-image PSD (I-05), every standalone-open plugin |
| **DOC-02** | **`ooxmlsdk` wasm feasibility spike** | Everything runs in-browser wasm; no wasm claim upstream — go/no-go for §5 | paged.sheet consolidation, paged.slide |
| **DOC-03** | Whole-document **read** door for exporters | Faithful standalone `.docx` save-back (§8.3); ADR-021 names the gap | IDML export, all whole-doc foreign export |
| **DOC-04** | Fidelity-reference harness for Word (no headless Word) | Tier-4 convergence + regression gate (§8/§Fidelity); mirrors the "no InDesign on dev machine" problem | paged.slide (no headless PowerPoint) |
| **DOC-05** | **Ratify "embed-or-open" as a platform pattern** | Generalize §2 so every content plugin declares embed + optional document faces uniformly | web, sheet, image, data, slide |
| **DOC-06** *(only if hybrid)* | C-1 per-glyph faithful text (multi-line, per-run face) | The fixed-render fallback wants better than C-1's single-line/default-font text; else use PDF/PNG derived | plugin-web text paint follow-ons |

Note the pleasant result: the **native-lowering choice removes** what would otherwise be the two biggest
items — foreign-flow fragmentation (**W-frag**) and the caret door (**C-9**) — from `paged.doc`'s
critical path entirely.

**Fidelity gate.** Like `corpus/` gates IDML render against InDesign exports (ΔE2000 / SSIM), `paged.doc`
needs a Word-reference corpus. Word can't run headless in CI, so references come from **Word-exported
PDFs** (authored once) and/or **LibreOffice `--headless` PDF renders** as a continuous proxy, with the
"authored-once Word PDF" as the ground truth for a curated set. This is DOC-04.

---

## 10. Round-trip & preservation policy (binding to ADR-007 / ADR-021)

- **`.docx`-origin content is preserved by *storage*, not reconstruction.** The `source.docx` OPC package
  is retained as the content-engine's native part (ADR-021: "each content engine stores its own native
  format"). Untouched → byte-identical on save. This dissolves most of the round-trip-fidelity risk the
  same way ADR-021's addendum dissolves it for IDML: loss appears **only** when the whole composition is
  flattened to a *single* foreign format, and then it is inherent and **diagnosed**, never silent.
- **Edits patch, they don't rewrite.** Provenance bindings make save-back a targeted patch of changed
  parts (mirrors `idml-export`'s streaming attribute patches + `sheet-xlsx`'s preservation).
- **Never destroy a document** — the `plugin-sheets` invariant, verbatim, for Word: unknown parts and
  unknown subtrees survive re-emission untouched; a zero-edit round-trip is byte-identical modulo
  normalization.

---

## 11. Learnings that flow back to `plugin-web`

`paged.doc` is not just a consumer of platform work — done right, it feeds `plugin-web` on several axes:

1. **It proves and hardens the shared `FlowId`/region-chain seam.** ADR-020/021 bless one renderer-
   neutral flow model "shared with IDML stories"; `paged.doc` is the *second heavy consumer* of that seam
   (native Word flow across pages), which pressure-tests it beyond IDML before `plugin-web`'s W-frag
   rungs 3–4 ever need it. Every sharp edge `paged.doc` files against `paged-flow` is an edge `plugin-web`
   doesn't hit later. The seam graduating from "IDML + Word" to "IDML + Word + web flow" is the
   ADR-021 generalization made concrete.
2. **A hybrid "lower-the-easy-part to native" template for web.** `paged.doc`'s hybrid posture — lower
   the 95% that maps to native, fall back to a fixed render only for the hard remainder — is directly
   transferable: much of real-world HTML (headings, paragraphs, lists, simple tables, inline images) maps
   to the *same* native primitives as Word. `plugin-web` could lower that clean subset to **native
   stories** (getting native caret/reflow/editing — sidestepping C-9 for text-heavy web content) and
   reserve Blitz for genuinely CSS-hard layout, instead of routing *all* web content through Blitz→C-1.
   `paged.doc` validates the "native lowering of foreign flowing text" machinery web would reuse.
3. **The native-editing-dodges-C-9 result generalizes.** If `paged.doc` demonstrates that editing a
   foreign-origin document via native text frames is superior to a plugin `editContext` (no caret door,
   real tools), that reframes `plugin-web`'s editing story for its text content — and may lower the
   priority of building C-9 at all.
4. **Save-back / preservation discipline is shared.** The provenance-binding + verbatim-carry-through
   pattern `paged.doc` builds for OOXML is the same pattern `plugin-web` needs to round-trip authored
   HTML/CSS without reformatting it — a reusable "edit the native projection, patch the source"
   technique.
5. **The fidelity-reference method transfers.** DOC-04's "author ground-truth once (Word PDF) + a headless
   proxy (LibreOffice) for continuous CI" is a template for `plugin-web`'s WPT/print-reference gating.
6. **DOC-05 (embed-or-open) is a `plugin-web` feature too.** Ratifying the pattern gives `plugin-web`
   "open a standalone `.html` as the whole canvas" on the same door `paged.doc` lands (DOC-01).

Conversely, `paged.doc` should **borrow from `plugin-web`**: the `web-model` discipline (pure, zero-DOM,
"a scanner not a parser, never crashes on bad input") is exactly how `docx-import` should treat malformed
`.docx`; and W-frag's `flow.rs` rungs are the reference if `paged.doc` ever *does* need a foreign-engine
fallback (§4.4).

---

## 12. Milestones

- **M0 — Foundation + spike.** `paged-ooxml` skeleton (OPC read/write, `_rels`, content-types,
  carry-through) wrapping `ooxmlsdk`; **DOC-02 wasm spike** (go/no-go). Bundle skeleton, manifest,
  CLI-validated. Exit: a `.docx` opens, its parts round-trip byte-identical (no rendering yet).
- **M1 — Tier 0/1 standalone read.** `docx-import` + `docx-lower` Tier 0–1 → native stories + sections +
  pagination; **DOC-01** `host.document.open(bytes)`; open a real multi-page `.docx` and *see* it,
  natively, across pages (read-only acceptable). Exit: a business memo/report renders faithfully.
- **M2 — Native editing + save-back.** Native text editing on the opened document; `docx-export` Tier 0–1
  save-back with preservation; **DOC-03** whole-doc export read. Exit: open → edit text/styles → save
  `.docx`, zero-edit round-trip byte-identical, edited round-trip faithful.
- **M3 — Embedded mode.** `objectType` + `editContext` + `host.parts` persistence; place a `.docx` in a
  host layout, edit in place, round-trip. Exit: the embed case (§8.1) end-to-end; DOC-05 pattern drafted.
- **M4 — Tier 2 structures.** Tables, images (inline+floating+wrap), headers/footers, hyperlinks, lists.
  Exit: most real-world `.docx` files render + round-trip.
- **M5 — Tier 3 apparatus + Tier 4 fidelity.** Notes, fields, comments, tracked changes, text boxes;
  **DOC-04** Word-reference gate; `paged-text` Word-compat line-breaking mode. Exit: fidelity-gated.
- **Cross-cut:** `paged-ooxml` consolidation offered to `sheet-xlsx` (opportunistic, §5.4); `paged.slide`
  scoped on the same foundation.

## 13. Open questions

- **Canonical on-disk form for a standalone Word doc:** round-trip the `.docx` in place, or wrap it in a
  `.paged` container holding `source.docx`? (Likely: `.docx` stays canonical while purely Word content;
  promote to `.paged` when the user adds non-Word Paged content — a natural upgrade prompt.)
- **How much of Word's layout must be *Word-exact* vs. *faithful-plausible*?** Sets the Tier-4 bar and
  whether the hybrid fixed-render fallback is a long-term necessity or a temporary scaffold.
- **`ooxmlsdk` "enum-vector children" ergonomics at scale** — does the wrapper cost stay bounded, or does
  it argue for contributing named-particle codegen upstream (we're a plausible contributor given three
  consumers)?
- **`.dotx` templates** and Word template semantics — in scope as a variant of import, or a later theme?
- **Tracked changes & comments as native concepts** — does the native model grow first-class revision/
  comment support (useful well beyond Word), or does `paged.doc` preserve-and-project them?

## 14. Licensing

`paged-ooxml` is a **shared, embeddable library → MPL-2.0 OR PMEL** (like `plugin-sdk`/viewer). The
`paged.doc` plugin (crates + bundle) is **AGPL-3.0 OR PMEL**, like the other first-party plugins.
`ooxmlsdk` (MIT OR Apache-2.0) is compatible as a vendored dependency of the MPL crate.

## 15. One-line recommendation

> Build `paged.doc` as an engine-plugin in the `plugin-sheets` shape, but lower DOCX onto the engine's
> **native** text + flow stack rather than forking a layout engine — because Word ≈ the DTP text model
> the engine already implements — which delivers native-quality editing (real caret, no C-9) and native
> pagination (no W-frag) at the price of a native↔OOXML save-back layer paid by a preservation invariant.
> Parse/serialize through a new shared **`paged-ooxml`** foundation over the vendored **`ooxmlsdk`**
> crate (gated on a wasm spike), reused by sheet and slide. Support **both** an embedded face (content
> object in a frame) and a standalone face (open the whole `.docx` as the canvas) on one new door
> (`host.document.open(bytes)`), and ratify that embed-or-open symmetry as a platform pattern for every
> content plugin.
