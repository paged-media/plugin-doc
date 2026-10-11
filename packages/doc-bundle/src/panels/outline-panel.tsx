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

// The paged.doc outline/styles panel — the real surface over the RETAINED
// LoweredDoc: document summary, heading/table outline, synthesized styles +
// minted swatches, every ADR-007 diagnostic (footnotes, headers/footers,
// frozen fields, styled-only anchors), and save-back readiness.
//
// Outline headings navigate: a click takes the user to the heading in the
// placed story (`host.viewport.revealText`); on a host without that door
// the rows are display-only. Display-only glue: all values come from the
// engine's already-computed lowering; derivations live in outline-model.ts
// (pure, unit-tested). No DOCX semantics here (CLAUDE.md hard rule).

import type { DockEdge, PanelProps } from "@paged-media/plugin-api";
import * as React from "react";

import {
  OUTLINE_CAP,
  headingOccurrence,
  outlineEntries,
  type DocPanelNotice,
  sortedDiagnostics,
  summarize,
  type DocStore,
} from "./outline-model.js";

export interface PanelDescriptor {
  title: string;
  component: React.ComponentType<PanelProps>;
  defaultDock: DockEdge;
  icon: string;
}

export interface SaveBackReadiness {
  /** The host can read an edited story back (structured story read). */
  readStory: boolean;
}

export interface OutlinePanelActions {
  /** Pick a .docx and insert it into the current document. */
  insert: () => void;
  /** False when the host has no file picker (the button is disabled). */
  canInsert: () => boolean;
  selectFrame: (frameId: unknown) => void;
  readiness: () => SaveBackReadiness;
  /** Take the user to a heading (its text and which repeat of that text it
   *  is). Absent → the host cannot navigate, and headings are plain rows.
   *  Resolves false when the heading is not on the page any more. */
  revealHeading?: (text: string, occurrence: number) => Promise<boolean>;
}

/** The menu's verb, used everywhere a document is placed. */
export const INSERT_LABEL = "Insert Word document…";
export const NO_PICKER_REASON =
  "This editor has no file picker. Open the .docx with File ▸ Open, or drop it on the canvas.";
/** Rows shown before a list collapses behind "Show all". */
export const LIST_CAP = 12;
/** Diagnostics are long sentences: fewer of them before "Show all". */
export const DIAGNOSTICS_CAP = 6;

const muted = "var(--pg-muted-fg)";
const section: React.CSSProperties = { marginTop: 14 };
const h4: React.CSSProperties = {
  margin: "0 0 4px",
  fontSize: 12,
  fontWeight: 600,
  color: muted,
};
const rowStyle: React.CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: 6,
  padding: "1px 0",
  minWidth: 0,
};
const ellipsis: React.CSSProperties = {
  overflow: "hidden",
  textOverflow: "ellipsis",
  whiteSpace: "nowrap",
  minWidth: 0,
};
const note: React.CSSProperties = { color: muted, fontSize: 12, margin: 0 };
/** A dot + text row. Grid, not flex: the host wraps overflowing flex rows,
 *  which would put a long message on a line of its own under its dot. */
const dotRow: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "7px minmax(0, 1fr)",
  columnGap: 6,
  alignItems: "start",
  fontSize: 12,
  padding: "1px 0",
};
/** Long engine messages: three lines, the full text on hover. */
const clamp3: React.CSSProperties = {
  display: "-webkit-box",
  WebkitLineClamp: 3,
  WebkitBoxOrient: "vertical",
  overflow: "hidden",
  overflowWrap: "anywhere",
};

const SEVERITY: Record<DocPanelNotice["severity"], { color: string; label: string }> = {
  error: { color: "var(--status-error)", label: "Error" },
  warning: { color: "var(--status-review)", label: "Warning" },
  info: { color: "var(--status-info)", label: "Note" },
};

function plural(n: number, one: string, many = `${one}s`): string {
  return `${n} ${n === 1 ? one : many}`;
}

/** A status dot (CSS, not a text glyph) with its severity as the label. */
function SeverityDot({ severity }: { severity: DocPanelNotice["severity"] }) {
  const s = SEVERITY[severity];
  return (
    <span
      role="img"
      aria-label={s.label}
      title={s.label}
      style={{
        width: 7,
        height: 7,
        marginTop: 6,
        borderRadius: "50%",
        background: s.color,
      }}
    />
  );
}

function Notice({ notice }: { notice: DocPanelNotice }) {
  return (
    <div
      role="status"
      data-doc-notice={notice.severity}
      style={{ ...dotRow, margin: "0 0 8px" }}
    >
      <SeverityDot severity={notice.severity} />
      <span>{notice.message}</span>
    </div>
  );
}

function InsertButton({ actions, style }: { actions: OutlinePanelActions; style?: React.CSSProperties }) {
  const can = actions.canInsert();
  return (
    <button
      type="button"
      data-doc-insert
      style={style}
      disabled={!can}
      title={can ? "Pick a .docx and insert it into this document" : NO_PICKER_REASON}
      onClick={() => actions.insert()}
    >
      {INSERT_LABEL}
    </button>
  );
}

/** A list that shows its first `LIST_CAP` rows and a toggle for the rest. */
function CappedList<T>({
  items,
  noun,
  render,
  cap = LIST_CAP,
}: {
  items: readonly T[];
  noun: string;
  render: (item: T, index: number) => React.ReactNode;
  cap?: number;
}) {
  const [all, setAll] = React.useState(false);
  const shown = all ? items : items.slice(0, cap);
  return (
    <>
      {shown.map(render)}
      {items.length > cap && (
        <button
          type="button"
          data-doc-show-all
          style={{ marginTop: 4, fontSize: 12 }}
          onClick={() => setAll((v) => !v)}
        >
          {all ? "Show fewer" : `Show all ${items.length} ${noun}`}
        </button>
      )}
    </>
  );
}

function readinessLine(live: boolean, readStory: boolean): string {
  if (live) return "Your edits are written back into the .docx when you export it.";
  if (readStory) return "Export writes the original .docx unchanged: this document's text was not captured.";
  return "Export writes the original .docx unchanged: this editor cannot read edited text back.";
}

export function makeOutlinePanel(store: DocStore, actions: OutlinePanelActions): PanelDescriptor {
  const Component: React.FC<PanelProps> = () => {
    const [, force] = React.useReducer((n: number) => n + 1, 0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    React.useEffect(() => store.subscribe(force), []);
    const doc = store.get();
    const notice = store.notice();

    if (!doc) {
      return (
        <div style={{ padding: 12, fontSize: 13 }} data-doc-panel="empty">
          {notice && <Notice notice={notice} />}
          <p style={{ ...note, fontSize: 13, marginBottom: 8 }}>
            Insert a Word document to see its outline, styles and diagnostics.
          </p>
          <InsertButton actions={actions} />
        </div>
      );
    }

    const { ir } = doc;
    const summary = summarize(ir);
    const outline = outlineEntries(ir);
    const diagnostics = sortedDiagnostics(ir);
    const ready = actions.readiness();
    const saveBackLive = ready.readStory && doc.storyId != null;
    const counts = [
      plural(summary.paragraphs, "paragraph"),
      summary.tables > 0 ? plural(summary.tables, "table") : null,
      summary.images > 0 ? plural(summary.images, "image") : null,
      summary.hyperlinks > 0 ? plural(summary.hyperlinks, "link") : null,
    ].filter(Boolean);

    return (
      <div style={{ padding: 12, fontSize: 13, lineHeight: 1.5 }} data-doc-panel="ready">
        {notice && <Notice notice={notice} />}
        {/* Document */}
        <div style={rowStyle}>
          <strong style={{ flex: 1, ...ellipsis }} title={doc.fileName}>
            {doc.fileName}
          </strong>
          {doc.frameId != null && (
            <button
              type="button"
              data-doc-select-frame
              title="Select the frame the document is placed in"
              onClick={() => actions.selectFrame(doc.frameId)}
            >
              Select frame
            </button>
          )}
        </div>
        <div style={note} data-doc-summary>
          {counts.join(" · ")}
        </div>
        <div style={note}>{summary.pageLine}</div>
        <InsertButton actions={actions} style={{ marginTop: 8 }} />

        {/* Outline */}
        <div style={section} data-doc-outline>
          <h4 style={h4}>Outline</h4>
          {outline.length === 0 ? (
            <p style={note}>This document has no headings.</p>
          ) : (
            outline.map((entry, i) => (
              <div
                key={i}
                data-doc-outline-entry={entry.kind}
                style={{
                  ...rowStyle,
                  paddingLeft: entry.kind === "heading" ? (entry.level - 1) * 12 : 0,
                  fontSize: 12,
                }}
              >
                {entry.kind === "heading" && actions.revealHeading ? (
                  <button
                    type="button"
                    data-doc-outline-reveal={i}
                    title={`Go to “${entry.text}”`}
                    onClick={() => {
                      const occurrence = headingOccurrence(outline, i);
                      void actions.revealHeading!(entry.text, occurrence).then((ok) => {
                        if (!ok) {
                          store.setNotice({
                            severity: "info",
                            message: `“${entry.text}” isn't on the page — the heading was edited, or it is in overset text.`,
                          });
                        }
                      });
                    }}
                    style={{
                      ...ellipsis,
                      display: "block",
                      width: "100%",
                      textAlign: "left",
                      background: "transparent",
                      border: "none",
                      padding: 0,
                      color: "inherit",
                      font: "inherit",
                      cursor: "pointer",
                    }}
                  >
                    {entry.text}
                  </button>
                ) : entry.kind === "heading" ? (
                  <span style={ellipsis} title={entry.text}>
                    {entry.text}
                  </span>
                ) : (
                  <span style={{ color: muted }}>
                    Table, {plural(entry.rows, "row")} × {plural(entry.cols, "column")}
                  </span>
                )}
              </div>
            ))
          )}
          {outline.length >= OUTLINE_CAP && (
            <p style={{ ...note, marginTop: 4 }}>Showing the first {OUTLINE_CAP} entries.</p>
          )}
        </div>

        {/* Styles + swatches */}
        <div style={section} data-doc-styles>
          <h4 style={h4}>
            Styles ({ir.styles.length}) and swatches ({ir.swatches.length})
          </h4>
          {ir.styles.length === 0 && <p style={note}>No styles.</p>}
          <CappedList
            items={ir.styles}
            noun="styles"
            render={(s) => (
              <div key={s.id} data-doc-style-row style={{ ...rowStyle, fontSize: 12 }}>
                <span style={{ ...ellipsis, flex: 1 }} title={s.name}>
                  {s.name}
                </span>
                <span style={{ ...ellipsis, flex: "none", maxWidth: "50%", color: muted, fontSize: 11 }}>
                  {s.collection === "paragraph" ? "Paragraph" : "Character"}
                  {s.basedOn ? `, based on ${s.basedOn.split("/").pop()}` : ""}
                </span>
              </div>
            )}
          />
          {ir.swatches.length > 0 && (
            <div style={{ ...rowStyle, flexWrap: "wrap", marginTop: 6 }}>
              {ir.swatches.map((sw) => (
                <span
                  key={sw.id}
                  data-doc-swatch
                  title={sw.name}
                  style={{
                    width: 14,
                    height: 14,
                    borderRadius: 3,
                    border: "1px solid var(--pg-border)",
                    // Swatch CONTENT colour: literal by design.
                    background: `rgb(${sw.value.join(",")})`,
                  }}
                />
              ))}
            </div>
          )}
        </div>

        {/* Diagnostics — what the lowering could not carry (ADR-007). */}
        <div style={section} data-doc-diagnostics>
          <h4 style={h4}>Diagnostics ({diagnostics.length})</h4>
          {diagnostics.length === 0 ? (
            <p style={note}>Everything in the document was carried over.</p>
          ) : (
            <CappedList
              items={diagnostics}
              noun="diagnostics"
              cap={DIAGNOSTICS_CAP}
              render={(d, i) => (
                <div key={i} data-doc-diagnostic={d.severity} style={dotRow}>
                  <SeverityDot severity={d.severity} />
                  <span style={clamp3} title={d.message}>
                    {d.message}
                  </span>
                </div>
              )}
            />
          )}
        </div>

        {/* Save-back readiness */}
        <div style={section} data-doc-readiness={saveBackLive ? "live" : "verbatim"}>
          <h4 style={h4}>Saving back to Word</h4>
          <p style={note}>{readinessLine(saveBackLive, ready.readStory)}</p>
        </div>
      </div>
    );
  };
  // U9 — the tab says what the panel IS ("Document outline"); the owning
  // bundle ("paged.doc") is carried by the host-stamped `source` provenance.
  return { title: "Document outline", component: Component, defaultDock: "right", icon: "panel-outline" };
}
