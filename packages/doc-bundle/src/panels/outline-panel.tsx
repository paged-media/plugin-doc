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
// frozen fields, styled-only anchors — previously computed and DISCARDED),
// and save-back readiness. Display-only glue: all values come from the
// engine's already-computed lowering; derivations live in outline-model.ts
// (pure, unit-tested). No DOCX semantics here (CLAUDE.md hard rule).

import type { DockEdge, PanelProps } from "@paged-media/plugin-api";
import * as React from "react";

import {
  OUTLINE_CAP,
  outlineEntries,
  sortedDiagnostics,
  summarize,
  type DocStore,
} from "./outline-model.js";

export interface PanelDescriptor {
  title: string;
  component: React.ComponentType<PanelProps>;
  defaultDock: DockEdge;
}

export interface SaveBackReadiness {
  /** `supports("document.readStory@1")` at ask time. */
  readStory: boolean;
}

const section: React.CSSProperties = { marginTop: 12 };
const h4: React.CSSProperties = {
  margin: "0 0 4px",
  fontSize: 11,
  textTransform: "uppercase",
  letterSpacing: "0.04em",
  opacity: 0.65,
};
const rowStyle: React.CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: 6,
  padding: "1px 0",
};
const ellipsis: React.CSSProperties = {
  overflow: "hidden",
  textOverflow: "ellipsis",
  whiteSpace: "nowrap",
};

const SEVERITY_COLOR: Record<string, string> = {
  error: "var(--status-error, #e5484d)",
  warning: "var(--status-warning, #f5a524)",
  info: "var(--status-info, currentColor)",
};

export function makeOutlinePanel(
  store: DocStore,
  onPlace: () => void,
  onSelectFrame: (frameId: unknown) => void,
  readiness: () => SaveBackReadiness,
): PanelDescriptor {
  const Component: React.FC<PanelProps> = () => {
    const [, force] = React.useReducer((n: number) => n + 1, 0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    React.useEffect(() => store.subscribe(force), []);
    const doc = store.get();

    if (!doc) {
      return (
        <div style={{ padding: 12, fontSize: 13 }} data-doc-panel="empty">
          <p style={{ margin: "0 0 8px", opacity: 0.7 }}>
            No Word document loaded. Use “Place Word document…” or open a{" "}
            <code>.docx</code>.
          </p>
          <button type="button" onClick={() => onPlace()}>
            Place Word document…
          </button>
        </div>
      );
    }

    const { ir } = doc;
    const summary = summarize(ir);
    const outline = outlineEntries(ir);
    const diagnostics = sortedDiagnostics(ir);
    const ready = readiness();
    const saveBackLive = ready.readStory && doc.storyId != null;

    return (
      <div
        style={{ padding: 12, fontSize: 13, lineHeight: 1.5 }}
        data-doc-panel="ready"
      >
        {/* Document */}
        <div style={rowStyle}>
          <strong style={{ flex: 1, minWidth: 0, ...ellipsis }}>
            {doc.fileName}
          </strong>
          {doc.frameId != null && (
            <button
              type="button"
              data-doc-select-frame
              title="Select the placed frame"
              onClick={() => onSelectFrame(doc.frameId)}
            >
              Select
            </button>
          )}
        </div>
        <div style={{ opacity: 0.7, fontSize: 12 }} data-doc-summary>
          {summary.paragraphs} paragraphs · {summary.tables} tables ·{" "}
          {summary.images} images · {summary.hyperlinks} links ·{" "}
          {summary.pageLine}
        </div>
        <button type="button" style={{ marginTop: 6 }} onClick={() => onPlace()}>
          Place Word document…
        </button>

        {/* Outline */}
        <div style={section} data-doc-outline>
          <h4 style={h4}>Outline</h4>
          {outline.length === 0 ? (
            <div style={{ opacity: 0.6, fontSize: 12 }}>
              No headings — {summary.paragraphs} body paragraph
              {summary.paragraphs === 1 ? "" : "s"}.
            </div>
          ) : (
            outline.map((entry, i) => (
              <div
                key={i}
                data-doc-outline-entry={entry.kind}
                style={{
                  ...rowStyle,
                  paddingLeft:
                    entry.kind === "heading" ? (entry.level - 1) * 12 : 0,
                  fontSize: 12,
                }}
              >
                {entry.kind === "heading" ? (
                  <span style={ellipsis}>{entry.text}</span>
                ) : (
                  <span style={{ opacity: 0.75 }}>
                    ▦ Table {entry.rows}×{entry.cols}
                  </span>
                )}
              </div>
            ))
          )}
          {ir.story.blocks.length > OUTLINE_CAP && (
            <div style={{ opacity: 0.6, fontSize: 11 }}>
              (outline capped at {OUTLINE_CAP} entries)
            </div>
          )}
        </div>

        {/* Styles + swatches */}
        <div style={section} data-doc-styles>
          <h4 style={h4}>
            Styles ({ir.styles.length}) · Swatches ({ir.swatches.length})
          </h4>
          {ir.styles.map((s) => (
            <div
              key={s.id}
              data-doc-style-row
              style={{ ...rowStyle, fontSize: 12 }}
            >
              <span style={ellipsis}>{s.name}</span>
              <span style={{ opacity: 0.55, fontSize: 11 }}>
                {s.collection === "paragraph" ? "¶" : "a"}
                {s.basedOn ? ` ← ${s.basedOn.split("/").pop()}` : ""}
              </span>
            </div>
          ))}
          {ir.swatches.length > 0 && (
            <div style={{ ...rowStyle, flexWrap: "wrap", marginTop: 4 }}>
              {ir.swatches.map((sw) => (
                <span
                  key={sw.id}
                  data-doc-swatch
                  title={sw.name}
                  style={{
                    width: 14,
                    height: 14,
                    borderRadius: 3,
                    border: "1px solid rgba(128,128,128,0.4)",
                    background: `rgb(${sw.value.join(",")})`,
                  }}
                />
              ))}
            </div>
          )}
        </div>

        {/* Diagnostics — the ADR-007 payload, previously discarded. */}
        <div style={section} data-doc-diagnostics>
          <h4 style={h4}>Diagnostics ({diagnostics.length})</h4>
          {diagnostics.length === 0 ? (
            <div style={{ opacity: 0.6, fontSize: 12 }}>None.</div>
          ) : (
            diagnostics.map((d, i) => (
              <div
                key={i}
                data-doc-diagnostic={d.severity}
                style={{ ...rowStyle, fontSize: 12, alignItems: "flex-start" }}
              >
                <span
                  style={{
                    color: SEVERITY_COLOR[d.severity],
                    fontSize: 11,
                    marginTop: 1,
                  }}
                >
                  ●
                </span>
                <span>{d.message}</span>
              </div>
            ))
          )}
        </div>

        {/* Save-back readiness */}
        <div
          style={section}
          data-doc-readiness={saveBackLive ? "live" : "verbatim"}
        >
          <h4 style={h4}>Save-back</h4>
          <div style={{ fontSize: 12, opacity: 0.8 }}>
            {saveBackLive
              ? "Edited save-back is LIVE — export re-emits your edits as a targeted .docx patch."
              : ready.readStory
                ? "Story not captured — export re-emits the original .docx verbatim."
                : "Host has no structured story read (document.readStory@1) — export re-emits the original .docx verbatim."}
          </div>
        </div>
      </div>
    );
  };
  return { title: "paged.doc", component: Component, defaultDock: "right" };
}
