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

// The paged.doc panel — a minimal status surface (loaded document + a place
// button). The styles panel / document outline is a later-tier feature; this
// keeps the contribution honest without over-promising.

import type { DockEdge, PanelProps } from "@paged-media/plugin-api";
import * as React from "react";

export interface PanelDescriptor {
  title: string;
  component: React.ComponentType<PanelProps>;
  defaultDock: DockEdge;
}

export function makeOutlinePanel(
  currentFileName: () => string | null,
  onPlace: () => void,
): PanelDescriptor {
  const Component: React.FC<PanelProps> = () => {
    const fileName = currentFileName();
    return (
      <div style={{ padding: "var(--space-3, 12px)", fontSize: 13 }}>
        <h3 style={{ margin: "0 0 8px" }}>paged.doc</h3>
        {fileName ? (
          <p>
            Loaded: <strong>{fileName}</strong>
          </p>
        ) : (
          <p style={{ opacity: 0.7 }}>
            No Word document loaded. Use “Place Word document…” or open a{" "}
            <code>.docx</code>.
          </p>
        )}
        <button type="button" onClick={() => onPlace()}>
          Place Word document…
        </button>
      </div>
    );
  };
  return { title: "paged.doc", component: Component, defaultDock: "right" };
}
