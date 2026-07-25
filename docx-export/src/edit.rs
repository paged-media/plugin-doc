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

//! The edit description the save-back patcher consumes — expressed in `docx-core`
//! semantics, keyed by LOWERED STORY coordinates (the same `(block, run)` space
//! the Lowered IR and the host mutations use). `diff` produces one of these from
//! two `LoweredDoc`s; the vertical slice hand-authors it.

use docx_core::RunProps;
use serde::{Deserialize, Serialize};

/// A set of run-level edits to apply to one document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EditSet {
    pub runs: Vec<RunEdit>,
}

/// One run's edit, addressed by lowered story `block` index and lowered `run`
/// index within that block's paragraph. `None` fields are "unchanged".
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunEdit {
    pub block: usize,
    pub run: usize,
    /// Replace the run's text (`Some`) — its single `<w:t>` is rewritten.
    pub new_text: Option<String>,
    /// Replace the run's effective DIRECT character properties (`Some`) — its
    /// `<w:rPr>` is rewritten (synthesized styles project to direct `w:rPr`).
    pub new_props: Option<RunProps>,
    /// Set (`Some(Some(id))`) or clear (`Some(None)`) the `<w:rStyle>` to a REAL
    /// Word style id; `None` leaves the existing `rStyle` untouched. Emitted only
    /// alongside `new_props` (rStyle is a child of `w:rPr`).
    pub rstyle: Option<Option<String>>,
}

impl RunEdit {
    /// A pure text change at `(block, run)`.
    pub fn text(block: usize, run: usize, new_text: impl Into<String>) -> Self {
        RunEdit {
            block,
            run,
            new_text: Some(new_text.into()),
            ..Default::default()
        }
    }

    /// A pure character-property change at `(block, run)`.
    pub fn props(block: usize, run: usize, new_props: RunProps) -> Self {
        RunEdit {
            block,
            run,
            new_props: Some(new_props),
            ..Default::default()
        }
    }
}
