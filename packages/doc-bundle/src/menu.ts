/**
 * paged.doc — the menu bar entries.
 *
 * One command, one entry — and it goes in the host's `Object ▸ Insert` menu rather than a top-level `Doc` menu, because a menu bar with a top level holding a single item is worse than no top level at all.
 *
 * Registered through `contribute.menu()` (plugin-api 0.2.33). Before it
 * there was no menu door at all, so every verb here lived behind Cmd+K
 * and nowhere else.
 *
 * This is the SAME path the host curates as a courtesy for this bundle,
 * which is exactly the case `fallbackFor` was built for: the entry keeps
 * its place for the user, and the host stops naming paged.doc by hand.
 *
 * WHY paged.doc CONTRIBUTES NO MENU OF ITS OWN. A DOCX lowers to real
 * text frames and real stories, so once a document is placed the verbs
 * that apply to it are the HOST's — Type, Edit, and the story tools —
 * not a parallel set. That is the same reason the contract's `mode`
 * defaults to "augment": doc is the plugin that disproves replace-only.
 * */

import type { BundleHost, Disposable } from "@paged-media/plugin-api";

const C = "media.paged.doc.command";

/** `[path, command suffix, group]`. */
const ENTRIES: [path: string, suffix: string, group: string][] = [
  ["Object/Insert Word document…", "placeDoc", "insert-plugin"],
];

/**
 * Register every entry; one Disposable drops them all. Degrades on a
 * host older than plugin-api 0.2.33 by contributing nothing and saying
 * so, rather than throwing and taking the bundle down over a menu.
 */
export function contributeMenu(host: BundleHost): Disposable {
  const contribute = host.contribute as BundleHost["contribute"] & {
    menu?: (c: {
      path: string;
      command: string;
      order?: number;
      group?: string;
    }) => Disposable;
  };
  if (typeof contribute.menu !== "function") {
    host.log.info(
      "host predates contribute.menu (plugin-api 0.2.33) — " +
        `${ENTRIES.length} menu entries not contributed; every command ` +
        "remains reachable through the command palette",
    );
    return { dispose() {} };
  }

  const handles: Disposable[] = [];
  const perGroup = new Map<string, number>();
  for (const [path, suffix, group] of ENTRIES) {
    const n = (perGroup.get(group) ?? 0) + 1;
    perGroup.set(group, n);
    handles.push(
      contribute.menu({ path, command: `${C}.${suffix}`, group, order: n * 10 }),
    );
  }
  host.log.info(`contributed ${handles.length} menu entries`);
  return {
    dispose() {
      for (const h of handles) h.dispose();
      handles.length = 0;
    },
  };
}

/** Exported for the bundle's own test. */
export const MENU_ENTRIES = ENTRIES;
export const MENU_COMMAND_PREFIX = C;
