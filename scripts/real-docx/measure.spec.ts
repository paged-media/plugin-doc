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
// ADR 029 decision 6, measured (docs/acceptance-real-docx.md). NOT an editor
// spec: the documents are the private corpus, and the Word faces are this
// machine's. scripts/real-docx-acceptance.sh copies it into an editor
// checkout's tests/e2e for one run.
//
// Opens each .docx through File▸Open and records, per paragraph of every
// story, which pages its lines land on. Offsets come from the engine
// (a paragraphBounds walk, byte offsets), never from JS string lengths.
//
//   REAL_DOCX=/abs/a.docx,/abs/b.docx REAL_OUT=/abs/dir REAL_FONTS=word|none \
//   REAL_RELOAD=1 IDML_CANVAS_TEST_PORT=5291 npx playwright test tests/e2e/real-docx-measure.spec.ts
//
// REAL_FONTS=word registers Word's own faces (from Word.app) before the
// open. REAL_RELOAD=1 then saves the poured document and loads it again:
// the engine builds its font table once, at load, from the fonts the
// stories reference THEN, and the skeleton's stories are empty, so without
// the reload every poured family falls back to the default face.
import { test, expect } from "@playwright/test";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { basename } from "node:path";

import { openCanvas } from "../fidelity/canvas-driver";

const DOCS = (process.env.REAL_DOCX ?? "").split(",").filter(Boolean);
const OUT = process.env.REAL_OUT ?? "/tmp";
const FONT_MODE = (process.env.REAL_FONTS ?? "none") + (process.env.REAL_RELOAD === "1" ? "-reload" : "");
const DF = "/Applications/Microsoft Word.app/Contents/Resources/DFonts";
const SUPP = "/System/Library/Fonts/Supplemental";

// Word's own font files (local machine only — never committed).
const WORD_FACES: Array<[string, string | null, string]> = [
  ["Times New Roman", null, `${DF}/times.ttf`],
  ["Times New Roman", "Bold", `${DF}/timesbd.ttf`],
  ["Times New Roman", "Italic", `${DF}/timesi.ttf`],
  ["Times New Roman", "Bold Italic", `${DF}/timesbi.ttf`],
  ["Arial", null, `${DF}/arial.ttf`],
  ["Arial", "Bold", `${DF}/arialbd.ttf`],
  ["Arial", "Italic", `${DF}/ariali.ttf`],
  ["Arial", "Bold Italic", `${DF}/arialbi.ttf`],
  ["Calibri", null, `${DF}/Calibri.ttf`],
  ["Calibri", "Bold", `${DF}/Calibrib.ttf`],
  ["Calibri", "Italic", `${DF}/Calibrii.ttf`],
  ["Calibri", "Bold Italic", `${DF}/Calibriz.ttf`],
  ["Verdana", null, `${DF}/Verdana.ttf`],
  ["Verdana", "Bold", `${DF}/Verdana Bold.ttf`],
  ["Tahoma", null, `${DF}/tahoma.ttf`],
  ["Tahoma", "Bold", `${DF}/tahomabd.ttf`],
  ["Century", null, `${DF}/Century.ttf`],
  ["Lucida Sans Unicode", null, `${DF}/Lucida Sans Unicode.ttf`],
  ["Lucida Sans", null, `${DF}/Lucida Sans.ttf`],
  ["Symbol", null, `${DF}/symbol.ttf`],
  ["Wingdings", null, `${DF}/Wingdings.ttf`],
  ["Courier New", null, `${SUPP}/Courier New.ttf`],
  ["Courier New", "Bold", `${SUPP}/Courier New Bold.ttf`],
  ["Arial Unicode MS", null, `${SUPP}/Arial Unicode.ttf`],
];

type Registries = {
  commands: { invoke: (id: string) => Promise<unknown> };
  importers: { resolve: (name: string) => { id: string } | null };
};
type CanvasGlobal = { __canvas: { ready: boolean; registries: Registries } };

for (const DOC of DOCS) {
  test(`measure ${basename(DOC)} [fonts=${FONT_MODE}]`, async ({ page }) => {
    test.setTimeout(30 * 60_000);
    const diag: string[] = [];
    page.on("console", (m) => diag.push(`[${m.type()}] ${m.text()}`));
    await openCanvas(page);
    await page.evaluate(() =>
      (globalThis as unknown as CanvasGlobal).__canvas.registries.commands.invoke("paged.file.new"),
    );
    await page.waitForFunction(
      () => (globalThis as unknown as CanvasGlobal).__canvas.ready === true,
      null,
      { timeout: 30_000 },
    );
    await expect
      .poll(
        () =>
          page.evaluate(
            () =>
              (globalThis as unknown as CanvasGlobal).__canvas.registries.importers.resolve("x.docx")
                ?.id ?? null,
          ),
        { timeout: 30_000 },
      )
      .toBe("media.paged.doc.importer.docx");

    const registerWordFonts = async () => {
      if (!FONT_MODE.startsWith("word")) return;
      const faces = WORD_FACES.filter(([, , f]) => existsSync(f));
      await page.route("**/__word-fonts/*", (route) => {
        const i = Number(route.request().url().split("/__word-fonts/")[1]);
        return route.fulfill({ status: 200, contentType: "font/ttf", body: readFileSync(faces[i][2]) });
      });
      await page.evaluate(async (list) => {
        const c = (globalThis as unknown as {
          __canvas: { client: { registerFont: (f: string, b: Uint8Array, s: string | null) => Promise<void> } };
        }).__canvas.client;
        for (let i = 0; i < list.length; i++) {
          const res = await fetch(`/__word-fonts/${i}`);
          await c.registerFont(list[i][0], new Uint8Array(await res.arrayBuffer()), list[i][1]);
        }
      }, faces.map(([f, s]) => [f, s] as [string, string | null]));
    };
    await registerWordFonts();
    const t0 = Date.now();
    const chooser = page.waitForEvent("filechooser");
    const opened = page.evaluate(() =>
      (globalThis as unknown as CanvasGlobal).__canvas.registries.commands.invoke("paged.file.openIdml"),
    );
    await (await chooser).setFiles(DOC);
    await opened;
    if (FONT_MODE.startsWith("word")) {
      await page.waitForTimeout(3000);
      await page.unroute("**/__word-fonts/*").catch(() => {});
      await registerWordFonts();
    }
    // Wait for growth to settle: page count stable for 3 consecutive polls.
    let last = -1;
    let stable = 0;
    for (let i = 0; i < 300 && stable < 3; i++) {
      await page.waitForTimeout(2000);
      const n = await page.evaluate(async () => {
        const c = (globalThis as unknown as { __canvas: { client: { executeScript: (s: string) => Promise<{ output: string[] }> } } }).__canvas.client;
        return (JSON.parse((await c.executeScript("paged.pages()")).output[0] ?? "[]") as unknown[]).length;
      });
      stable = n === last ? stable + 1 : 0;
      last = n;
    }
    const openMs = Date.now() - t0;
    // RELOAD: core builds its font table ONCE at load from the fonts the
    // stories reference then; the pour introduces every family after the
    // skeleton load, so a reload is the only way to lay the poured text out
    // in the registered faces.
    if (process.env.REAL_RELOAD === "1") {
      await page.evaluate(async () => {
        const c = (globalThis as any).__canvas.client;
        const bytes = await c.exportPaged();
        await c.loadDocument(bytes);
      });
      await page.waitForTimeout(5000);
      let l2 = -1, st2 = 0;
      for (let i = 0; i < 300 && st2 < 3; i++) {
        await page.waitForTimeout(2000);
        const n = await page.evaluate(async () => (JSON.parse((await (globalThis as any).__canvas.client.executeScript("paged.pages()")).output[0] ?? "[]") as unknown[]).length);
        st2 = n === l2 ? st2 + 1 : 0;
        l2 = n;
      }
    }

    const result = await page.evaluate(async () => {
      type Rect = { pageId: string; topPt: number; leftPt: number };
      const c = (globalThis as unknown as {
        __canvas: {
          client: {
            send: (m: unknown) => Promise<{ kind: string; payload?: { content?: { paragraphs?: Array<{ paragraphStyle?: string; runs: Array<{ text: string }> }> } } }>;
            selectionGeometry: (s: unknown) => Promise<Rect[]>;
            paragraphBounds: (id: string, o: number) => Promise<{ start: number; end: number } | null>;
            executeScript: (s: string) => Promise<{ output: string[] }>;
          };
        };
      }).__canvas.client;
      const pages = JSON.parse((await c.executeScript("paged.pages()")).output[0] ?? "[]") as Array<Record<string, unknown>>;
      const pageIds = pages.map((p) => p.selfId as string);
      const stories = JSON.parse((await c.executeScript("paged.stories()")).output[0] ?? "[]") as Array<{ selfId: string }>;
      const out: Array<{ storyId: string; paras: Array<{ text: string; style?: string; start: number; end: number; pages: number[]; top: number | null }> }> = [];
      for (const { selfId } of stories) {
        const r = await c.send({ kind: "requestStoryContent", payload: { storyId: selfId } });
        const content = r.payload?.content?.paragraphs ?? [];
        const paras: Array<{ text: string; style?: string; start: number; end: number; pages: number[]; top: number | null }> = [];
        let off = 0;
        for (let i = 0; i < content.length; i++) {
          const b = await c.paragraphBounds(selfId, off);
          if (!b) break;
          let rects: Rect[] = [];
          try {
            rects = await c.selectionGeometry({ storyId: selfId, start: b.start, end: Math.max(b.end, b.start), affinity: false });
          } catch {
            rects = [];
          }
          const pg = [...new Set(rects.map((x) => pageIds.indexOf(x.pageId)))];
          paras.push({
            text: content[i].runs.map((x) => x.text).join(""),
            style: content[i].paragraphStyle,
            start: b.start,
            end: b.end,
            pages: pg,
            top: rects.length ? rects[0].topPt : null,
            lines: new Set(rects.map((x) => `${x.pageId}@${Math.round(x.topPt)}`)).size,
          } as never);
          if (b.end + 1 <= off) break;
          off = b.end + 1;
        }
        out.push({ storyId: selfId, paras });
      }
      return { pages, stories: out };
    });
    const name = basename(DOC, ".docx");
    const pdf = await page.evaluate(async () => {
      const c = (globalThis as unknown as { __canvas: { client: { exportPdf: (o: unknown) => Promise<{ bytes: Uint8Array; diagnostics: string[] }> } } }).__canvas.client;
      const r = await c.exportPdf({});
      return { bytes: Array.from(r.bytes), diagnostics: r.diagnostics.slice(0, 50) };
    });
    writeFileSync(`${OUT}/${name}.ours.${FONT_MODE}.pdf`, Buffer.from(pdf.bytes));
    diag.push(...pdf.diagnostics.map((d) => `[pdf] ${d}`));
    writeFileSync(
      `${OUT}/${name}.ours.${FONT_MODE}.json`,
      JSON.stringify({ doc: DOC, fontMode: FONT_MODE, openMs, ...result, console: diag.filter((d) => /warn|error|diagnos|doc/i.test(d)).slice(0, 400) }, null, 1),
    );
    console.log(`${name}: ${result.pages.length} pages, ${result.stories.length} stories, open ${openMs} ms`);
  });
}
