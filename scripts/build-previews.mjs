#!/usr/bin/env node
/**
 * Cut every element's preview out of its deck's print.
 *
 * `docs/DESIGN.md` section 3: previews are PowerPoint's own rendering, because
 * nothing else draws an SSF element the way PowerPoint does. There is no
 * PowerPoint in CI, so the owner prints each deck to PDF and this reads the
 * print.
 *
 *   npx playwright install chromium      # once
 *   node scripts/build-previews.mjs
 *
 * WHERE THE PIXELS COME FROM. pdf.js rasterises a page, inside the Chromium
 * that `pane-shots` already uses — so a real canvas does the drawing and this
 * repo gains no native dependency and nothing at runtime. The page is rendered
 * ONCE and every element on it is cut from that one raster; the 4:3 deck's
 * flowchart slide alone yields ten.
 *
 * WHAT IS CUT is not decided here. `src/core/catalogue/cut.ts` owns the
 * geometry — the air, the neighbours painted out, the rotated mask — and is
 * unit-tested, because a cut that is slightly wrong still produces a picture
 * that looks plausible. This file turns those rectangles into files.
 *
 * NAMES ARE DERIVED, NOT HASHED. Section 11 asks for hashed names, and a hash
 * of the rendered bytes cannot go in the committed index: font rasterisation
 * differs between machines, so the index would stop matching the decks on
 * whichever machine did not render it, and CI's "catalogue is committed" check
 * would fail for a reason that has nothing to do with the catalogue. The file
 * is `previews/<element id>.png` and the catalogue's own content `version` is
 * the cache key. The index therefore does not change when a preview is rebuilt.
 */
import { mkdirSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { chromium } from "playwright";
import { cutsFor } from "../dist-lib/core/catalogue/cut.js";

const OUT = "public/catalogue";
const MAX_EDGE = Number(process.env.PREVIEW_MAX_EDGE ?? 480);
const PAGE_SCALE = Number(process.env.PREVIEW_PAGE_SCALE ?? 2);

const DECKS = [
  { size: "16:9", print: "template/library-16x9.pdf", dir: "16x9" },
  { size: "4:3", print: "template/library-4x3.pdf", dir: "4x3" },
];

const PDFJS = readFileSync("node_modules/pdfjs-dist/build/pdf.min.mjs", "utf8");
const PDFJS_WORKER = readFileSync("node_modules/pdfjs-dist/build/pdf.worker.min.mjs", "utf8");

const catalogue = JSON.parse(readFileSync(`${OUT}/catalogue.json`, "utf8"));

const browser = await chromium.launch();
const page = await browser.newPage();
page.on("pageerror", (e) => console.error("preview page error:", e.message));
await page.setContent("<!doctype html><title>previews</title><body></body>");
await page.addScriptTag({ content: PDFJS, type: "module" });
await page.evaluate((workerSource) => {
  // pdf.js wants its worker as a URL; give it one made out of the source we
  // already have on disk, so nothing is fetched.
  const blob = new Blob([workerSource], { type: "text/javascript" });
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
}, PDFJS_WORKER);

let written = 0;
const report = [];

for (const deck of DECKS) {
  const elements = catalogue.sizes[deck.size].elements;
  const cuts = cutsFor(elements);
  const dir = `${OUT}/${deck.dir}/previews`;
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(dir, { recursive: true });

  const pdfBase64 = readFileSync(deck.print).toString("base64");
  await page.evaluate(async (b64) => {
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    window.__doc = await window.pdfjsLib.getDocument({ data: bytes }).promise;
  }, pdfBase64);

  const pages = Number(await page.evaluate(() => window.__doc.numPages));
  if (pages !== new Set(elements.map((e) => e.slide)).size && pages < Math.max(...elements.map((e) => e.slide))) {
    throw new Error(
      `${deck.print}: ${pages} pages, but an element names slide ${Math.max(...elements.map((e) => e.slide))}`,
    );
  }

  // group the cuts by page, so each page is rasterised once
  const byPage = new Map();
  for (const cut of cuts) {
    if (!byPage.has(cut.page)) byPage.set(cut.page, []);
    byPage.get(cut.page).push(cut);
  }

  for (const [pageNo, pageCuts] of [...byPage.entries()].sort((a, b) => a[0] - b[0])) {
    const images = await page.evaluate(
      async ({ pageNo, pageCuts, maxEdge, scale }) => {
        const pdfPage = await window.__doc.getPage(pageNo);
        const viewport = pdfPage.getViewport({ scale });
        const sheet = document.createElement("canvas");
        sheet.width = Math.ceil(viewport.width);
        sheet.height = Math.ceil(viewport.height);
        const sheetCtx = sheet.getContext("2d");
        sheetCtx.fillStyle = "#ffffff";
        sheetCtx.fillRect(0, 0, sheet.width, sheet.height);
        await pdfPage.render({ canvasContext: sheetCtx, viewport }).promise;

        const out = [];
        for (const cut of pageCuts) {
          const sx = cut.crop.x * sheet.width;
          const sy = cut.crop.y * sheet.height;
          const sw = cut.crop.w * sheet.width;
          const sh = cut.crop.h * sheet.height;
          const ratio = Math.min(1, maxEdge / Math.max(sw, sh));
          const dw = Math.max(1, Math.round(sw * ratio));
          const dh = Math.max(1, Math.round(sh * ratio));

          const tile = document.createElement("canvas");
          tile.width = dw;
          tile.height = dh;
          const ctx = tile.getContext("2d");
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, dw, dh);

          // page fractions -> this tile's pixels
          const toTileX = (fx) => ((fx - cut.crop.x) / cut.crop.w) * dw;
          const toTileY = (fy) => ((fy - cut.crop.y) / cut.crop.h) * dh;

          if (cut.mask) {
            ctx.save();
            ctx.beginPath();
            cut.mask.forEach((p, i) => {
              const x = toTileX(p.x);
              const y = toTileY(p.y);
              if (i === 0) ctx.moveTo(x, y);
              else ctx.lineTo(x, y);
            });
            ctx.closePath();
            ctx.clip();
            ctx.drawImage(sheet, sx, sy, sw, sh, 0, 0, dw, dh);
            ctx.restore();
          } else {
            ctx.drawImage(sheet, sx, sy, sw, sh, 0, 0, dw, dh);
          }

          for (const box of cut.whiteOut) {
            ctx.fillStyle = "#ffffff";
            ctx.fillRect(toTileX(box.x), toTileY(box.y), (box.w / cut.crop.w) * dw, (box.h / cut.crop.h) * dh);
          }

          // is anything actually there? an all-white tile is a cut that missed.
          const data = ctx.getImageData(0, 0, dw, dh).data;
          let ink = 0;
          for (let i = 0; i < data.length; i += 4) {
            if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) ink++;
          }

          out.push({
            id: cut.id,
            png: tile.toDataURL("image/png").split(",")[1],
            w: dw,
            h: dh,
            inkFraction: ink / (dw * dh),
          });
        }
        return out;
      },
      { pageNo, pageCuts, maxEdge: MAX_EDGE, scale: PAGE_SCALE },
    );

    for (const image of images) {
      const bytes = Buffer.from(image.png, "base64");
      writeFileSync(`${dir}/${image.id}.png`, bytes);
      written++;
      report.push({
        size: deck.size,
        id: image.id,
        page: pageNo,
        w: image.w,
        h: image.h,
        ink: image.inkFraction,
        bytes: bytes.length,
      });
    }
  }

  console.log(`previews: ${deck.size} — ${byPage.size} pages read, ${elements.length} cut into ${dir}`);
}

await browser.close();

const blank = report.filter((r) => r.ink < 0.001);
const total = report.reduce((sum, r) => sum + r.bytes, 0);
console.log(`previews: ${written} files, ${(total / 1024 / 1024).toFixed(1)} MB`);
if (blank.length) {
  console.error(`previews: ${blank.length} came out blank, which is a cut that missed:`);
  for (const b of blank.slice(0, 20)) console.error(`  ${b.size} ${b.id} (page ${b.page})`);
  process.exit(1);
}
writeFileSync(`${OUT}/previews-report.json`, JSON.stringify(report, null, 2) + "\n");
