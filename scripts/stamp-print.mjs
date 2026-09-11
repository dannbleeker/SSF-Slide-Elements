#!/usr/bin/env node
/**
 * Stamp each committed print with the deck it was taken from.
 *
 * Run this after re-printing, and commit the sidecars with the PDFs.
 * `test/print.test.ts` holds them to `scripts/print-provenance.mjs`, so a print
 * that drifts from its deck — or a deck edited without a re-print — goes red in
 * CI rather than producing previews cut from the wrong file.
 *
 *   node scripts/stamp-print.mjs --powerpoint 16.0.20326.20132
 */

import { readFileSync, writeFileSync } from "node:fs";
import JSZip from "jszip";
import { slideCountOf, stampFor } from "./print-provenance.mjs";

const DECKS = ["library-16x9", "library-4x3"];

const HOW =
  "File > Export > Create PDF/XPS > Options: range all, publish what Slides, " +
  "frame slides off, include hidden slides on, include comments off, include ink off, " +
  "optimise for Standard. Each read back off the dialog before publishing, because it " +
  "resets to its defaults between presentations.";

const arg = (flag) => {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : undefined;
};

const powerPoint = arg("--powerpoint");
if (!powerPoint) {
  console.error("stamp-print: --powerpoint <version> is required; it is the build that took the print");
  process.exit(2);
}
const takenAt = arg("--on") ?? new Date().toISOString().slice(0, 10);

for (const name of DECKS) {
  const deckPath = `template/${name}.pptx`;
  const printPath = `template/${name}.pdf`;
  const deckBytes = readFileSync(deckPath);
  const printBytes = readFileSync(printPath);

  const zip = await JSZip.loadAsync(deckBytes);
  const slides = slideCountOf(await zip.file("ppt/presentation.xml").async("string"));

  const stamp = stampFor({
    deck: `${name}.pptx`,
    print: `${name}.pdf`,
    deckBytes,
    printBytes,
    slides,
    powerPoint,
    takenAt,
    how: HOW,
  });

  const out = `template/${name}.print.json`;
  writeFileSync(out, JSON.stringify(stamp, null, 2) + "\n");
  console.log(
    `stamp-print: ${out} — ${stamp.slides} slides, ${stamp.pages} pages, deck ${stamp.deckSha256.slice(0, 12)}`,
  );
}
