/**
 * Splice every element in the catalogue into a real deck and check the result.
 *
 * The one gate that exercises the whole product without a PowerPoint: harvest
 * output in, a package PowerPoint would open out, ninety-eight times. Every
 * failure `splice.ts` can produce is structural, and structural failures are
 * exactly what a script can see.
 *
 *     npm run build:lib && node scripts/verify-splice.mjs
 */
import { readFileSync, readdirSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Pkg } from "../dist-lib/core/pptx/pkg.js";
import { splice } from "../dist-lib/core/element/splice.js";
import { verify } from "../dist-lib/core/element/verify.js";

if (typeof globalThis.btoa !== "function") globalThis.btoa = (s) => Buffer.from(s, "binary").toString("base64");
if (typeof globalThis.atob !== "function") globalThis.atob = (s) => Buffer.from(s, "base64").toString("binary");

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const catalogue = join(root, "src", "pane", "catalogue");
const index = JSON.parse(readFileSync(join(catalogue, "index.json"), "utf8"));
const deckName = readdirSync(join(root, "template")).find((f) => f.toLowerCase().endsWith(".pptx"));
const deckBytes = new Uint8Array(readFileSync(join(root, "template", deckName)));

let failed = 0;
let checked = 0;
const t0 = Date.now();

for (const entry of index.elements) {
  const payload = JSON.parse(readFileSync(join(catalogue, "elements", `${entry.id}.json`), "utf8"));
  // A fresh package per element: a splice mutates, and reusing one would test
  // ninety-eight elements piled onto one slide rather than each on its own.
  const pkg = await Pkg.open(deckBytes);
  const slides = await pkg.slidePaths();
  // Slide 1 is the cover — a slide with a picture and a title, which is a
  // realistic thing to drop an element onto and is not one of the library's
  // own content slides.
  const target = slides[0];
  const before = (await pkg.rels(target)).length;
  const report = await splice(pkg, target, payload, { x: 914400, y: 914400 });
  const findings = await verify(pkg, target);
  checked += 1;
  if (findings.length > 0) {
    failed += 1;
    console.error(`FAIL ${entry.id}`);
    for (const f of findings.slice(0, 5)) console.error(`     ${f.kind}: ${f.detail}`);
  } else if (report.shapes !== payload.shapes.length) {
    failed += 1;
    console.error(`FAIL ${entry.id}: spliced ${report.shapes} of ${payload.shapes.length} shapes`);
  }
  void before;
}

const secs = ((Date.now() - t0) / 1000).toFixed(1);
console.log(`\n${checked - failed}/${checked} elements spliced cleanly in ${secs}s`);
if (failed > 0) process.exit(1);
