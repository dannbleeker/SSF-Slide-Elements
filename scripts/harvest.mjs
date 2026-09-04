/**
 * Turn the library deck into the catalogue the pane ships with.
 *
 * Run after any change to `template/*.pptx`:
 *
 *     npm run build:lib && node scripts/harvest.mjs
 *
 * (Two flat halves rather than one nested `npm run`, because a nested one is
 * blocked by AppLocker on the owner's Windows box. The sibling projects both
 * learned this; every npm script here stays flat.)
 *
 * Writes `src/pane/catalogue/index.json` and one payload per element. Both are
 * committed: they are the add-in's content, and a build that fetches them from
 * anywhere at runtime would be the one thing this design rules out.
 */
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Pkg } from "../dist-lib/core/pptx/pkg.js";
import { harvest } from "../dist-lib/core/catalogue/harvest.js";
import { previewSvg } from "../dist-lib/core/preview/svg.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const templateDir = join(root, "template");
const outDir = join(root, "public", "catalogue");

/**
 * `btoa`/`atob` are DOM globals. The engine uses them so one implementation
 * runs in the pane and in the suite; Node has had them since 16, but they are
 * not on the global in every runner, so they are ensured rather than assumed.
 */
if (typeof globalThis.btoa !== "function") {
  globalThis.btoa = (s) => Buffer.from(s, "binary").toString("base64");
}
if (typeof globalThis.atob !== "function") {
  globalThis.atob = (s) => Buffer.from(s, "base64").toString("binary");
}

const decks = readdirSync(templateDir).filter((f) => f.toLowerCase().endsWith(".pptx"));
if (decks.length === 0) {
  console.error(`no .pptx in ${templateDir} — the library deck is the input to this script`);
  process.exit(1);
}
if (decks.length > 1) {
  console.error(`${decks.length} decks in ${templateDir}; expected exactly one library deck:\n  ${decks.join("\n  ")}`);
  process.exit(1);
}

const deckName = decks[0];
const pkg = await Pkg.open(new Uint8Array(readFileSync(join(templateDir, deckName))));
const { index, payloads, notes } = await harvest(pkg, {
  libraryName: deckName,
  preview: (shapes, bounds) => previewSvg(shapes, bounds),
});

rmSync(outDir, { recursive: true, force: true });
mkdirSync(join(outDir, "elements"), { recursive: true });
writeFileSync(join(outDir, "index.json"), `${JSON.stringify(index, null, 2)}\n`);
for (const payload of payloads) {
  writeFileSync(join(outDir, "elements", `${payload.id}.json`), JSON.stringify(payload));
}

const bytes = payloads.reduce((n, p) => n + JSON.stringify(p).length, 0);
const indexBytes = JSON.stringify(index).length;
console.log(`library      ${deckName}`);
console.log(`sections     ${index.sections.length}`);
console.log(`elements     ${index.elements.length}`);
console.log(`index        ${(indexBytes / 1024).toFixed(0)} KB  (loaded once, at startup)`);
console.log(
  `payloads     ${(bytes / 1024 / 1024).toFixed(1)} MB across ${payloads.length} files (fetched one at a time)`,
);

const skipped = notes.filter((n) => n.level === "skipped");
const warnings = notes.filter((n) => n.level === "warning");
if (skipped.length > 0) {
  console.log(`\nskipped ${skipped.length} slide(s):`);
  for (const n of skipped) console.log(`  slide ${n.slide}: ${n.detail}`);
}
if (warnings.length > 0) {
  console.log(`\n${warnings.length} warning(s):`);
  for (const n of warnings.slice(0, 40)) console.log(`  slide ${n.slide}: ${n.detail}`);
  if (warnings.length > 40) console.log(`  … and ${warnings.length - 40} more`);
}

// A section with nothing in it means a divider whose elements all failed to
// harvest — the shape of failure that otherwise reads as a successful run with
// a slightly shorter list.
const empty = index.sections.filter((s) => !index.elements.some((e) => e.section === s.id));
if (empty.length > 0) {
  console.error(`\n${empty.length} section(s) harvested no elements: ${empty.map((s) => s.name).join(", ")}`);
  process.exit(1);
}
