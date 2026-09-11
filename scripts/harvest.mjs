#!/usr/bin/env node
/**
 * Harvest the two library decks into the committed catalogue.
 *
 *   npm run harvest
 *
 * Reads `template/library-16x9.pptx`, `template/library-4x3.pptx` and
 * `template/names.en.json`, and writes `public/catalogue/`:
 *
 *   catalogue.json                 both sizes: every element without its markup
 *   <size>/elements/<id>.json      one element's markup, relationships and parts
 *   <size>/parts/<package path>    the media, charts, embeddings and tags the elements carry
 *
 * and `public/catalogue.html`, the catalogue page the site serves.
 *
 * CI runs this and fails when the result differs from what is committed, so
 * the decks, the names file and the catalogue cannot disagree. The two decks
 * must carry the same keys: an element that exists in one size only is refused
 * here, because the pane never scales an element from the other size.
 *
 * The engine does the reading (`src/core/catalogue/harvest.ts`, pure); this
 * script only touches the file system. It needs `npm run build:lib` first.
 */
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { Pkg, harvest, HarvestError } from "../dist-lib/core/index.js";
import { catalogueHtml } from "./catalogue-page.mjs";

const DECKS = [
  ["16:9", "library-16x9.pptx", "16x9"],
  ["4:3", "library-4x3.pptx", "4x3"],
];
const OUT = "public/catalogue";

const names = JSON.parse(readFileSync("template/names.en.json", "utf8"));
const sizes = {};
const written = [];

rmSync(OUT, { recursive: true, force: true });
for (const [size, file, dir] of DECKS) {
  let result;
  try {
    const pkg = await Pkg.open(new Uint8Array(readFileSync(join("template", file))));
    result = await harvest(pkg, { size, names });
  } catch (error) {
    if (error instanceof HarvestError) {
      console.error(`harvest: ${error.message}`);
      for (const problem of error.problems) console.error(`  - ${problem}`);
      process.exit(2);
    }
    throw error;
  }
  const { catalogue, parts } = result;
  for (const el of catalogue.elements) {
    const { markup, ...rest } = el;
    write(join(OUT, dir, "elements", `${el.id}.json`), JSON.stringify(markup, null, 2) + "\n");
    sizes[size] ??= { ...catalogue, elements: [] };
    sizes[size].elements.push(rest);
  }
  for (const [path, content] of parts) write(join(OUT, dir, "parts", path), content);
  console.log(
    `harvest: ${size} deck: ${catalogue.elements.length} elements in ${catalogue.categories.length} categories, ` +
      `${catalogue.elements.filter((e) => e.kind === "part").length} of them parts, ${parts.size} carried parts`,
  );
}

// Both decks carry the same keys, or the pane would have to scale an element from the other size (refused by design).
const keys = Object.fromEntries(
  Object.entries(sizes).map(([size, c]) => [size, new Set(c.elements.map((e) => e.key))]),
);
const onlyIn = (a, b) => [...keys[a]].filter((k) => !keys[b].has(k));
const missing = [
  ...onlyIn("16:9", "4:3").map((k) => `"${k}" is only in the 16:9 deck`),
  ...onlyIn("4:3", "16:9").map((k) => `"${k}" is only in the 4:3 deck`),
];
if (missing.length) {
  console.error("harvest: the two decks do not carry the same elements:");
  for (const m of missing) console.error(`  - ${m}`);
  process.exit(2);
}

// The version is a hash of the content, so it changes exactly when the catalogue does and never otherwise.
const body = JSON.stringify(sizes);
const version = createHash("sha256").update(body).digest("hex").slice(0, 12);
write(join(OUT, "catalogue.json"), JSON.stringify({ version, sizes }, null, 2) + "\n");

// The catalogue page on the site (`docs/DESIGN.md` section 3): every element
// with its picture and its name, for browsing outside PowerPoint. Written from
// the same catalogue the pane reads and diffed by CI, so it cannot fall behind
// the decks.
writeFileSync("public/catalogue.html", catalogueHtml(sizes, version));

console.log(`harvest: catalogue ${version}, ${written.length} files under ${OUT}, and public/catalogue.html`);

function write(path, content) {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, content);
  written.push(path);
}
