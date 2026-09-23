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
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { Pkg, harvest, HarvestError } from "../dist-lib/core/index.js";
import { catalogueHtml } from "./catalogue-page.mjs";
import { OUT, publish, stage } from "./catalogue-out.mjs";

const DECKS = [
  ["16:9", "library-16x9.pptx", "16x9"],
  ["4:3", "library-4x3.pptx", "4x3"],
];
// Everything is written to a STAGING directory and the committed catalogue is
// replaced only at the end, once both decks have harvested and agreed. See
// `catalogue-out.mjs`: this script has two early exits, and opening with an
// `rmSync` of `public/catalogue` left the committed `catalogue.json` deleted
// with no replacement on either of them.
const WORK = stage();

const names = JSON.parse(readFileSync("template/names.en.json", "utf8"));
const sizes = {};
const written = [];
/** Every element's markup, rolled up for the version hash below. */
const markupHash = createHash("sha256");

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
  sizes[size] ??= { ...catalogue, elements: [] };
  for (const el of catalogue.elements) {
    const { markup, ...rest } = el;
    const body = JSON.stringify(markup, null, 2) + "\n";
    write(join(WORK, dir, "elements", `${el.id}.json`), body);
    // The markup goes into the VERSION even though it is written to its own
    // file: see the hash below for what it is a cache key for.
    markupHash.update(`${dir}/${el.id}\u0000${body}`);
    sizes[size].elements.push(rest);
  }
  for (const [path, content] of parts) write(join(WORK, dir, "parts", path), content);
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

// The same elements is not the same CATALOGUE. The check above compares keys
// only, so the two decks can file one element under different categories — and
// the pane's chips, the summary count and `catalogue.html` all come straight
// from `library.categories`, which is per size.
//
// Measured on the committed catalogue, 2026-09-23: 16:9 carries 11 categories
// and 4:3 carries 10, "Hvide kasser med sorte overskrifter" exists only in
// 16:9, and five elements sit under a different heading in the two decks. A
// user on a 4:3 deck sees a different library from a user on a 16:9 one.
//
// Reported rather than refused, deliberately: this is the owner's DECK content
// and only they can move a slide between collections, so failing here would
// block every harvest until the decks are re-cut. It is loud, it names every
// one, and it is written down in `docs/BACKLOG.md`.
const catsOf = (size) => new Set(sizes[size].categories.map((c) => c.key));
const cats = { "16:9": catsOf("16:9"), "4:3": catsOf("4:3") };
const filedUnder = (size) => Object.fromEntries(sizes[size].elements.map((e) => [e.key, e.category.key]));
const at169 = filedUnder("16:9");
const drift = [
  ...[...cats["16:9"]].filter((k) => !cats["4:3"].has(k)).map((k) => `category "${k}" is only in the 16:9 deck`),
  ...[...cats["4:3"]].filter((k) => !cats["16:9"].has(k)).map((k) => `category "${k}" is only in the 4:3 deck`),
  ...sizes["4:3"].elements
    .filter((e) => at169[e.key] !== undefined && at169[e.key] !== e.category.key)
    .map((e) => `"${e.key}" is under "${at169[e.key]}" at 16:9 and "${e.category.key}" at 4:3`),
];
if (drift.length) {
  console.error(`harvest: WARNING — the two decks disagree about ${drift.length} categorisation(s):`);
  for (const d of drift) console.error(`  - ${d}`);
  console.error("  The pane shows a different library depending on the deck's shape. Fix the decks, not the code.");
}

// The version is a hash of the content, so it changes exactly when the
// catalogue does and never otherwise.
//
// It hashed `sizes` alone, which is the INDEX — every element without its
// markup. That is not what the version is used for. `previewUrl` in
// `src/pane/catalogue.ts` makes it the `?v=` on every preview picture, and the
// header of `scripts/build-previews.mjs` says why: the preview name is derived
// from the element's id rather than hashed, so the version is the only thing
// that can retire a cached one. A deck edit that changes PIXELS but no metadata
// — a recolour, a corrected typo inside a shape, a line weight — left every
// committed field untouched, so the version did not move and a returning user
// kept the old picture. The markup is not covered by it either, and is fetched
// with no cache key at all.
//
// So the markup rolls in above, and the print stamps here: those carry the SHA
// of the deck each PDF was printed from, which is the closest thing the tree
// has to "the pixels changed".
for (const [, file] of DECKS) {
  const stamp = `template/${file.replace(/\.pptx$/, ".print.json")}`;
  if (existsSync(stamp)) markupHash.update(readFileSync(stamp));
}
const body = JSON.stringify(sizes) + "\u0000" + markupHash.digest("hex");
const version = createHash("sha256").update(body).digest("hex").slice(0, 12);
write(join(WORK, "catalogue.json"), JSON.stringify({ version, sizes }, null, 2) + "\n");

// The catalogue page on the site (`docs/DESIGN.md` section 3): every element
// with its picture and its name, for browsing outside PowerPoint. Written from
// the same catalogue the pane reads and diffed by CI, so it cannot fall behind
// the decks.
writeFileSync("public/catalogue.html", catalogueHtml(sizes, version));

// The last thing, and the only destructive one: everything above has run, so
// there is a whole catalogue to put in place of the committed one.
publish(WORK, OUT);

console.log(`harvest: catalogue ${version}, ${written.length} files under ${OUT}, and public/catalogue.html`);

function write(path, content) {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, content);
  written.push(path);
}
