/**
 * Exports the product never calls.
 *
 * An export nothing imports is a comment that compiles. It type-checks, it has
 * a test of its own, the test is green, and the shipped add-in never runs a
 * line of it — so the thing it was written to fix is still broken. This family
 * has spent four separate sessions on that exact shape: a fix written and one
 * call site missed, an alias the splice was said to adopt shapes by while the
 * splice called something else, a helper the harvest was meant to use.
 *
 * `test/dead-exports.test.ts` holds the repo to this. It is a sweep rather than
 * a memory, because remembering is what failed the last four times.
 *
 * Two things make a sweep like this honest:
 *
 * 1. **Prose is stripped first, but interpolations are not.** A name that
 *    appears only in the paragraph explaining it is not a use, and a sweep that
 *    counted it would report a dead export as alive. A name inside `${…}` IS a
 *    use, and the first version of this sweep — reading `withoutTsProse`, which
 *    blanks a template literal whole — called `nameOfRatio` dead while the pane
 *    ran it on every borrowed deck. `withoutTsText` is the one shade of
 *    difference that costs, and it lives with the other three strippers.
 * 2. **A declaration is not a use.** A name occurring ONCE in the file that
 *    declares it is that declaration. Occurring more than once means the file
 *    dispatches or re-uses it internally, which is how every landing strategy
 *    in `src/core/splice/landing.ts` is reached — a sweep without this rule
 *    called all seven of them dead.
 *
 * What it reports is not automatically a defect. Some exports are deliberate:
 * a seam a sibling repository reaches, or a composition kept for the test that
 * compares the two libraries' palettes. Those are recorded in `ALLOWED` below,
 * each with the reason — and the test also fails when one of those names stops
 * existing, so an exception cannot outlive the thing it excused.
 */
import { readdirSync, readFileSync } from "node:fs";
import { join, sep } from "node:path";
import { isMain } from "./is-main.mjs";
import { withoutTsText } from "./without-prose.mjs";

/** Directories that are not the repo's own source. */
const SKIP = new Set(["node_modules", "dist", "dist-lib", ".git", "coverage", "public", "template", "docs"]);

/**
 * Exports that are reached only by tests ON PURPOSE, and why.
 *
 * An entry here is a decision, not a silence: it names the export and states
 * what reaches it instead of the product. `test/dead-exports.test.ts` requires
 * every one of these to still be an export that exists, so removing the thing
 * removes the excuse with it.
 */
export const ALLOWED = {
  "src/core/pptx/clone.ts::notesPathFor":
    "the sibling merge reaches it; this repo's clone writes the notes path itself",
  "src/core/pptx/theme.ts::themeColoursFor":
    "one call over the chain, for the test that holds the two libraries to the same palette; the harvest needs the chain's parts separately and calls them apart",
  "src/pane/steps.ts::STEPS":
    "the enumeration the gates sweep — test/docs.test.ts holds the manual to every step in it; the pane draws one step at a time and never wants the list",

  // `Pkg`'s own members, which this sweep could not see until 2026-09-16 and
  // which were therefore never triaged. Each is reached by its test and by
  // nothing the add-in runs.
  //
  // **Ask the SIBLING before calling one of these spare.** `src/core` is shared
  // engine, and "nothing in this repo calls it" is not the same claim as
  // "nothing calls it" — a distinction two of these rows got wrong for a day.
  "src/core/pptx/pkg.ts::cachedParts":
    "a diagnostic, and the only exact way to state a held-part property — test/pptx-tags.test.ts uses it to hold reading a deck to a count that does not grow with its slides",
  "src/core/pptx/pkg.ts::partNames":
    "a measurement rather than a manipulation: the package-integrity checks enumerate parts through it, and nothing the add-in runs needs the list",
  "src/core/pptx/pkg.ts::release":
    "the memory tool, measured in the sibling merge at 1697 MB against 93 — this repo clones ONE slide per insert, so nothing here reaches the scale it exists for. Keep it: the scale is a feature away",
  "src/core/pptx/pkg.ts::addContentTypeDefault":
    "the Default-over-Override route its own comment describes, for a caller embedding hundreds of pictures. This repo's splice carries a handful of parts and declares each one, so only its test reaches it",
  "src/core/pptx/pkg.ts::maybeText":
    "the forgiving read. Every caller here asks `has` first and then `text`, so the undefined arm is reached only by its own test",
  // These two were recorded on 2026-09-16 as SUPERSEDED and put to the owner as
  // deletion candidates. THAT WAS WRONG, and the correction is the reason the
  // rows now carry a sibling file and line. `src/core/pptx/pkg.ts` is shared
  // engine, and the verdict had been reached by reading THIS repo only. Both
  // are called by SSF-Merge, checked against its whole tree — 61 source files
  // fetched and grepped, not sampled, after GitHub's code search returned empty
  // for a symbol already proven to be in that repo and so proved only that the
  // search was broken. Deleting either would have broken the sibling at the
  // next port.
  "src/core/pptx/pkg.ts::nextMediaNumber":
    "the sibling merge calls it — `src/core/merge/images.ts:204`, `ppt/media/image${this.pkg.nextMediaNumber()}.${extension}`. Unused HERE: `freeName` in splice/carry.ts numbers every family through the general `nextNumber`, and carried media is named by fingerprint (`ssf-<hash>-<len>.emf`) rather than by extending the image sequence at all",
  "src/core/pptx/pkg.ts::removeSlide":
    "the sibling merge calls it — `src/office/merge.ts:381`, `if (!keep.has(path)) await pkg.removeSlide(path);` — and it takes `orphanedParts` with it, which is why that one is private rather than listed here. Unused HERE: the three packages handed to PowerPoint are reduced by `keepOnly`, which UNLISTS slides and leaves their parts (probe question 1's unlisted arm, the half that touches least). Really removing a slide and its orphans is a capability this add-in never reaches for",
};

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {string[]}
 */
function walk(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (SKIP.has(entry.name)) continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) walk(path, out);
    else out.push(path);
  }
  return out;
}

/**
 * How many times a name appears in a body of code, as a whole word.
 *
 * @param {string} code
 * @param {string} name
 * @returns {number}
 */
function mentions(code, name) {
  return (code.match(new RegExp("\\b" + name + "\\b", "g")) ?? []).length;
}

/**
 * The public methods of every `export class` in a file.
 *
 * WITHOUT THIS THE SWEEP COULD NOT SEE THE BIGGEST SURFACE IN THE REPO. It
 * matched `^export function` and `^export const` and nothing else, so `Pkg` —
 * twenty-eight public members, the package layer the whole engine is built on —
 * was never examined at all, and the sweep reported "0 unexcused" while seven
 * of those members were reached by nothing but their own tests. A gate blind to
 * a whole shape of declaration is a gate that cannot go red for it.
 *
 * Classes here are top level, so a line beginning `}` at column 0 closes one.
 * `private` and `protected` members are not a surface anything outside could
 * call and are skipped; a `constructor` is not a member.
 *
 * @param {string} text
 * @returns {{ name: string, cls: string }[]}
 */
function classMethods(text) {
  /** @type {{ name: string, cls: string }[]} */
  const found = [];
  /** @type {string | undefined} */
  let cls;
  for (const line of text.split("\n")) {
    const open = line.match(/^export (?:abstract )?class (\w+)/);
    if (open) {
      cls = open[1];
      continue;
    }
    if (cls === undefined) continue;
    if (/^\}/.test(line)) {
      cls = undefined;
      continue;
    }
    if (/^ {2}(?:private|protected|readonly)\b/.test(line)) continue;
    const m = line.match(/^ {2}(?:static )?(?:async )?(?:get |set )?([a-zA-Z_$][\w$]*)\s*[(<]/);
    if (!m || m[1] === "constructor") continue;
    found.push({ name: /** @type {string} */ (m[1]), cls });
  }
  return found;
}

/**
 * Every export in `src/` that no other file in `src/` or `scripts/` reaches.
 *
 * Answers `{ file, name, kind, reach }`, where `reach` is `"tests"` when a test
 * imports it and `"nothing"` when not even that does. Types are left out: an
 * exported interface names a shape for a reader, and one used only inside its
 * own file is ordinary rather than dead.
 *
 * @param {string} root
 * @returns {{ file: string, name: string, kind: string, reach: "tests" | "nothing" }[]}
 */
export function deadExports(root = ".") {
  const files = walk(root)
    .map((f) => f.split(sep).join("/").replace(/^\.\//, ""))
    .filter((f) => /\.(ts|mjs)$/.test(f) && !/(^|\/)zz/.test(f));
  /** @type {Map<string, string>} */
  const code = new Map(files.map((f) => [f, withoutTsText(readFileSync(f, "utf8"))]));

  /** @type {{ file: string, name: string, kind: string, reach: "tests" | "nothing" }[]} */
  const out = [];
  for (const [file, text] of code) {
    if (!file.startsWith("src/")) continue;
    /** @type {Map<string, string>} */
    const declared = new Map();
    for (const m of text.matchAll(/^export (?:async )?function (\w+)/gm)) declared.set(m[1], "function");
    for (const m of text.matchAll(/^export const (\w+)/gm)) declared.set(m[1], "const");
    for (const { name, cls } of classMethods(text)) declared.set(name, `method on ${cls}`);
    for (const [name, kind] of declared) {
      if (mentions(text, name) > 1) continue; // dispatched or re-used inside its own file
      const elsewhere = [...code].filter(([other, body]) => other !== file && mentions(body, name) > 0).map(([f]) => f);
      const product = elsewhere.filter((f) => !f.startsWith("test/"));
      if (product.length > 0) continue;
      out.push({ file, name, kind, reach: elsewhere.length > 0 ? "tests" : "nothing" });
    }
  }
  return out.sort((a, b) => a.file.localeCompare(b.file) || a.name.localeCompare(b.name));
}

/**
 * The key an entry is recorded under in `ALLOWED`.
 *
 * @param {{ file: string, name: string }} row
 * @returns {string}
 */
export const keyOf = (row) => `${row.file}::${row.name}`;

if (isMain(import.meta.url)) {
  const rows = deadExports();
  for (const row of rows) {
    const excuse = ALLOWED[keyOf(row)];
    console.log(
      `${(excuse ? "allowed" : row.reach).padEnd(8)} ${row.file} :: ${row.name}${excuse ? ` — ${excuse}` : ""}`,
    );
  }
  console.log(`\n${rows.filter((r) => !ALLOWED[keyOf(r)]).length} unexcused`);
}
