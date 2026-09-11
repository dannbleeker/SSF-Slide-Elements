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
