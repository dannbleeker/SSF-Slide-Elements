/**
 * Every reference this repository's prose makes, and whether it lands.
 *
 * This project explains itself by pointing: a comment says `docs/DESIGN.md`
 * section 6, a doc links `[the backlog](docs/BACKLOG.md)`, the manual links a
 * heading inside itself. There are 130 section citations and 32 anchored links,
 * and until now nothing checked a single one of them.
 *
 * That is the quiet kind of rot. Renumbering one heading in `docs/DESIGN.md`
 * silently turns dozens of sentences into directions to the wrong place, and
 * nothing goes red — the docs still build, the suite still passes, and the next
 * reader follows a pointer to a section about something else. The four stale
 * claims fixed by hand on 2026-09-11 were the same shape one level up: prose
 * that was true when written.
 *
 * Three questions, and each is answerable from the files alone:
 *
 * 1. does every relative markdown link name a file that exists;
 * 2. does every `#anchor` name a heading in the file it points at;
 * 3. does every "`docs/X.md` section N" name a section `docs/X.md` has.
 *
 * `test/references.test.ts` asks them, and asks each COUNT as well — a
 * reference gate whose pattern stops matching reports a clean sweep of nothing,
 * which is the failure mode of every guard in this repository that has ever
 * been wrong.
 */
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { join, sep } from "node:path";

const SKIP = new Set(["node_modules", "dist", "dist-lib", ".git", "coverage"]);

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
    else out.push(path.split(sep).join("/").replace(/^\.\//, ""));
  }
  return out;
}

/** Every file whose prose can carry a reference: the docs, the code and the site. */
export function referenceFiles(root = ".") {
  return walk(root).filter((f) => /\.(md|ts|mjs|html)$/.test(f) && !/(^|\/)zz/.test(f));
}

/**
 * A heading's anchor, the way GitHub and most renderers build one.
 *
 * @param {string} heading
 * @returns {string}
 */
export function anchorOf(heading) {
  return heading
    .toLowerCase()
    .replace(/[^\w\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-");
}

/**
 * The numbered sections of a document: `## 7. The gear` answers `7`.
 *
 * Only NUMBERED headings, because only those are what a citation names. A
 * document with none answers an empty map, and a citation into it is a miss
 * rather than a pass — that is the case where a document has been rewritten
 * without its headings and every pointer into it is now wrong.
 *
 * @param {string} path
 * @returns {Map<number, string>}
 */
export function sectionsOf(path) {
  /** @type {Map<number, string>} */
  const out = new Map();
  if (!existsSync(path)) return out;
  for (const line of readFileSync(path, "utf8").split("\n")) {
    const m = /^#{1,3} (\d+)\.\s+(.*)$/.exec(line);
    if (m) out.set(Number(m[1]), m[2] ?? "");
  }
  return out;
}

/**
 * Resolve a link target written inside `from` against the repository root.
 *
 * @param {string} from
 * @param {string} target
 * @returns {string}
 */
function resolveFrom(from, target) {
  if (target.startsWith("/")) return target.slice(1);
  const base = from.includes("/") ? from.slice(0, from.lastIndexOf("/")) : ".";
  return join(base, target).split(sep).join("/");
}

/**
 * Every relative markdown link, with where it points and whether that is there.
 *
 * `[text](path#anchor)` — external links and `mailto:` are somebody else's
 * problem and are left out rather than fetched.
 */
export function markdownLinks(root = ".") {
  /** @type {{ file: string, raw: string, path: string, anchor?: string, exists: boolean }[]} */
  const out = [];
  for (const file of referenceFiles(root).filter((f) => f.endsWith(".md"))) {
    const text = readFileSync(file, "utf8");
    for (const m of text.matchAll(/\[[^\]]*\]\(([^)]+)\)/g)) {
      const raw = m[1] ?? "";
      if (/^(https?:|mailto:|#!)/.test(raw)) continue;
      const [target, anchor] = raw.split("#");
      const path = target ? resolveFrom(file, target) : file;
      out.push({ file, raw, path, anchor, exists: existsSync(path) });
    }
  }
  return out;
}

/** Which of those links point at a file that is not there. */
export function brokenLinks(root = ".") {
  return markdownLinks(root)
    .filter((link) => !link.exists)
    .map((link) => `${link.file} -> ${link.raw}`);
}

/** Which of those links name an anchor the target file has no heading for. */
export function brokenAnchors(root = ".") {
  /** @type {string[]} */
  const out = [];
  for (const link of markdownLinks(root)) {
    if (!link.anchor || !link.exists) continue;
    const headings = readFileSync(link.path, "utf8")
      .split("\n")
      .filter((line) => /^#{1,6} /.test(line))
      .map((line) => anchorOf(line.replace(/^#+\s*/, "")));
    if (!headings.includes(anchorOf(decodeURIComponent(link.anchor)))) out.push(`${link.file} -> ${link.raw}`);
  }
  return out;
}

/**
 * Every "`docs/X.md` section N" in the repository, code comments included.
 *
 * "sections 4, 6 and 8" is one citation naming three, and each is checked: the
 * form is what the prose actually uses.
 */
export function sectionCitations(root = ".") {
  /** @type {{ file: string, doc: string, section: number, exists: boolean }[]} */
  const out = [];
  /** @type {Map<string, Map<number, string>>} */
  const cache = new Map();
  for (const file of referenceFiles(root)) {
    const text = readFileSync(file, "utf8");
    for (const m of text.matchAll(/`?(docs\/[A-Za-z]+\.md)`?\s+sections?\s+([\d,\sand]+)/gi)) {
      const doc = m[1] ?? "";
      if (!cache.has(doc)) cache.set(doc, sectionsOf(doc));
      const sections = cache.get(doc);
      for (const number of (m[2] ?? "").match(/\d+/g) ?? []) {
        out.push({ file, doc, section: Number(number), exists: sections.has(Number(number)) });
      }
    }
  }
  return out;
}

/** Which of those name a section the document does not have. */
export function brokenCitations(root = ".") {
  return sectionCitations(root)
    .filter((cite) => !cite.exists)
    .map((cite) => `${cite.file}: ${cite.doc} has no section ${cite.section}`);
}

/**
 * The questions the probe actually asks, read off the READER.
 *
 * `scripts/read-answers.mjs` prints one numbered heading per question, and it
 * is the honest source for "what does a round answer": the snippet collects
 * arms, the reader is what turns them into questions somebody reads. The two
 * canonical lists — `docs/DESIGN.md` section 13 and `CLAUDE.md`'s "Open
 * questions for the real host" — are prose, and prose is what drifts.
 *
 * It drifted on 2026-09-12: the jump added question 7 to the probe, to
 * `docs/PROBE.md` and to the reader, and to neither list. A round could have
 * answered a question no document was expecting an answer to, and `CLAUDE.md`
 * went on saying the snippet "asks all of them" while asking one more.
 *
 * @param {string} root
 * @returns {number[]}
 */
export function probeQuestions(root = ".") {
  const reader = readFileSync(join(root, "scripts/read-answers.mjs"), "utf8");
  const numbers = [...reader.matchAll(/console\.log\("\\n(\d+)\.\s/g)].map((m) => Number(m[1]));
  return [...new Set(numbers)].sort((a, b) => a - b);
}

/**
 * The numbers of the ordered list that follows a heading, up to the next one.
 *
 * @param {string} path
 * @param {RegExp} heading
 * @returns {number[]}
 */
export function listedQuestions(path, heading) {
  if (!existsSync(path)) return [];
  const lines = readFileSync(path, "utf8").split("\n");
  const at = lines.findIndex((line) => heading.test(line));
  if (at < 0) return [];
  /** @type {number[]} */
  const out = [];
  for (const line of lines.slice(at + 1)) {
    if (/^#{1,3} /.test(line)) break;
    const m = /^(\d+)\.\s/.exec(line);
    if (m) out.push(Number(m[1]));
  }
  return out;
}

/** Where the two canonical lists live, and how each heading is spelled. */
export const QUESTION_LISTS = [
  ["docs/DESIGN.md", /^## \d+\. Open questions for the host$/],
  ["CLAUDE.md", /^## Open questions for the real host$/],
];

/**
 * Every question the probe asks that a canonical list does not carry, and
 * every question a list carries that the probe does not ask.
 *
 * Both directions, because both are the same defect wearing different clothes:
 * an unlisted question is one nobody is expecting an answer to, and a listed
 * question the probe dropped is a round somebody plans that will never happen.
 *
 * @param {string} root
 * @returns {string[]}
 */
export function questionDrift(root = ".") {
  const asked = probeQuestions(root);
  /** @type {string[]} */
  const out = [];
  for (const [path, heading] of QUESTION_LISTS) {
    const listed = listedQuestions(join(root, /** @type {string} */ (path)), /** @type {RegExp} */ (heading));
    for (const n of asked)
      if (!listed.includes(n)) out.push(`${path} does not list question ${n}, which the probe asks`);
    for (const n of listed)
      if (!asked.includes(n)) out.push(`${path} lists question ${n}, which the probe does not ask`);
  }
  return out;
}

/** The heading whose table says which directory owns what. */
const WHERE_THINGS_LIVE = "## Where things live";

/**
 * The rows of `CLAUDE.md`'s directory table: the paths a row is about, and the
 * files its prose names.
 *
 * @param {string} root
 * @returns {{ paths: string[], files: string[] }[]}
 */
export function directoryRows(root = ".") {
  const text = readFileSync(join(root, "CLAUDE.md"), "utf8");
  const at = text.indexOf(WHERE_THINGS_LIVE);
  if (at < 0) return [];
  /** @type {{ paths: string[], files: string[] }[]} */
  const out = [];
  for (const line of text.slice(at + WHERE_THINGS_LIVE.length).split("\n")) {
    if (line.startsWith("## ")) break;
    if (!line.startsWith("| `")) continue;
    const [, first = "", rest = ""] = /^\|([^|]*)\|(.*)$/.exec(line) ?? [];
    const paths = [...first.matchAll(/`([^`]+)`/g)].map((m) => m[1] ?? "");
    const files = [...rest.matchAll(/`([\w.-]+\.(?:ts|mjs))`/g)].map((m) => m[1] ?? "");
    out.push({ paths, files });
  }
  return out;
}

/** The source directories the table is expected to account for. */
const OWNED = ["src/core", "src/host", "src/office", "src/pane", "scripts", "probe", "public", "template"];

/**
 * Where the directory table and the tree disagree.
 *
 * Three questions, and the third is the one that caught the drift this gate
 * was written for. On 2026-09-12 the `src/host/` row named three files and the
 * directory held eight: `coalesce`, `insert`, `jump`, `links` and `timeout`
 * had arrived over four increments and the table said nothing about any of
 * them, while the `src/core/` row still promised the splice was "next".
 *
 * A row is held EXHAUSTIVE only when it chose to enumerate: the directory is
 * flat and the row already names at least one `.ts` file in it. A row that
 * describes its directory in prose — `src/office/`, `scripts/` — is not made
 * to list its files, because that is a different kind of sentence and forcing
 * it would produce a table nobody reads. `.d.ts` is left out: an ambient
 * declaration is not a module anyone navigates to.
 *
 * @param {string} root
 * @returns {string[]}
 */
export function directoryTableProblems(root = ".") {
  const rows = directoryRows(root);
  /** @type {string[]} */
  const out = [];
  const named = new Set(rows.flatMap((row) => row.paths.map((p) => p.replace(/\/$/, ""))));
  for (const dir of OWNED) if (!named.has(dir)) out.push(`the table has no row for ${dir}/`);

  for (const row of rows) {
    for (const path of row.paths) {
      if (!existsSync(join(root, path))) out.push(`the table names ${path}, which is not there`);
    }
    // A row can be about a FILE — `docs/PROBE.md` has one — and only a
    // directory has files to enumerate.
    const dir = (row.paths[0] ?? "").replace(/\/$/, "");
    if (!dir || !existsSync(join(root, dir)) || !statSync(join(root, dir)).isDirectory()) continue;
    const entries = readdirSync(join(root, dir), { withFileTypes: true });
    // Every module the row names is somewhere under the directory it is about.
    // Only walked when the row names one: `public/` holds the built catalogue,
    // and walking it for a row that enumerates nothing is 551 files a run.
    if (row.files.length > 0) {
      const under = new Set(walk(join(root, dir)).map((f) => f.slice(f.lastIndexOf("/") + 1)));
      for (const file of row.files)
        if (!under.has(file)) out.push(`the ${dir}/ row names ${file}, which is not under it`);
    }
    // And, for a flat directory whose row already enumerates, all of them.
    const flat = !entries.some((e) => e.isDirectory());
    const modules = entries.filter((e) => e.isFile() && e.name.endsWith(".ts") && !e.name.endsWith(".d.ts"));
    const enumerated = modules.some((m) => row.files.includes(m.name));
    if (!flat || !enumerated) continue;
    for (const module of modules) {
      if (!row.files.includes(module.name)) out.push(`the ${dir}/ row does not name ${module.name}`);
    }
  }
  return out;
}

/** The heading whose fenced block lists what a maintainer can run. */
const COMMANDS = "## Commands";

/**
 * Every `npm run` command `CLAUDE.md`'s Commands block names, and every script
 * `package.json` actually has.
 *
 * The block is where a session looks for what it may run, and on 2026-09-12 it
 * was missing five: `previews` — a step the Pages deploy runs — along with
 * `print-stamp`, `bench`, `dead-exports` and `release:check`. A command nobody
 * knows about is a step nobody takes, and `npm run previews` is the one that
 * puts the pictures on the tiles.
 *
 * `format:check` and `test:count` are the two the block need not repeat: the
 * first is `format` in its read-only spelling and the second is already there.
 * Everything else must be listed, and everything listed must exist.
 *
 * @param {string} root
 * @returns {string[]}
 */
export function commandDrift(root = ".") {
  const text = readFileSync(join(root, "CLAUDE.md"), "utf8");
  const at = text.indexOf(COMMANDS);
  if (at < 0) return ["CLAUDE.md has no Commands section"];
  const block = text.slice(at).split("```")[1] ?? "";
  const listed = new Set([...block.matchAll(/npm (?:run )?([\w:-]+)/g)].map((m) => m[1] ?? ""));
  const scripts = Object.keys(
    /** @type {{ scripts: Record<string, string> }} */ (JSON.parse(readFileSync(join(root, "package.json"), "utf8")))
      .scripts,
  );
  /** @type {string[]} */
  const out = [];
  for (const name of listed) {
    if (!scripts.includes(name)) out.push(`the Commands block names npm run ${name}, which package.json has not`);
  }
  for (const name of scripts) {
    if (name === "format:check") continue;
    if (!listed.has(name)) out.push(`package.json has ${name} and the Commands block does not name it`);
  }
  return out;
}
