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
import { existsSync, readdirSync, readFileSync } from "node:fs";
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
