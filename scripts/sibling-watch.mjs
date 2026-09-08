#!/usr/bin/env node
/**
 * Watch both sibling projects for findings nobody here has answered.
 *
 * Most of what this repo knows about the PowerPoint host was learned by
 * SSF-Charts over its real-host rounds and carried into SSF-Merge, and until
 * now it would arrive here by somebody happening to read those repos. That is
 * not a process; it is luck with a good afternoon attached. SSF-Merge measured
 * what the luck cost it: 44 hand-copied citations, five carrying a counter that
 * went stale the next time a round ran.
 *
 * So: look every week, automatically, and say only what is NEW.
 *
 * **`TRIAGED` is the load-bearing half.** Without it the sweep reports the same
 * seventy findings every Monday and is filtered within a month. With it the
 * report is precisely "here is something nobody here has looked at", and the
 * table doubles as the record of what was decided — which is the question a
 * reader actually has when they meet a finding referenced in a comment.
 *
 * Usage:
 *   node scripts/sibling-watch.mjs                  # fetch and report
 *   node scripts/sibling-watch.mjs --from dir/      # report from saved copies
 *   node scripts/sibling-watch.mjs --json           # machine-readable output
 *
 * Exit 0 when everything is answered, 3 when something is not, anything else
 * when the sweep itself broke. A broken run must never read as a quiet week.
 *
 * **Raw file reads only, never the GitHub API.** All three repositories are
 * public, so `raw.githubusercontent.com` answers without a token — which means
 * this runs in CI and in an agent session alike.
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { isMain } from "./is-main.mjs";

/**
 * Where each sibling's files are read from. The canonical repository names:
 * SSF-Charts was once called PowerChart, and the old name still redirects, but
 * a redirect is the thing that one day stops.
 */
export const RAW = {
  charts: "https://raw.githubusercontent.com/dannbleeker/SSF-Charts/main",
  merge: "https://raw.githubusercontent.com/dannbleeker/SSF-Merge/main",
};

/**
 * The siblings' curated tables, and nothing else.
 *
 * SSF-Charts' `CLAUDE.md` is the fullest record and the worst feed: thousands
 * of lines of prose that diffs into noise. Its keyed tables are small and
 * already worded for a machine to compare — they are what that project uses to
 * gate its own fake against a real host. SSF-Merge publishes no keyed table of
 * its own host answers (they are prose in its CLAUDE.md and sheets under its
 * docs/host-answers/), so what is watched there is its `TRIAGED` table: a
 * finding Merge has triaged that has no row here is worth a look, and a key
 * Merge dropped is worth knowing about.
 *
 * `repo` says where the file is fetched from; `space` says whose findings the
 * keys name (Merge's TRIAGED keys are Charts' findings, so both are `charts`
 * and one row answers a finding wherever it appears); `kind` is how the key is
 * spelled, or `self` when the key already carries its kind.
 *
 * `mayBeEmpty`: SSF-Charts' `PENDING_QUESTIONS` is empty by design between a
 * probe's commit and its answering round, and SSF-Merge's watcher has been red
 * on it since 2026-09-02 because an empty table parses as "unreadable". Here a
 * table that is PRESENT but parses to nothing reads as empty; an absent table
 * is still a broken sweep, for every source.
 */
export const SOURCES = [
  { repo: "charts", space: "charts", path: "scripts/office-js-watch.mjs", table: "KNOWN_ISSUES", kind: "issue" },
  { repo: "charts", space: "charts", path: "scripts/host-baseline.mjs", table: "FAKE_BASELINE", kind: "question" },
  { repo: "charts", space: "charts", path: "scripts/host-baseline.mjs", table: "KNOWN_DIVERGENCES", kind: "question" },
  { repo: "charts", space: "charts", path: "scripts/host-baseline.mjs", table: "UNSTABLE_ANSWERS", kind: "question" },
  {
    repo: "charts",
    space: "charts",
    path: "scripts/host-baseline.mjs",
    table: "PENDING_QUESTIONS",
    kind: "question",
    mayBeEmpty: true,
  },
  { repo: "merge", space: "charts", path: "scripts/sibling-watch.mjs", table: "TRIAGED", kind: "self" },
];

/**
 * The keys of one `export const NAME = { … }` table.
 *
 * A regex over two-space-indented keys rather than an import, because importing
 * would run the sibling's code — a weekly job that executes a file fetched over
 * the network is a supply chain, not a sweep. Both quoted and bare keys, since
 * the siblings use both; a QUOTED key may hold a colon, which is how SSF-Merge
 * spells its `TRIAGED` keys (`"issue:1650":`) and which the sibling's own
 * reader could not see — it read `issue` and stopped.
 *
 * Returns null when the table is not there at all, or when it parses to
 * nothing: a rename or a reindent upstream would otherwise report "nothing new"
 * forever, which is the failure this whole file exists to prevent. Whether an
 * empty table is a broken one is the caller's decision (`mayBeEmpty`).
 *
 * @param {string} source
 * @param {string} table
 * @returns {string[] | null}
 */
export function tableKeys(source, table) {
  const at = source.indexOf(`export const ${table}`);
  if (at < 0) return null;
  const rest = source.slice(at);
  const end = rest.search(/^\};/m);
  const body = end < 0 ? rest : rest.slice(0, end);
  const keys = [...body.matchAll(/^ {2}(?:"([A-Za-z0-9][A-Za-z0-9._:-]*)"|([A-Za-z0-9][A-Za-z0-9._-]*)):/gm)].map((m) =>
    String(m[1] ?? m[2]),
  );
  return keys.length === 0 ? null : keys;
}

/**
 * Whether a table is DECLARED in a source, whatever it holds.
 *
 * @param {string} source
 * @param {string} table
 * @returns {boolean}
 */
export function tableDeclared(source, table) {
  return source.includes(`export const ${table}`);
}

/**
 * The id a finding is triaged under.
 *
 * Namespaced by whose finding it is, never by where it was read: a Charts
 * finding read from Merge's table is the same finding, and one row answers it.
 *
 * @param {{space: string, kind: string}} source
 * @param {string} key
 */
export function idOf(source, key) {
  return source.kind === "self" ? `${source.space}:${key}` : `${source.space}:${source.kind}:${key}`;
}

/**
 * What this repo has answered about the siblings' findings.
 *
 * Every key carries its reason, and **"no exposure" is a real answer** — the
 * most common one worth writing down, because an untriaged finding is
 * indistinguishable from an unnoticed one.
 *
 * Seeded on 2026-09-08 from the first sweep's report. Where SSF-Merge's verdict
 * holds here by construction — package route, one insert, no shape-level API —
 * its reason is kept; where Merge ADOPTED something in code this repo does not
 * have yet, the row is RELEVANT and says to re-triage when the insert lands.
 * Keys are `charts:issue:N` and `charts:question:id`, the sibling's own
 * spelling under whose finding it is, so a rename there surfaces as a new
 * finding rather than silently matching nothing.
 */
export const TRIAGED = {
  "charts:issue:1650":
    "ADOPTED as doctrine (CLAUDE.md host rules), no code yet — a slide add whose sync never resolves though the slide lands. The insert will be one `insertSlidesFromBase64`, and the deck DELTA is the evidence, never the absence of an error; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:2328":
    "NO EXPOSURE — slideMaster.shapes throws GeneralException on the web. Nothing here reads a master; a spliced slide carries its layout relationship inside the package, and the API is never asked.",
  "charts:issue:2699":
    "NO EXPOSURE — a blank slide that is only blank to the eye. NO EXPOSURE — nothing here inspects whether a slide is blank.",
  "charts:issue:2775":
    "RELEVANT — addTextBox deletes the SELECTED shape, web only. No exposure to the call: nothing here adds a text box. The CLASS matters: an insert lands while the user may have something selected, and SSF-Merge has a probe question (`insertWhileSelectedProbe`) asking whether a SLIDE insert survives a standing selection, unanswered as of 2026-09-08; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:2172":
    "NO EXPOSURE — addGeometricShape refused on a completely blank slide. NO EXPOSURE — no shape is ever added through the API.",
  "charts:issue:2780":
    "NO EXPOSURE — a documented caveat with no code depending on it, in the sibling and here alike. NO EXPOSURE.",
  "charts:issue:2881":
    "NO EXPOSURE — complex SVG renders wrong through the picture path. NO EXPOSURE — nothing here inserts a picture.",
  "charts:issue:2903":
    "ADOPTED as doctrine (CLAUDE.md host rules), no code yet — a stale shape proxy answers `InvalidParam passed to GetItem(id)`. It is why anything this add-in needs to remember goes into the PACKAGE before the insert and never through a proxy afterwards; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:3014":
    "NO EXPOSURE — powerPoint's API has no grouping story. NO EXPOSURE — no shape-level work at all.",
  "charts:issue:3083":
    "NO EXPOSURE — setSelectedShapes([]) does not clear the selection on the web. NO EXPOSURE — this add-in never calls `setSelectedShapes`, which is what puts `getSelectedSlides` on the safe part of that surface.",
  "charts:issue:3269":
    "NO EXPOSURE — office.js cannot read speaker notes at all. An element is spliced onto a slide whose notes page is untouched, inside the package; the API gap is real and never reached for.",
  "charts:issue:3309": "NO EXPOSURE — sVG cannot be read back out of a shape. NO EXPOSURE.",
  "charts:issue:3698":
    "RELEVANT — a picture cannot be inserted while a shape is selected, and setSelectedShapes([]) may never resolve. No exposure to the picture half; the selection half is the same class as #2775: this add-in inserts SLIDES, which is neither case, and nobody has established that it is safe (SSF-Merge asks it, unanswered as of 2026-09-08); re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:3826":
    "NO EXPOSURE — a freshly-added slide's layout shapes throw GeneralException. NO EXPOSURE — nothing here reads a slide's layout through the API.",
  "charts:issue:4272":
    "RELEVANT — a collection load of more than ~50 items answers short. The deck read will page `getItemAt` at 20 for the same reason a sibling does; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:4906":
    "NO EXPOSURE — slideLayout.shapes throws on decks built from a custom template. NO EXPOSURE — see #3826.",
  "charts:issue:5455":
    "NO EXPOSURE — generalException reading ParagraphFormat.horizontalAlignment. NO EXPOSURE — text is replaced in the package, never through a text range.",
  "charts:issue:6079":
    "RELEVANT — PowerPoint on the WEB uppercases tag keys internally and then requires the uppercased spelling to read them back. Nothing here writes a tag yet; any key this add-in writes must be uppercase from the first day; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:2474":
    "RELEVANT — `SlideRange.id` lacks the `#XYZ` suffix the deck's own list carries. Whatever names the slide the user is on must match by prefix and refuse two matches rather than guess; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:3565":
    "NO EXPOSURE — context.sync taking progressively longer every run. NO EXPOSURE, and the sibling's own note says why: it is WORD FOR MAC, not PowerPoint web, so it is not evidence about this host either.",
  "charts:issue:6867":
    "RELEVANT — `Slide.exportAsBase64` omits modern comments and `ppt/authors.xml`, and SSF-Merge, which clones from the presentation-level export, has a probe question open on whether that call drops the same parts (unanswered as of 2026-09-08). Which read of the deck this add-in uses is an open question in CLAUDE.md, and this is one of the facts that will decide it; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:3784":
    "RELEVANT — shape TAGS are lost when a shape is cut and pasted on PowerPoint web. Documented rather than guarded, as both siblings decided: nothing this add-in writes onto a shape may be something a later step depends on finding; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:6266":
    "NO EXPOSURE — getImageAsBase64 differs between Mac and Windows for content add-ins. NO EXPOSURE — nothing here rasterises anything.",
  "charts:issue:6498":
    "RELEVANT — shapes inserted on the web may appear in the slide PREVIEW but not the main view without a refresh. A support answer rather than a defect to fix: a user reporting that an element is missing may be seeing this, and the deck delta will say it landed; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:issue:5022":
    "NO EXPOSURE — context.sync runs indefinitely when shapes are re-read after an image insert. NO EXPOSURE — no image inserts and no shape re-reads.",
  "charts:issue:5101":
    "NO EXPOSURE — a placeholder keeps `type: Placeholder` when reused. NO EXPOSURE — nothing here reads a shape's type.",
  "charts:issue:5264":
    "NO EXPOSURE — a part of the object model Office.js cannot reach. NO EXPOSURE — recorded as a limitation on both sides.",
  "charts:issue:5849": "NO EXPOSURE — shape.group throws GeneralException. NO EXPOSURE — no grouping.",
  "charts:issue:5896": "NO EXPOSURE — reported alongside another SVG defect. NO EXPOSURE.",
  "charts:issue:6363":
    'RELEVANT — `PowerPoint.run`\'s batching fails to load properties reliably after `context.sync()`, web only. The deck read will batch `load("id")` across `getItemAt` handles, which is precisely this shape; re-triage when the insert lands; no host code in this repo as of 2026-09-08.',
  "charts:issue:2714": "NO EXPOSURE — setSelectedDataAsync converts points to pixels. NO EXPOSURE — never called.",
  "charts:question:getcount-populates-same-sync":
    "RELEVANT — a host whose count is right while the list is empty. The deck read will trust the scalar count and page by index, never a collection load; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:getitemat-past-end":
    "RELEVANT — paging by index and removing the replaced slide by index both step past the end if the deck moved; what the host does there bounds both; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:which-end-a-short-read-drops":
    "RELEVANT — office-js#4272 again: a short read that is not a prefix makes a slide NUMBER wrong rather than merely a list shorter, and this add-in addresses the slide the user is on by position; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:how-many-collection-reads-a-context-survives":
    "RELEVANT — the deck read will make one `PowerPoint.run` per page, deliberately, so no context accumulates reads; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:delete-then-lookup":
    "ADOPTED as doctrine (CLAUDE.md host rules), no code yet — the replaced slide is removed by position, highest index first, and the deck is re-counted rather than the call believed; whether a deleted slide still resolves is exactly what made by-id clean-up unsafe; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:scratch-slides-returned":
    "ADOPTED as doctrine (CLAUDE.md host rules), no code yet — a sibling's by-id clean-up left 45 blank slides; any removal here is positional and clamped; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:shapes-by-index-vs-items":
    "NO EXPOSURE — a question about a SHAPE collection. Nothing here reads shapes.",
  "charts:question:creationid-on-fresh-shape":
    "NO EXPOSURE — a shape's creation id read back through the API. A spliced slide keeps the creation ids its markup carries, and nothing asks the host for one.",
  "charts:question:creationid-survives-a-sync": "NO EXPOSURE — see `creationid-on-fresh-shape`.",
  "charts:question:creationid-survives-grouping":
    "NO EXPOSURE — see `creationid-on-fresh-shape`; there is no grouping here either.",
  "charts:question:load-isnullobject-populates":
    "NO EXPOSURE — this add-in will never call `getItemOrNullObject`. Slides are reached by `getItemAt`, which is a different code path.",
  "charts:question:load-id-populates-isnullobject": "NO EXPOSURE — see `load-isnullobject-populates`.",
  "charts:question:getitemornullobject-missing": "NO EXPOSURE — see `load-isnullobject-populates`.",
  "charts:question:shape-add-fresh-slide-proxy": "NO EXPOSURE — adds a shape to a slide proxy. No shape work here.",
  "charts:question:shape-add-held-slide-proxy":
    "NO EXPOSURE — see `shape-add-fresh-slide-proxy`. Also in UNSTABLE_ANSWERS, so nothing may be built on its answer in either repo.",
  "charts:question:shape-add-held-slide-proxy-again": "NO EXPOSURE — the partner of `shape-add-held-slide-proxy`.",
  "charts:question:shape-resolve-held-slide-proxy": "NO EXPOSURE — resolving a held shape proxy.",
  "charts:question:shape-add-fresh-getitem-slide": "NO EXPOSURE — see `shape-add-fresh-slide-proxy`.",
  "charts:question:shape-add-positional-slide-proxy": "NO EXPOSURE — see `shape-add-fresh-slide-proxy`.",
  "charts:question:shape-proxy-survives-one-sync":
    "NO EXPOSURE — no proxy is held across a sync here; an insert holds only slide indices.",
  "charts:question:shapes-items-count-honest":
    "NO EXPOSURE to the SHAPE collection. The equivalent question about the SLIDE collection is the one that matters here, and it is `issue:4272`'s.",
  "charts:question:shapes-items-via-positional-slide": "NO EXPOSURE — see `shapes-items-count-honest`.",
  "charts:question:tags-add-same-key-twice":
    "NO EXPOSURE to the API question. The package equivalent is real: if this add-in ever writes a tag part it takes the next free `tagN` rather than overwriting `tag1.xml`, which would destroy another tool's tags. Nothing writes one today.",
  "charts:question:tags-on-fresh-shape":
    "NO EXPOSURE — tags are written into the package, never onto a shape through the API.",
  "charts:question:tag-through-refetched-shape": "NO EXPOSURE — see `tags-on-fresh-shape`.",
  "charts:question:how-many-syncs-a-creation-handle-survives":
    "NO EXPOSURE — no handle from a creation call is ever held here; an insert returns nothing we keep.",
  "charts:question:collection-read-poisons-the-creation-handle":
    "NO EXPOSURE — see `how-many-syncs-a-creation-handle-survives`.",
  "charts:question:does-a-failed-group-poison-the-tag": "NO EXPOSURE — no grouping.",
  "charts:question:addgroup-returns-usable": "NO EXPOSURE — no grouping.",
  "charts:question:group-children-via-getcount": "NO EXPOSURE — no grouping.",
  "charts:question:group-reports-its-children": "NO EXPOSURE — no grouping.",
  "charts:question:group-of-existing-shape-readable": "NO EXPOSURE — no grouping.",
  "charts:question:binding-names-shape-later":
    "NO EXPOSURE — bindings as a route around id refusals, REJECTED on a sibling's evidence: the host rejects the batch carrying the binding, with a control arm proving it was the binding. Nothing here binds.",
  "charts:question:picture-then-shape-read": "NO EXPOSURE — no picture inserts.",
  "charts:question:slide-layout-readable": "NO EXPOSURE — nothing here reads a layout through the API.",
  "charts:question:layouts-readable": "NO EXPOSURE — see `slide-layout-readable`.",
  "charts:question:untrack-available":
    "NO EXPOSURE — an insert holds a handful of proxies for one batch, so there is nothing to untrack. A sibling measured it unavailable on this host anyway.",
  "charts:question:untrack-available-on-shape": "NO EXPOSURE — see `untrack-available`.",
  "charts:issue:6329":
    "RELEVANT — PowerPoint on the web forces a full presentation save on every `context.sync()`, read-only syncs included. Every host call this add-in will make is a sync, so they stay few and batched; re-triage when the insert lands; no host code in this repo as of 2026-09-08.",
  "charts:question:rotation-keeps-the-unrotated-box":
    "NO EXPOSURE — what the API reports as a rotated shape's box. Nothing here reads a shape through the API; an element arrives as slide markup in the package.",
  "charts:question:named-preset-resolves":
    "NO EXPOSURE — whether a geometric preset name resolves when drawn through the API. Nothing here draws through the shape collection.",
};

/**
 * The closed vocabulary every `TRIAGED` reason opens with.
 *
 * A prefix rather than prose, because something has to be able to ASK which
 * findings we acted on. `docs/SIBLING.md` is the human record and it can fall
 * behind this table silently; `test/sibling.test.ts` asserts the ledger
 * mentions every finding that is not NO EXPOSURE.
 */
export const VERDICTS = ["NO EXPOSURE", "ADOPTED", "RELEVANT"];

/** @param {unknown} reason */
export function verdictOf(reason) {
  return VERDICTS.find((v) => typeof reason === "string" && reason.startsWith(v));
}

/**
 * Findings present upstream with no row in `TRIAGED`.
 *
 * @param {{source: {repo: string, space: string, table: string, kind: string}, keys: string[]}[]} tables
 * @param {Record<string, string>} [triaged]
 */
export function untriaged(tables, triaged = TRIAGED) {
  /** @type {Map<string, {id: string, space: string, kind: string, key: string, tables: string[]}>} */
  const seen = new Map();
  for (const { source, keys } of tables) {
    for (const key of keys) {
      const id = idOf(source, key);
      if (Object.prototype.hasOwnProperty.call(triaged, id)) continue;
      const where = `${source.repo}/${source.table}`;
      const hit = seen.get(id);
      if (hit) hit.tables.push(where);
      else
        seen.set(id, {
          id,
          space: source.space,
          kind: source.kind === "self" ? key.split(":")[0] : source.kind,
          key,
          tables: [where],
        });
    }
  }
  return [...seen.values()];
}

/**
 * Every table, from whatever `read` hands back for a repo and a path.
 *
 * @param {(repo: string, path: string) => string} read
 */
export function tablesFrom(read) {
  /** @type {Map<string, string>} */
  const cache = new Map();
  const out = [];
  for (const source of SOURCES) {
    const cacheKey = `${source.repo}/${source.path}`;
    if (!cache.has(cacheKey)) cache.set(cacheKey, read(source.repo, source.path));
    const text = cache.get(cacheKey);
    let keys = tableKeys(text, source.table);
    if (keys === null) {
      if (source.mayBeEmpty && tableDeclared(text, source.table)) keys = [];
      else
        throw new Error(
          `${cacheKey}: no keys read from ${source.table} — it was renamed, moved, emptied or reformatted upstream. ` +
            `A sweep that cannot read the table must say so rather than report a quiet week.`,
        );
    }
    out.push({ source, keys });
  }
  return out;
}

function report(found) {
  if (!found.length) return "Nothing new: every finding in the siblings' curated tables has a row in `TRIAGED`.\n";
  const issues = found.filter((f) => f.kind === "issue");
  const questions = found.filter((f) => f.kind !== "issue");
  const lines = [
    "## Sibling findings nobody here has answered",
    "",
    "Each needs a row in `TRIAGED` in `scripts/sibling-watch.mjs` and, unless it is",
    "no exposure, a line in `docs/SIBLING.md`. **“No exposure” is a real answer** —",
    "say it rather than leaving the row out, or this comes back next Monday looking",
    "identical.",
    "",
  ];
  if (issues.length) {
    lines.push(`### office-js issues (${issues.length})`, "");
    for (const f of issues) {
      const n = f.key.replace(/^issue:/, "");
      lines.push(`- [#${n}](https://github.com/OfficeDev/office-js/issues/${n}) — ${f.tables.join(", ")}`);
    }
    lines.push("");
  }
  if (questions.length) {
    lines.push(`### Host questions (${questions.length})`, "");
    for (const f of questions) lines.push(`- \`${f.id}\` — ${f.tables.join(", ")}`);
    lines.push("");
  }
  return lines.join("\n");
}

async function main(argv) {
  const json = argv.includes("--json");
  const fromAt = argv.indexOf("--from");
  const dir = fromAt >= 0 ? argv[fromAt + 1] : undefined;
  const read = dir
    ? (repo, path) => readFileSync(join(dir, repo, String(path.split("/").pop())), "utf8")
    : await (async () => {
        const bodies = new Map();
        for (const source of SOURCES) {
          const key = `${source.repo}/${source.path}`;
          if (bodies.has(key)) continue;
          const res = await fetch(`${RAW[source.repo]}/${source.path}`);
          if (!res.ok) throw new Error(`${key}: the sibling repo answered ${res.status}`);
          bodies.set(key, await res.text());
        }
        return (repo, path) => String(bodies.get(`${repo}/${path}`));
      })();

  const tables = tablesFrom(read);
  const found = untriaged(tables);
  process.stdout.write(
    json
      ? `${JSON.stringify({ found, tables: tables.map((t) => ({ ...t.source, keys: t.keys.length })) }, null, 2)}\n`
      : report(found),
  );
  return found.length ? 3 : 0;
}

if (isMain(import.meta.url)) {
  main(process.argv.slice(2)).then(
    (code) => process.exit(code),
    (err) => {
      process.stderr.write(`sibling-watch: ${err instanceof Error ? err.message : String(err)}\n`);
      process.exit(1);
    },
  );
}
