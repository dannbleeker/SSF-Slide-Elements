#!/usr/bin/env node
/**
 * Change the code on purpose and see whether the suite notices.
 *
 * Coverage answers "did a test RUN this line". It cannot answer "would a test
 * FAIL if this line were wrong", and the gap between those two is where a
 * vacuous assertion lives. Three were found by hand in this repo in one week —
 * a flag whose removal changed nothing, a four-way check that only ever
 * exercised its first arm, and a count compared against itself — each in a file
 * at or near 100% coverage. This is that hand check, run over every line it can
 * reach.
 *
 * **Not a gate, and deliberately not wired into CI.** A mutation score in a
 * required check becomes a number people move rather than a question people
 * answer, and the whole value here is the list of survivors, which a human
 * reads one at a time. `test/release.test.ts` holds CI's step list; this script
 * is absent from it on purpose.
 *
 * Stryker is the off-the-shelf answer and was not taken: a runtime dependency
 * and a config format, against a script this repo can read end to end, for a
 * job whose output is a report rather than a verdict.
 *
 * ## What it mutates
 *
 * The pure decision code only — `src/host` and the flat pure files of
 * `src/pane`. Nothing that talks to Office.js, and nothing in `src/core`, whose
 * tests are the slow half of the suite.
 *
 * The operators are the mistakes this repo has actually made, not a textbook's
 * list: a comparison boundary, a boolean operator, a dropped negation, a
 * deleted `??` fallback, a guard clause that stops guarding, and a number off
 * by one.
 *
 * ## What it does NOT do
 *
 * It is a text mutator with a hand-written mask, not a parser. It knows where
 * comments and string literals are and leaves them alone; it does not know
 * types, so a mutation that will not compile is simply killed by the run like
 * any other. That is the safe direction — a mutant reported dead when it never
 * ran is a wasted second, while a mutant reported ALIVE that never ran would be
 * a lie. Nothing here can produce the second.
 *
 * ## Two tiers, and the reason
 *
 * Measured 2026-09-13 on an idle machine: the whole suite is 57.6 s, and
 * `test/splice.test.ts` alone is 50.7 of them, because it harvests both 1.5 MB
 * library decks and sweeps 117 elements. That one file IS the suite's
 * wall-clock; everything else fits in its shadow across four workers.
 *
 * These figures replace the ones this file carried until today — 166 s and 82 s,
 * dated 2026-09-12 — which are about 2.9x too high. They were taken while the
 * sweep, a coverage run and CI polling shared the same four cores, and were
 * written down as if they were a property of the suite. A number copied from a
 * live counter carries the date it was taken; it also has to be taken on a
 * machine that is doing nothing else, or the date does not save it.
 *
 * The files a mutation of `src/host` can possibly affect are all in the fast
 * half. So each mutant runs against the fast tier, and then every SURVIVOR is
 * re-run against the whole suite before it is reported — because a survivor of a
 * partial run is not a survivor, it is an untested guess.
 *
 * Usage: `node scripts/mutants.mjs [--only <substring>] [--list]`
 */
import { execFileSync } from "node:child_process";
import { appendFileSync, existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { isMain } from "./is-main.mjs";
import { failedNames } from "./test-count.mjs";

/** The pure decision code. Everything here is covered by the fast tier. */
export const TARGETS = [
  "src/host/capability.ts",
  "src/host/coalesce.ts",
  "src/host/errors.ts",
  "src/host/insert.ts",
  "src/host/jump.ts",
  "src/host/links.ts",
  "src/host/memory.ts",
  "src/host/probe.ts",
  "src/host/theme.ts",
  "src/host/timeout.ts",
  "src/pane/card.ts",
  "src/pane/search.ts",
  "src/pane/steps.ts",
  "src/pane/storage.ts",
  "src/pane/used.ts",

  // The engine. Added 2026-09-13, once `FAST` made it affordable: without a
  // first-tier map these mutations cost about seventeen hours, because every
  // core file except `catalogue/cut.ts` is imported into `test/splice.test.ts`.
  // `cut.ts` has no `FAST` row because it needs none — one test reaches it.
  //
  // `src/core/catalogue/types.ts` was here and is not: it declares interfaces
  // and holds no executable line, so once the boundary operator stopped
  // matching type arguments it offered nothing to mutate. Its ten "survivors"
  // were all the tool changing `Record<string, string>` into
  // `Record<=string, string>`, which vitest erases before anything runs.
  "src/core/catalogue/boxes.ts",
  "src/core/catalogue/cut.ts",
  "src/core/catalogue/harvest.ts",
  "src/core/catalogue/runs.ts",
  "src/core/catalogue/tags.ts",
  "src/core/catalogue/text.ts",
  "src/core/pptx/base64.ts",
  "src/core/pptx/clone.ts",
  "src/core/pptx/layout.ts",
  "src/core/pptx/pkg.ts",
  "src/core/pptx/tags.ts",
  "src/core/pptx/theme.ts",
  "src/core/pptx/xml.ts",
  "src/core/splice/carry.ts",
  "src/core/splice/colours.ts",
  "src/core/splice/landing.ts",
  "src/core/splice/listing.ts",
  "src/core/splice/remove.ts",
  "src/core/splice/shapes.ts",
  "src/core/splice/splice.ts",
];

/**
 * Mutations that CANNOT be killed, with the proof for each.
 *
 * A record, NOT a suppression list. Every one of these still appears in the
 * report; they are only labelled, so the next reader does not spend an afternoon
 * re-deriving what this afternoon already settled. `docs/SIBLING.md` keeps a
 * borrowed finding dated for the same reason.
 *
 * Established on 2026-09-12, by making the mutation and looking
 * at what changed rather than by reasoning about the code. Matching is on file,
 * operator and text — NOT on line number, because a line number rots on the next
 * edit and a ledger that rots silently is worse than none.
 *
 * An entry that stops matching anything is reported too, as a line to re-verdict:
 * either the code moved under it or a test now kills it, and both mean the
 * reasoning below needs reading again.
 */
export const EQUIVALENT = [
  {
    file: "src/host/jump.ts",
    what: "boundary",
    was: "<",
    why: "`split`'s `at < 0`, which routes an id carrying a '#' against one with none. Re-verdicted 2026-09-13: the earlier note described the line as `indexOf(\"#\") > 0`, which this file has never contained — the measurement stood, the description did not. First established over 1,003,578 id pairs against the original `split`, then re-checked over 5,476 pairs built from every id shape the harvest produces after `split` was given a `[string, boolean]` signature. Zero disagreements both times, and the second number is the one that describes the code as it stands. The read-back that the borrowed setSelectedSlides measurement rests on is line 97 and the empty-anchor guard, and both are pinned.",
  },
  {
    file: "src/host/jump.ts",
    what: "off-by-one",
    was: "0",
    why: "the `0` in that same `at < 0`, shifted to `at < 1`: an id whose only '#' sits at position 0 has an empty prefix, and the `ap === \"\"` guard on the line below rejects it either way. A mutation of the comparison in the other direction IS killed by test/jump.test.ts, so the routing is held; this one cannot be observed.",
  },
  {
    file: "src/host/probe.ts",
    what: "boundary",
    was: ">",
    why: "line 85 is only reached after the line above returned for `landed === o.expected`, so `>` and `>=` agree on every remaining input, NaN and -0 included. The branch itself is live and covered.",
  },
  {
    file: "src/pane/search.ts",
    what: "guard",
    was: "return true",
    why: 'an early-out that returns exactly what the code below it would return: with an empty query `"".split(/\\s+/)` is `[""]` and every string includes `""`. Corrected 2026-09-13 — the earlier note said it runs on every keystroke, which is backwards. It is TAKEN only when the box is EMPTY; while a user types, the comparison falls straight through. What it saves is building the haystack for all 117 elements on every render in the pane\'s default browsing state, measured at ~31 microseconds. That is a real fast path on a real state, not the dead defensive branch CLAUDE.md deletes on sight, so the line stays.',
  },
  {
    file: "src/pane/search.ts",
    what: "off-by-one",
    was: "0",
    why: "a counter whose values only ever meet each other in a comparison, so adding one to every count leaves the ordering identical — and the function returns names, not counts.",
  },
  {
    file: "src/pane/search.ts",
    what: "off-by-one",
    was: "1",
    why: "two cases: the same counter as above, and a matrix row seed whose last column is never read.",
  },
  {
    file: "src/pane/storage.ts",
    what: "boundary",
    was: ">",
    why: 'killable only by `Object.is(storedScroll({ scroll: -0 }), 0)`. A -0 cannot reach storage — the pane\'s own write cannot produce one and `JSON.stringify(-0)` is "0" — and -0 behaves identically to 0 in every use a scroll offset has, so the assertion would pin the sign of a zero and go red for a rewrite that changed no behaviour.',
  },
  {
    file: "src/core/catalogue/harvest.ts",
    what: "boundary",
    was: "<",
    why: "both index loops in `relIdsIn`, over `node.attributes` and over `node.childNodes`. Widened to `<=`, each runs one extra turn at `i === length`, and @xmldom/xmldom answers `null` for an index past the end of either collection — measured directly on 2026-09-13, `attributes.item(length)` and `childNodes.item(length)` are both `null`. The `if (attr && …)` and `if (child && …)` guards on the next line absorb it, so nothing is read and nothing is added. Killing it would mean asserting on a DOM read that never happens.",
  },
  {
    file: "src/core/catalogue/cut.ts",
    what: "boundary",
    was: "<",
    why: "`clamp01`'s `n < 0 ? 0 : …`. The two comparisons disagree on exactly one input, `n === 0`, which is true of +0 and -0: at +0 both give +0, and at -0 the original returns -0 where the mutant returns the literal +0 — the same number to every operator except `Object.is`. Measured 2026-09-13 over 3,020,012 values (11 specials: ±0, ±1, ±Infinity, NaN, ±Number.MIN_VALUE, 1±EPSILON; 3,000,000 uniform randoms in [-2, 2]; 20,001 exact twenty-thousandths of a page): zero disagreements under `===`, one under `Object.is`, at -0. It is unreachable as well as unobservable — `withAir` cannot make a -0, the committed catalogue's boxes hold none, and `JSON.stringify(-0)` is \"0\" — so the assertion would pin the sign of a zero, which the `storage.ts` entry above rejects for the same reason.",
  },
  {
    file: "src/core/catalogue/cut.ts",
    what: "boundary",
    was: ">",
    why: "`clamp01`'s `n > 1 ? 1 : n`. They disagree only at `n === 1`, where the original returns `n` — which IS 1 — and the mutant returns the literal 1. Unlike the `<` above there is no -0 caveat, because 1 has one representation: over the same 3,020,012 values on 2026-09-13, zero disagreements under `===` AND zero under `Object.is`. Nothing can observe it.",
  },
  {
    file: "src/core/catalogue/harvest.ts",
    what: "fallback",
    was: "?? 0",
    why: 'the slide size: `Number(sldSz?.getAttribute("cx") ?? 0)` and the `cy` beside it. The fallback has exactly two inputs that reach it. With no `<p:sldSz>` at all the optional chain gives undefined, so the mutant computes `Number(undefined)` = NaN where the original computes 0; with a `<p:sldSz>` carrying no `cx`, `getAttribute` returns null and `Number(null)` is 0, which is the fallback\'s own value, so there the two are identical. NaN and 0 both fail `if (!(width > 0 && height > 0))` on the very next line, which throws `HarvestError("the deck states no slide size")` before either number is read anywhere else — same throw, same message, same problem list. Verified 2026-09-13. Every OTHER mutation of those three lines is killed by test/catalogue.test.ts "the size the deck states".',
  },
  {
    file: "src/core/catalogue/harvest.ts",
    what: "fallback",
    was: "?? key",
    why: 'the two English-name fallbacks, `name: name ?? key` in `category` and `return name ?? key` in `nameOf`. Each is taken exactly when `options.names` has no English name for that key — which is the same condition as the line above it, where a problem is pushed onto `problems`, and a non-empty `problems` throws `HarvestError` before the catalogue is returned (twice: after the slide loop and again after the carried parts). So on every input where the fallback and the mutant differ, there is nothing to observe: the harvest fails either way with the same list. They stay because they keep the type honest — the catalogue\'s `name` fields are `string`, not `string | undefined`, and the mutant only runs at all because vitest strips types rather than checking them. The problems themselves are pinned by test/catalogue.test.ts "fails naming every key and category without an English name, and a key used twice".',
  },
  {
    file: "src/core/catalogue/boxes.ts",
    what: "boolean",
    was: "||",
    why: "`topLevelShapes`' `if (!node || node.nodeType !== 1) continue`. `node` is never falsy inside the loop — the bound is `i < tree.childNodes.length` and @xmldom/xmldom answers null only outside 0..length-1 — so `&&` makes the whole condition permanently false, which is the guard deleted. See the `guard`/`continue` entry below for why deleting it changes nothing. Both measured 2026-09-13. This entry is keyed on file and operator, so it would also label a future `||` survivor elsewhere in boxes.ts; the other two `||`s in the file, on the rotation clamp and the table-column read, are killed by test/boxes.test.ts.",
  },
  {
    file: "src/core/catalogue/boxes.ts",
    what: "guard",
    was: "continue",
    why: "the same early exit in `topLevelShapes`, dropped. Everything after it is a cast and a test of `el.localName` against six element names. Measured 2026-09-13 against @xmldom/xmldom by parsing a `<p:spTree>` holding a text node, a comment, a CDATA section and a processing instruction beside a `<p:sp>`: `localName` is null on all four non-elements, `Array.prototype.includes(null)` is false, so a fallen-through node is not pushed and no property access throws. A declared entity expands to a text node rather than an EntityReference, so those four and Element are the whole domain. The line is NOT deleted: it is what makes the `as Element` cast below it honest, and removing it would rest the walk on an implementation detail of a foreign DOM.",
  },
];

/**
 * Whether a ledger entry is about a given survivor line.
 *
 * `judgeSurvivor` and `staleEquivalents` are the two readings of the same
 * question — "is this recorded?" and "is anything still recording this?" — and
 * they were written out separately, so a fix to one silently left the other
 * matching a different set. One function now, called by both.
 *
 * @param {{ file: string, what: string, was: string }} one
 * @param {string} where a survivor line as the report prints it
 * @returns {boolean}
 */
function matchesEntry(one, where) {
  return where.startsWith(`${one.file}:`) && where.includes(`  ${one.what}  ${JSON.stringify(one.was)} -> `);
}

/**
 * A survivor annotated with the reason it cannot be killed, when there is one.
 *
 * @param {string} where a survivor line as the report prints it
 * @returns {{ known: boolean, why: string }}
 */
export function judgeSurvivor(where) {
  // The operator and the ORIGINAL text together, as one segment ending in the
  // arrow — not two separate `includes`. A report line reads
  // `file:12  boundary  "<=" -> "<"`, and `includes('"<"')` matches that line
  // as well as the one it is about, so an entry recorded for `<` would have
  // labelled a future `<=` survivor a known equivalent. Nothing was mislabelled
  // when this was found on 2026-09-13 — `cut.ts`'s `"<=" -> "<"` is killed —
  // but a ledger that can quietly annotate the wrong line is worse than none.
  const match = EQUIVALENT.find((one) => matchesEntry(one, where));
  return match ? { known: true, why: match.why } : { known: false, why: "" };
}

/**
 * Which recorded entries matched nothing this run, and so need re-verdicting.
 *
 * `swept` is the files the run actually covered, and leaving it out was a defect
 * in this gate rather than a nicety. `--only host/jump,host/probe,pane/search`
 * on 2026-09-13 reported `src/pane/storage.ts` as an entry gone stale, which was
 * nonsense: storage.ts was not swept, so of course nothing of its matched. A
 * gate that cries wolf on every partial run teaches the reader to skim past it,
 * which is the same failure as a gate that cannot fail at all.
 *
 * An entry for a file outside the run is not evidence either way, so it is
 * simply not judged.
 *
 * @param {string[]} survivors
 * @param {string[]} swept the source files this run mutated
 * @returns {string[]}
 */
export function staleEquivalents(survivors, swept) {
  return EQUIVALENT.filter(
    (one) => swept.includes(one.file) && !survivors.some((where) => matchesEntry(one, where)),
  ).map((one) => `${one.file}  ${one.what}  ${JSON.stringify(one.was)}`);
}

/**
 * Which test files can reach a source file, by following imports.
 *
 * The first version of this ran every mutant against a "fast tier" — the whole
 * suite minus the nine files that harvest a real library deck. Measured
 * 2026-09-13: that tier is 15.4 seconds, and 348 mutations of it is an hour and
 * a half. Running only the tests that can actually SEE the mutated file is 1 to
 * 3 seconds, and it is also the more honest question: a mutation of
 * `src/host/insert.ts` that `test/splice.test.ts` fails to notice tells nobody
 * anything.
 *
 * The safety is not in this map, and that is not a formality. SIX test files
 * read source as TEXT rather than importing it, so this function cannot see
 * them: `docs.test.ts` and `security.test.ts` both read `src/pane/main.ts`,
 * `security.test.ts` also reads `links.ts`, `memory.ts`, `steps.ts` and
 * `storage.ts`, `dead-exports.test.ts` reads `src/pane/catalogue.ts`, and
 * `sibling.test.ts` reads `src/office/powerpoint.ts`. Any of them can fail on a
 * mutation this map says is invisible to them.
 *
 * The safety is in the whole suite being re-run against every survivor before it
 * is reported, so an import this misses costs a wasted minute and cannot produce
 * a false survivor. Those six files are the receipt for why that re-check is not
 * an optimisation waiting to be removed.
 *
 * @param {string} from a file path, relative to the repository root
 * @returns {string[]} the source paths it imports, resolved and relative
 */
export function importsOf(from) {
  const text = readFileSync(from, "utf8");
  const here = dirname(from);
  /** @type {string[]} */
  const out = [];
  // Three spellings, because missing one silently shrinks the tests a mutant
  // runs against: `from "…"`, a bare side-effect `import "…"`, and the dynamic
  // `import("…")` the pane's wiring test uses to load `main.ts`.
  for (const m of text.matchAll(/(?:from|import)\s*\(?\s*["']([^"']+)["']/g)) {
    const spec = m[1] ?? "";
    if (!spec.startsWith(".")) continue;
    const base = join(here, spec).replace(/\.js$/, "");
    for (const candidate of [`${base}.ts`, join(base, "index.ts")]) {
      if (existsSync(candidate)) {
        out.push(candidate);
        break;
      }
    }
  }
  return out;
}

/**
 * Every test file from which `target` is reachable.
 *
 * @param {string} target
 * @returns {string[]}
 */
export function testsReaching(target) {
  const tests = readdirSync("test")
    .filter((name) => name.endsWith(".test.ts"))
    .map((name) => join("test", name));
  return tests.filter((test) => {
    const seen = new Set([test]);
    const queue = [test];
    while (queue.length) {
      const next = queue.pop();
      if (next === undefined) break;
      for (const dep of importsOf(next)) {
        if (dep === target) return true;
        if (seen.has(dep)) continue;
        seen.add(dep);
        queue.push(dep);
      }
    }
    return false;
  });
}

/**
 * The tests a mutant of a given file runs against in the FIRST tier, where that
 * is cheaper than everything the import map reaches.
 *
 * `testsReaching` keeps its job: it defines the SOUND set, and it is the
 * fallback for any file with no entry here. This map is an optimisation and
 * nothing more, and it is safe for exactly one reason — every survivor is still
 * re-run against the WHOLE suite before it is reported. A file whose entry here
 * is too narrow costs a wasted minute on a false first-tier survivor; it cannot
 * put a survivor in the report that is not one.
 *
 * The reason it exists: `test/splice.test.ts` is 50.7 s of a 57.6 s suite,
 * because it harvests both 1.5 MB library decks and sweeps 117 elements. Every
 * file in `src/core` except `catalogue/cut.ts` is imported into it, so the
 * "reaching set" of a core file is the whole suite in all but name — `pkg.ts`
 * pulls 23 test files at 62.2 s, which is SLOWER than simply running everything.
 * Measured 2026-09-13: extending the sweep to `src/core` without this map costs
 * about 17 hours of first tier. With it, about an hour.
 *
 * Every row was measured rather than guessed, on 2026-09-13. The method: start
 * from the file's full reaching set, drop each expensive test file in turn, and
 * KEEP the drop only when that file's own statement and branch coverage is
 * unchanged. Then time what survives with no coverage instrumentation, which is
 * what a tier run actually costs. Both numbers are on every row.
 *
 * Sixteen of the nineteen rows are LOSSLESS — identical coverage to the whole
 * suite. Three are not, and are here anyway: `splice/carry.ts` (92.45 branches
 * against 96.22), `splice/colours.ts` (83.33 against 87.5) and
 * `splice/splice.ts` (89.7 against 97.05). Falling back to their sound sets
 * would cost 57.6 s a mutant — 106 minutes for their 112 mutations, more than
 * the other seventeen files put together. A lossy tier costs one 57.6 s
 * whole-suite re-check per mutant it fails to kill, and there will be a handful,
 * not fifty. The re-check is what makes that trade safe; without it the loss
 * would be false survivors in the report rather than wasted minutes.
 *
 * Measured total for the engine's mutations: about 66 minutes of first tier,
 * plus one 57.6 s re-check per survivor. That figure was taken when the boundary
 * operator still matched type arguments, over 998 engine mutations; dropping
 * that noise on 2026-09-13 took the engine to 757 and the whole set to 1073, so
 * the real cost is lower and has not been re-timed.
 */
export const FAST = {
  // 80 mutations, 5.9 s a mutant, 100/100 statements/branches
  "src/core/catalogue/boxes.ts": ["test/boxes.test.ts", "test/catalogue.test.ts", "test/splice-malformed.test.ts"],
  // 110 mutations, 6.2 s a mutant, 96.92/90.51 statements/branches
  "src/core/catalogue/harvest.ts": ["test/catalogue.test.ts", "test/colours.test.ts"],
  // 31 mutations, 6.1 s a mutant, 100/91.66 statements/branches
  "src/core/catalogue/runs.ts": ["test/catalogue.test.ts", "test/runs.test.ts"],
  // 4 mutations, 2.9 s a mutant, 100/100 statements/branches
  "src/core/catalogue/tags.ts": ["test/validators-deck.test.ts"],
  // 21 mutations, 5.8 s a mutant, 100/90 statements/branches
  "src/core/catalogue/text.ts": ["test/catalogue.test.ts", "test/splice-malformed.test.ts", "test/text.test.ts"],
  // 10 mutations, 2.3 s a mutant, 100/100 statements/branches
  "src/core/pptx/base64.ts": [
    "test/base64.test.ts",
    "test/integrity.test.ts",
    "test/package-surface.test.ts",
    "test/package-valid.test.ts",
    "test/pptx-clone.test.ts",
    "test/pptx-layout.test.ts",
    "test/pptx-malformed.test.ts",
    "test/pptx-tags.test.ts",
    "test/pptx.test.ts",
    "test/slide-listing.test.ts",
    "test/splice-malformed.test.ts",
  ],
  // 52 mutations, 1.1 s a mutant, 100/97.01 statements/branches
  "src/core/pptx/clone.ts": ["test/pptx-clone.test.ts", "test/splice-malformed.test.ts"],
  // 71 mutations, 1.1 s a mutant, 100/100 statements/branches
  "src/core/pptx/layout.ts": ["test/pptx-layout.test.ts", "test/splice-malformed.test.ts"],
  // 187 mutations, 3.7 s a mutant, 99.15/98.96 statements/branches
  "src/core/pptx/pkg.ts": [
    "test/integrity.test.ts",
    "test/package-surface.test.ts",
    "test/package-valid.test.ts",
    "test/pptx-clone.test.ts",
    "test/pptx-layout.test.ts",
    "test/pptx-malformed.test.ts",
    "test/pptx-tags.test.ts",
    "test/pptx.test.ts",
    "test/security.test.ts",
    "test/slide-listing.test.ts",
    "test/splice-carry.test.ts",
    "test/splice-malformed.test.ts",
  ],
  // 65 mutations, 1.1 s a mutant, 100/97.5 statements/branches
  "src/core/pptx/tags.ts": ["test/pptx-tags.test.ts", "test/splice-malformed.test.ts"],
  // 38 mutations, 1.1 s a mutant, 96.66/95.83 statements/branches
  "src/core/pptx/theme.ts": ["test/colours.test.ts", "test/splice-malformed.test.ts"],
  // 17 mutations, 2.1 s a mutant, 97.56/84.21 statements/branches
  "src/core/pptx/xml.ts": [
    "test/integrity.test.ts",
    "test/package-surface.test.ts",
    "test/package-valid.test.ts",
    "test/pptx-clone.test.ts",
    "test/pptx-layout.test.ts",
    "test/pptx-malformed.test.ts",
    "test/pptx-tags.test.ts",
    "test/pptx.test.ts",
    "test/slide-listing.test.ts",
    "test/splice-malformed.test.ts",
  ],
  // 46 mutations, 2.7 s a mutant, 100/92.45 statements/branches — LOSSY, see below
  "src/core/splice/carry.ts": ["test/splice-carry.test.ts", "test/splice-malformed.test.ts"],
  // 15 mutations, 1.1 s a mutant, 94.59/83.33 statements/branches — LOSSY, see below
  "src/core/splice/colours.ts": ["test/colours.test.ts", "test/splice-malformed.test.ts"],
  // 55 mutations, 1.2 s a mutant, 100/100 statements/branches
  "src/core/splice/landing.ts": ["test/splice-landing.test.ts", "test/splice-malformed.test.ts"],
  // 8 mutations, 1.1 s a mutant, 100/100 statements/branches
  "src/core/splice/listing.ts": ["test/slide-listing.test.ts", "test/splice-malformed.test.ts"],
  // 27 mutations, 3.3 s a mutant, 98.24/93.1 statements/branches
  "src/core/splice/remove.ts": ["test/splice-malformed.test.ts", "test/splice-remove.test.ts"],
  // 75 mutations, 13.0 s a mutant, 99.5/97.52 statements/branches
  "src/core/splice/shapes.ts": ["test/splice-malformed.test.ts", "test/splice-shapes.test.ts"],
  // 51 mutations, 3.1 s a mutant, 91.07/89.7 statements/branches — LOSSY, see below
  "src/core/splice/splice.ts": ["test/splice-malformed.test.ts", "test/validators-deck.test.ts"],
};

/**
 * The first tier for a file: its measured subset, or everything that reaches it.
 *
 * @param {string} file
 * @returns {string[]}
 */
export function fastTests(file) {
  const named = /** @type {Record<string, string[]>} */ (FAST)[file];
  return named ?? testsReaching(file);
}

/**
 * A copy of `text` in which every comment and string literal is a space, and
 * every code character is itself, at the SAME offset.
 *
 * Offsets have to survive, because a mutation is found in the mask and applied
 * to the original. Any stripper that shortens the text — every one in
 * `without-prose.mjs` does — cannot be used for this.
 *
 * A regex literal is recognised by what precedes it, which is the standard
 * heuristic and not a parser: after a value, `/` divides; after an operator or
 * an opening bracket, it opens a pattern. Getting it wrong masks a little code
 * or reveals a little text, and either way the worst case is a mutant that does
 * not compile, which the run kills.
 *
 * @param {string} text
 * @returns {string}
 */
export function codeMask(text) {
  const out = text.split("");
  const blank = (from, to) => {
    for (let i = from; i < to && i < out.length; i += 1) if (out[i] !== "\n") out[i] = " ";
  };
  let at = 0;
  let lastCode = "";
  while (at < text.length) {
    const two = text.slice(at, at + 2);
    if (two === "//") {
      const end = text.indexOf("\n", at);
      blank(at, end === -1 ? text.length : end);
      at = end === -1 ? text.length : end;
      continue;
    }
    if (two === "/*") {
      const end = text.indexOf("*/", at + 2);
      const stop = end === -1 ? text.length : end + 2;
      blank(at, stop);
      at = stop;
      continue;
    }
    const ch = text[at];
    if (ch === '"' || ch === "'" || ch === "`") {
      let end = at + 1;
      while (end < text.length) {
        if (text[end] === "\\") {
          end += 2;
          continue;
        }
        if (text[end] === ch) break;
        end += 1;
      }
      blank(at, Math.min(end + 1, text.length));
      at = end + 1;
      lastCode = ch === "`" ? "`" : '"';
      continue;
    }
    if (ch === "/" && /[(,=:[!&|?{;+\-*%<>]/.test(lastCode)) {
      let end = at + 1;
      let inClass = false;
      while (end < text.length && text[end] !== "\n") {
        if (text[end] === "\\") {
          end += 2;
          continue;
        }
        if (text[end] === "[") inClass = true;
        else if (text[end] === "]") inClass = false;
        else if (text[end] === "/" && !inClass) break;
        end += 1;
      }
      blank(at, Math.min(end + 1, text.length));
      at = end + 1;
      lastCode = "/";
      continue;
    }
    if (!/\s/.test(ch)) lastCode = ch;
    at += 1;
  }
  return out.join("");
}

/**
 * Every mutation this script can make to one file.
 *
 * Each is `{ at, was, now, what }` — an offset into the ORIGINAL text, the
 * exact characters there, what to put in their place, and the name of the
 * operator, so a survivor names the change rather than a line number.
 *
 * @param {string} text
 * @returns {{ at: number, was: string, now: string, what: string }[]}
 */
export function mutationsOf(text) {
  const mask = codeMask(text);
  /** @type {{ at: number, was: string, now: string, what: string }[]} */
  const found = [];
  const add = (at, was, now, what) => {
    if (mask.slice(at, at + was.length) !== was) return;
    // `was` came out of the MASK, where a string literal is spaces, so the
    // report would show a blanked line. Its LENGTH is what the edit needs; the
    // original text is what a reader needs.
    found.push({ at, was: text.slice(at, at + was.length), now, what });
  };

  // A comparison boundary, and ONLY a comparison. The operator must have
  // whitespace on both sides, which is what separates `a < b` from the `<` of
  // `Record<string, string>`, the `>` of `Promise<T>` and the shifts in
  // `hash >>> 0` and `n >> 16`. Prettier puts spaces around every binary
  // operator and `npm run format:check` is a CI step, so in THIS repo the rule
  // is exact rather than a heuristic.
  //
  // Measured 2026-09-13, over the 36 files of `TARGETS`: 373 sites matched the
  // operator, and 272 of them — 73% — were type arguments or shifts. Every one
  // was a wasted run, and every one of the type-argument ones was guaranteed to
  // be reported ALIVE, because vitest strips types without checking them, so
  // `Record<=string, string>` is erased before anything executes. That is where
  // all ten of `src/core/catalogue/types.ts`'s survivors came from. Checked
  // against the sweep's own report: of the 34 boundary survivors it had found
  // at that point, this rule drops exactly those ten and keeps the other 24.
  for (const m of mask.matchAll(/(^|[^=<>!])(<=|>=|<|>)(?!=)/g)) {
    const at = (m.index ?? 0) + (m[1] ?? "").length;
    const op = m[2] ?? "";
    if (!/\s/.test(mask[at - 1] ?? "") || !/\s/.test(mask[at + op.length] ?? "")) continue;
    /** @type {Record<string, string>} */
    const flip = { "<": "<=", ">": ">=", "<=": "<", ">=": ">" };
    add(at, op, flip[op] ?? op, "boundary");
  }

  // A boolean operator, and a dropped negation.
  for (const m of mask.matchAll(/&&|\|\|/g)) {
    add(m.index ?? 0, m[0], m[0] === "&&" ? "||" : "&&", "boolean");
  }
  for (const m of mask.matchAll(/(^|[^!=<>+\-*/%&|^])!(?![=])/g)) {
    add((m.index ?? 0) + (m[1] ?? "").length, "!", " ", "negation");
  }

  // A `??` fallback deleted, so the left-hand side stands alone. Only where the
  // fallback is one simple token; anything longer needs a parser.
  for (const m of mask.matchAll(/\?\?\s*(?:[A-Za-z_$][\w$.]*|\d+|\[\]|\{\})/g)) {
    add(m.index ?? 0, m[0], "", "fallback");
  }

  // A guard clause that stops guarding: `if (…) return x;` on one line, with
  // the body emptied and the condition left exactly as it was, so what the
  // suite loses is the EARLY EXIT and nothing else.
  for (const m of mask.matchAll(/\)\s*(return[^;\n]*|continue|break);/g)) {
    add((m.index ?? 0) + m[0].indexOf(m[1] ?? ""), m[1] ?? "", "void 0", "guard");
  }

  // A number off by one. Not `0` on its own inside an index, which is every
  // other line — `0` is included anyway, because an off-by-one at zero is the
  // one this repo actually shipped.
  for (const m of mask.matchAll(/(^|[^\w.$])(\d+)(?![\w.])/g)) {
    const at = (m.index ?? 0) + (m[1] ?? "").length;
    const was = m[2] ?? "";
    add(at, was, String(Number(was) + 1), "off-by-one");
  }

  return found.sort((a, b) => a.at - b.at);
}

/**
 * Which line a mutation sits on, for a report a person reads.
 *
 * @param {string} text
 * @param {number} at
 * @returns {number}
 */
export function lineOf(text, at) {
  return text.slice(0, at).split("\n").length;
}

/**
 * What a run said about a mutant: it lived, a test killed it, or the run failed
 * for a reason that was not the mutation.
 *
 * This answered a plain boolean at first — a non-zero exit meant killed — and
 * that is wrong in the direction that matters. A run can exit non-zero without a
 * single test failing: a worker the kernel killed, a transform error, a machine
 * under load. Counting those as kills makes the sweep UNDER-report, and a
 * mutation tool that quietly misses survivors is worse than none, because the
 * clean part of the report is the whole product.
 *
 * Measured 2026-09-12, which is why this exists. The first full sweep reported
 * `src/host/probe.ts:657` killed. Re-running that file after a round of fixes —
 * which can only ever kill MORE mutants, never fewer — reported it alive, and by
 * hand it is alive: the second `&&` on that line, which lets the reader subtract
 * against a count the sheet does not carry. The kill was a run that failed for
 * its own reasons while several other things shared the same four cores.
 *
 * So the JSON report is read back, exactly as `test-count.mjs` learned to do,
 * and a kill counts only when the report NAMES a failed test.
 *
 * A run also gets a WALL-CLOCK LIMIT, and that is the second thing measured the
 * hard way. On 2026-09-13 a first-tier run of the twelve files that reach
 * `pkg.ts` stopped responding: the vitest process and three workers sat blocked
 * with one second of CPU between them for three hours and forty minutes, and
 * because `execFileSync` has no timeout by default the sweep waited with them.
 * Nothing said so — the survivor file simply stopped growing at 910 of 1345, and
 * a reader coming back to it would have read a stalled run as a slow one.
 *
 * The limit is deliberately far above any honest run: the whole suite is 57.6 s
 * and the slowest first tier is about 13 s, so ten minutes is more than ten
 * times the worst case and cannot cut a real run short. A run that hits it is
 * INCONCLUSIVE, never a kill — the same direction as every other unclear
 * answer, and `tryMutation` will give it one more go.
 *
 * @param {string[]|null} files null for the whole suite
 * @param {string} out where the JSON report goes
 * @returns {"survived"|"killed"|"inconclusive"}
 */
export const RUN_LIMIT_MS = 10 * 60 * 1000;

export function verdictOf(files, out) {
  const args = [
    join("node_modules", "vitest", "vitest.mjs"),
    "run",
    "--reporter=json",
    `--outputFile=${out}`,
    "--bail=1",
  ];
  if (files) args.push(...files);
  try {
    execFileSync(process.execPath, args, { stdio: "ignore", timeout: RUN_LIMIT_MS, killSignal: "SIGKILL" });
    return "survived";
  } catch (error) {
    return verdictOfFailure(error, out);
  }
}

/**
 * What a FAILED run means. Its own function so the two ways a run can fail
 * without answering — a timeout, and a report that names nothing — can be held
 * by a test without waiting ten minutes for a real one.
 *
 * A timeout must not read the report: the file on disk is whatever the PREVIOUS
 * run left there, so reading it would answer this mutant with the last one's
 * result. That is the same shape of mistake as counting a non-zero exit a kill,
 * and it fails in the same bad direction.
 *
 * @param {unknown} error what `execFileSync` threw
 * @param {string} out where the JSON report goes
 * @returns {"killed"|"inconclusive"}
 */
export function verdictOfFailure(error, out) {
  if (/** @type {NodeJS.ErrnoException} */ (error)?.code === "ETIMEDOUT") return "inconclusive";
  try {
    return failedNames(JSON.parse(readFileSync(out, "utf8"))).length > 0 ? "killed" : "inconclusive";
  } catch {
    return "inconclusive";
  }
}

/**
 * Put the mutation in place, run the tests, and ALWAYS put the original back.
 *
 * The restore is in a `finally` because a half-written file is the one way this
 * script can do damage. The sweep works in a copy now, but that copy is what
 * every later mutant is applied to, so a botched restore would poison the rest
 * of the run rather than the reader's tree.
 *
 * An inconclusive run is tried once more before it is believed, and once is the
 * limit: a mutant that cannot get a straight answer out of two runs is reported
 * as inconclusive rather than guessed at in either direction.
 *
 * @param {string} file
 * @param {string} text the original
 * @param {string} mutated
 * @param {string[]|null} tests null for the whole suite
 * @param {string} out
 * @returns {"survived"|"killed"|"inconclusive"}
 */
function tryMutation(file, text, mutated, tests, out) {
  writeFileSync(file, mutated);
  try {
    const first = verdictOf(tests, out);
    return first === "inconclusive" ? verdictOf(tests, out) : first;
  } finally {
    writeFileSync(file, text);
  }
}

/**
 * A throwaway copy of the repository to do the mutating in.
 *
 * The first two attempts at a full sweep mutated the working tree itself, and
 * restored each file within a second or two — which is correct and still
 * unpleasant: for the hour the sweep runs, `git status` is never trustworthy,
 * every editor in the directory sees files changing under it, and an
 * interrupted run can leave a source file wrong. None of that is worth saving
 * a copy of 332 MB that takes a few seconds.
 *
 * `cp -a` rather than a git worktree, because the sweep must see the tree AS IT
 * IS — including anything uncommitted and the built `dist-lib` and harvested
 * `public/catalogue` that some tests read. A worktree would silently sweep a
 * different tree from the one the reader is looking at.
 *
 * @returns {string}
 */
function workspace() {
  const dir = mkdtempSync(join(tmpdir(), "ssf-mutants-"));
  execFileSync("cp", ["-a", ".", dir], { stdio: "ignore" });
  return dir;
}

function main() {
  const argv = process.argv.slice(2);
  const onlyAt = argv.indexOf("--only");
  // A comma-separated list, not one substring: re-checking the files a round of
  // fixes touched is the ordinary use, and eight separate runs would each copy
  // the tree again and each re-report the same equivalent mutants.
  const only = (onlyAt === -1 ? "" : (argv[onlyAt + 1] ?? "")).split(",").filter(Boolean);
  const listing = argv.includes("--list");
  const files = TARGETS.filter((f) => only.length === 0 || only.some((part) => f.includes(part)));

  if (!listing) {
    const dir = workspace();
    console.log(`mutants: working in ${dir}, the tree you are in is untouched`);
    const sweep = () => rmSync(dir, { recursive: true, force: true });
    process.on("exit", sweep);
    // A killed sweep is the normal way this ends when a reader has seen enough,
    // and `exit` does not run on a signal. Both of the first two attempts were
    // killed, so this is the ordinary case rather than the exceptional one.
    for (const signal of ["SIGINT", "SIGTERM"]) {
      process.on(signal, () => {
        sweep();
        process.exit(1);
      });
    }
    process.chdir(dir);
  }

  const plan = files.map((file) => {
    const text = readFileSync(file, "utf8");
    return { file, text, mutations: mutationsOf(text), tests: fastTests(file) };
  });
  const total = plan.reduce((sum, one) => sum + one.mutations.length, 0);
  console.log(`mutants: ${total} mutations across ${plan.length} files`);
  // Survivors go to a file as well as the console, because the sweep is over an
  // hour and the console is where an interrupted run's output goes to die.
  const report = join(tmpdir(), "ssf-mutants-survivors.txt");
  if (!listing) {
    writeFileSync(report, "");
    console.log(`mutants: survivors also written to ${report}`);
  }
  if (listing) {
    for (const { file, text, mutations, tests } of plan) {
      console.log(`  ${file} <- ${tests.join(" ") || "NOTHING"}`);
      for (const m of mutations) {
        console.log(`  ${file}:${lineOf(text, m.at)}  ${m.what}  ${JSON.stringify(m.was)} -> ${JSON.stringify(m.now)}`);
      }
    }
    return;
  }

  /** @type {string[]} */
  const survivors = [];
  /** Mutants no run would give a straight answer about. Never counted as killed. */
  const unclear = [];
  const out = join(tmpdir(), "ssf-mutants-report.json");
  let done = 0;
  for (const { file, text, mutations, tests } of plan) {
    if (!tests.length) console.log(`\nmutants: NOTHING IMPORTS ${file} — every mutation of it will survive`);
    for (const m of mutations) {
      const mutated = text.slice(0, m.at) + m.now + text.slice(endOf(m));
      const verdict = tryMutation(file, text, mutated, tests, out);
      done += 1;
      const where = `${file}:${lineOf(text, m.at)}  ${m.what}  ${JSON.stringify(m.was)} -> ${JSON.stringify(m.now)}`;
      if (verdict === "inconclusive") {
        unclear.push(where);
        console.log(`mutants: INCONCLUSIVE  ${where}`);
        continue;
      }
      if (verdict === "killed") {
        if (done % 10 === 0) console.log(`mutants: ${done}/${total}, ${survivors.length} surviving so far`);
        continue;
      }
      // A survivor of the fast tier is not a survivor yet. Re-run it against
      // everything before it goes in the report.
      if (tryMutation(file, text, mutated, null, out) === "survived") {
        survivors.push(where);
        // Written and printed as it is found, not only in the summary. A sweep
        // of the whole set is over an hour, and a run that is interrupted
        // should still have told the reader everything it knew at the time.
        console.log(`mutants: SURVIVED  ${where}`);
        appendFileSync(report, `${where}\n`);
      }
    }
  }
  if (unclear.length) {
    console.log(`\nmutants: ${unclear.length} INCONCLUSIVE — two runs each, neither naming a failed test:`);
    for (const one of unclear) console.log(`  ${one}`);
    console.log("These are NOT kills. Re-run them before believing anything about them.");
  }
  if (!survivors.length) {
    console.log("mutants: no survivors — every mutation the suite could see, it saw");
    return;
  }
  console.log(`mutants: ${survivors.length} survived the WHOLE suite —`);
  for (const one of survivors) {
    const judged = judgeSurvivor(one);
    console.log(`  ${one}${judged.known ? "   [known equivalent, see EQUIVALENT]" : ""}`);
    if (judged.known) console.log(`      ${judged.why}`);
  }
  const stale = staleEquivalents(survivors, files);
  if (stale.length) {
    console.log(`\nmutants: ${stale.length} recorded equivalent(s) matched nothing this run — re-verdict them:`);
    for (const one of stale) console.log(`  ${one}`);
  }
  console.log(
    "\nEach is either a case the suite is missing or a line that does not matter. Both are findings; the second gets deleted.",
  );
}

/**
 * Where a mutation's original text ends.
 *
 * @param {{ at: number, was: string }} m
 * @returns {number}
 */
function endOf(m) {
  return m.at + m.was.length;
}

if (isMain(import.meta.url)) main();
