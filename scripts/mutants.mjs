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
 * types.
 *
 * This paragraph used to say that a mutation which will not compile "is simply
 * killed by the run like any other", and that nothing here could report a
 * mutant ALIVE that never ran. BOTH halves were false, and each was proven so
 * on 2026-09-14:
 *
 * - Vitest transpiles with esbuild, which strips types without checking them,
 *   so a type-only breakage runs clean and was reported as a survivor. `tsc`
 *   is now asked about every survivor before it goes in the report, and one it
 *   refuses is TYPE-KILLED rather than alive.
 * - A mutation can also stop the code terminating — delete a loop's only exit
 *   and the suite never finishes. That was reported INCONCLUSIVE, which is a
 *   bucket nobody reads. It has its own outcome now, and is asked twice, because
 *   a wedged machine looks the same on one ask and a broken loop does not.
 *
 * Four outcomes, then, not two: killed, survived, did not terminate, and
 * type-killed — plus inconclusive for a run that answered nothing at all.
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
import { dirname, join, relative } from "node:path";
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
  // first-tier map these mutations cost about thirteen hours, because every
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
    what: "boundary",
    was: "<",
    why: "`topLevelShapes`' loop bound, `i < tree.childNodes.length`, widened to `<=`. The same shape as the `harvest.ts` entry above: one extra turn at `i === length`, `childNodes.item(length)` is null, and the `if (!node || …) continue` on the next line absorbs it. Worth reading beside its own history — the first engine sweep reported this mutant KILLED, and the re-sweep of 2026-09-13 reported it alive on a suite that had only grown. Nothing about the line got weaker, and the mutant is provably behaviour-identical, so the kill was false: a test that failed under load for its own reasons, in a run sharing four cores with four other agents. `verdictOf` re-checks every SURVIVOR against the whole suite and re-checks no KILL at all, which is the asymmetry that let it through.",
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
  {
    file: "src/core/pptx/pkg.ts",
    what: "boundary",
    was: ">",
    why: "Four copies of one maximum loop — `nextFree` (line 67), `countable`'s `n > 0` (76), `addRel`'s relationship-id scan (437) and `appendSldId`'s slide-id scan (860). Widened to `>=`, each assigns a value equal to the one it already holds; `countable` additionally admits 0, and `nextFree` starts its maximum at 0, so a 0 in the set cannot raise it (a `\\d+` capture cannot be negative). Measured 2026-09-14 by differential snapshot: both committed library decks, five fixture shapes, eight crafted packages, every counter, every content type, every slide list, the held-document count and 187 resolver pairs — 0 disagreements for each of the four. NOTE: this entry is keyed on file and operator, so it would also label a future `>` survivor elsewhere in pkg.ts; the fifth `>` in the file, `appendSldId`'s `id > MAX_SLIDE_ID` on line 864, is killed by test/pptx.test.ts \"uses the highest id the format allows rather than refusing it\".",
  },
  {
    file: "src/core/pptx/pkg.ts",
    what: "boundary",
    was: "<",
    why: "Two root-part comparisons: `ownerOf`'s `i < 0` (line 750) and `resolvePath`'s `slash < 0` (1026). `resolvePath` is identical on ALL inputs — at slash === 0 the else-branch is `slice(0, 0)`, which is the empty string the then-branch returns. `ownerOf` differs only for a rels path beginning with `/_rels/`, and a ZIP item name carries no leading slash (OPC maps a part name to an item name by stripping it), so `indexOf(\"/_rels/\")` is -1 or ≥ 1; `_rels/.rels`, the case the branch exists for, gives -1 either way. Measured 2026-09-14: 0 disagreements over the whole snapshot domain for both. The file's other two `<` survivors are NOT equivalent and are now killed: `relsPathFor`'s `slash < 0` (403), where an absolute part name with no directory takes the wrong arm, and `extensionOf`'s `dot < 0` (1049), where a dotfile loses its extension.",
  },
  {
    file: "src/core/pptx/pkg.ts",
    what: "off-by-one",
    was: "0",
    why: "The same two root-part comparisons shifted (`i < 1` at 750, `slash < 1` at 1026), which are the same mutants as the `<` entry above and unreachable for the same reasons, plus `appendSldId`'s `getAttribute(\"id\") ?? 1` (859) — 1 is no more able to beat a maximum that starts at MIN_SLIDE_ID - 1 = 255 than 0 is. Measured 2026-09-14: 0 disagreements each. The other four `0` survivors in the file ARE killed: `noteWritten`'s `?? 1` (146), `relsPathFor` (403), `addRel`'s `?? 1` (436) and `extensionOf` (1049).",
  },
  {
    file: "src/core/pptx/pkg.ts",
    what: "fallback",
    was: "?? 0",
    why: "All four of the file's `?? 0` fallbacks: `usedNumbers` (136), `noteWritten` (146), `addRel`'s id read (436) and `appendSldId`'s id read (859). Dropping the fallback turns the absent value into NaN where the original made it 0, and every one of the four is immediately compared — `countable(NaN)` is false as `countable(0)` is, and `NaN > max` is false as `0 > max` is for every maximum those loops can hold (0, and 255). Measured 2026-09-14: 0 disagreements over the whole snapshot domain. Two of the four were already recorded BELT AND BRACES by hand in pptx-malformed.test.ts on 2026-09-12; this measures the other two and all four together. They stay because they keep the arithmetic honest — a NaN that happens to lose a comparison is not the same statement as a zero.",
  },
  {
    file: "src/core/pptx/pkg.ts",
    what: "guard",
    was: "continue",
    why: 'The `!rId` early exits in `slidePaths` (575) and `removeSlide` (625). Falling through, `targets.get(null)` is undefined: in `slidePaths` the `if (target)` below drops it, and in `removeSlide` undefined is never equal to the slidePath so the next line skips it. The map cannot be keyed by the empty string either, because `relTargets` refuses an entry with no Id. Measured 2026-09-14: 0 disagreements. pptx-malformed.test.ts labels both BELT AND BRACES by hand, with the behaviour pinned rather than the line. The third `continue` in the file, `relTargets`\' `!id || !target` on line 514, is NOT equivalent — dropping it crashes the resolver on a null Target — and is killed by test/pptx.test.ts "resolves the rest of the slide list past it rather than raising".',
  },
  {
    file: "src/core/pptx/pkg.ts",
    what: "guard",
    was: "return",
    why: 'The referrer scan\'s `if (!path.endsWith(".rels")) return` (line 742), inside `zip.forEach`. Without it the scan also collects JSZip\'s directory entries (`ppt/slides/_rels/`) and any stray file under a `_rels/` folder; `ownerOf` turns each into an owner name whose rels path the package does not hold, so `relatedParts` answers [] immediately, parses nothing and contributes no referrer. Measured 2026-09-14 over both library decks, five fixture shapes and eight crafted packages: package state, content types, slide lists, held-document count and every answer identical — the only difference is four to six extra `relatedParts` calls that read nothing. pptx-malformed.test.ts reached the same verdict by hand on 2026-09-12 ("the extra entry is a duplicate"); this is the measurement over the whole domain. Contrast line 733\'s `return []`, whose removal IS observable — it reads and HOLDS every .rels in the deck — and which is now killed.',
  },
  {
    file: "src/core/pptx/pkg.ts",
    what: "guard",
    was: "return segment",
    why: "`decodeSegment`'s fast path (line 981). `decodeURIComponent` changes nothing in a string containing no `%`, and throws only on a malformed escape — which the try/catch below already answers with the segment itself. Measured 2026-09-14 over every part name in both committed library decks, the fixture and the crafted packages, and 187 owner/target pairs covering `..`, `.`, absolute targets, `%20`, `%2E%2E` and `100%.png`: 0 disagreements. The line is a fast path on the common case, not a decision.",
  },
  {
    file: "src/core/pptx/tags.ts",
    what: "guard",
    was: "continue",
    why: "The two nodeType guards this file has, in `nvPrOf` (line 150) and `idOf` (line 173), and the third in `readShapeTags`'s `walk` (line 330). Established by MEASUREMENT on 2026-09-14, not by reading the code: each mutant was built as its own module and run beside the original over 21 packages — 10 fixture decks (group, picture, pretty-printed picture, graphicFrame, connector, extLst, custData, dangling tag reference, a reference with no r:id, shapes with no id and no cNvPr, an mc:AlternateContent modern chart, a vendor tag part, and a spTree carrying a comment, a processing instruction, a CDATA section and text beside and inside a shape), 11 decks already stamped at top level, inside a group, inside a group inside a group and inside a non-group, and both committed library decks — comparing readShapeTags per slide, usedInDeck, and the markup left by a writeShapeTags pass over every taggable top-level shape. Zero disagreements, and 500 randomised tag distributions added none. The library says why: @xmldom/xmldom answers firstChild null, childNodes.length 0, localName null and namespaceURI null for nodeType 3, 4, 7 and 8, so `child()` reads nothing, `nvPrOf` answers undefined, and the `grpSp` test is false. None of the three is deleted — each is what makes the `as Element` cast on the line below it honest, the same reason the boxes.ts entry gives. Keyed on file and operator, so it would also label a future `continue` guard elsewhere in tags.ts; the `return undefined` guard on line 305 is NOT covered by it, and is killed by a test as of 2026-09-14.",
  },
  {
    file: "src/core/pptx/tags.ts",
    what: "fallback",
    was: "?? undefined",
    why: '`idOf`\'s `return cNvPr.getAttribute("id") ?? undefined` (line 175). `idOf` is module-private with one caller, `shapeId: idOf(shape) ?? ""`, and `null ?? ""` and `undefined ?? ""` are both "". Measured 2026-09-14 over the 21-package corpus above, including the two fixtures whose shapes state no id and carry no `<p:cNvPr>` at all: zero disagreements, read and write. It stays because it keeps the declared `string | undefined` true under `strict`; vitest strips types rather than checking them, which is the only reason the mutant runs.',
  },
  {
    file: "src/core/pptx/tags.ts",
    what: "boundary",
    was: "<",
    why: "`usedInDeck`'s slide loop, `for (let i = 0; i < paths.length; i++)` widened to `<=` (line 386). One extra turn at `i === length`, where `paths[i]` is undefined and `readShapeTags` asks `pkg.has(undefined)` — JSZip's `file(undefined)` returns null, measured directly on 2026-09-14 — so it returns [] before reading anything and `found` is untouched. Zero disagreements over the 21-package corpus and 500 randomised tag distributions. The same shape as the loop-bound entries already recorded for harvest.ts and boxes.ts.",
  },
  {
    file: "src/core/pptx/tags.ts",
    what: "fallback",
    was: "?? 0",
    why: "Both fallbacks in `usedInDeck`'s sort key, `(a.slides[0] ?? 0) - (b.slides[0] ?? 0)` (line 410). An entry is only put in `found` by the same iteration that does `entry.slides.add(i + 1)`, so `slides` is never empty and neither fallback is ever taken. Measured 2026-09-14: zero disagreements over 500 randomised decks — five slides, four element ids, shapes at top level and inside groups, 1 to 7 stamps per deck, no empty baseline — and over the 21-package corpus. They stay because `noUncheckedIndexedAccess` is on, so `a.slides[0]` is `number | undefined` and the expression would not typecheck without them.",
  },
  {
    file: "src/core/pptx/tags.ts",
    what: "off-by-one",
    was: "0",
    why: 'The two `?? 0` fallbacks of the same sort key turned into `?? 1` (line 410): they change the value of a branch nothing reaches, for the reason in the `?? 0` entry above, and measured in the same two runs with zero disagreements. NOT the two index literals on that line — `a.slides[0] -> a.slides[1]` is killed by test/pptx-tags.test.ts "orders elements by where they first appear", and `b.slides[0] -> b.slides[1]` is killed as of 2026-09-14 by "orders an element by the FIRST slide it is on, not by a later one", which was written because 500 randomised decks disagreed 183 times. This entry is keyed on file and operator, so it would also label a future `0` survivor elsewhere in tags.ts; read the line before trusting it.',
  },
  {
    file: "src/core/pptx/layout.ts",
    what: "guard",
    was: "continue",
    why: "line 47, the `if (node.nodeType !== 1) continue;` in `placeholderOf`. Measured over every node kind @xmldom/xmldom can put inside an element — text, CDATA, PI, comment; it never builds an entity-reference node, a DTD entity arrives as text — and all four answer firstChild null and childNodes.length 0. `child()` therefore answers undefined for them and the `if (nvPr)` below skips on, which is what the `continue` did. 28 mixed-noise decks, 0 disagreements.  ALSO: line 129, the `if (node.nodeType !== 1) continue;` over the slide's own shape tree. A non-element handed to `placeholderOf` walks an empty `childNodes` and answers undefined, and `if (!ph || !wants(ph, want)) continue` one line below is the same skip. Measured with indentation, comments, CDATA, a processing instruction and an entity in the tree: 0 disagreements over 71 probe cases.  ALSO: line 147, the `if (node.nodeType !== 1) continue;` over the layout and master trees. Line 129's argument, in the second loop, measured in the same sweep with the noise placed in the layout's tree. 0 disagreements.  ALSO: line 143, the `if (!part) continue;` over `[layout, master]`. Equivalent at RUNTIME and not in the types. `spTreeOf` already answers undefined for a part the package does not hold and `if (!tree) continue` catches that, and `pkg.has(undefined)` is false on every package JSZip opens from a .pptx, so the extra iteration reads and answers nothing — six chain cases, 0 disagreements. The guard is what narrows `string | undefined` to `string`: removing it fails `npx tsc --noEmit` with TS2345 at 144,38. The suite cannot kill this mutant because vitest transpiles without typechecking; `npm run typecheck` does kill it. Strictly it is not equivalent on a package holding a zip entry literally named `undefined` — JSZip's `file(undefined)` then answers that entry — which no OPC producer writes and `relatedOfType` can never return.",
  },
  {
    file: "src/core/pptx/layout.ts",
    what: "fallback",
    was: "?? undefined",
    why: 'line 154, the `?? undefined` in `(ph.getAttribute("idx") ?? undefined) !== idx`. Exhaustive over the value domain. `getAttribute` answers a string or null; the comparison is reached only when `idx` is a string, and null and undefined are both unequal to every string, while a string is unaffected by the `??`. 13 slide/layout idx pairings, 0 disagreements.',
  },
  {
    file: "src/core/pptx/layout.ts",
    what: "fallback",
    was: "?? 0",
    why: 'line 173, the `?? 0` in `Number(sz?.getAttribute("cx") ?? 0)`. The `??` only ever fires when there is no `<p:sldSz>` at all, turning undefined into 0 instead of NaN. 0 fails `width <= 0`, NaN fails `!Number.isFinite(width)`, both take the same arm of the same `if` and return the same 4:3 default, and nothing reads `width` before it. 24 `<p:sldSz>` spellings, 0 disagreements.  ALSO: line 174, the `?? 0` in `Number(sz?.getAttribute("cy") ?? 0)`. Line 173\'s argument for the height, measured in the same 24-case sweep: 0 disagreements. Distinct from the off-by-one mutant on the same line — `?? 1` IS observable and now has a test.',
  },
  {
    file: "src/core/pptx/base64.ts",
    what: "boundary",
    was: "<",
    why: 'TWO separate lines in this file carry a `<` that cannot be widened observably, so one entry keyed on file and operator covers both; both measured 2026-09-14 by running the original loop and the mutated loop side by side over 3,019 byte lengths (0 to 299,960, every length 0-8, every k*CHUNK-2 .. k*CHUNK+2 for k in 1..4, and 3,000 uniform randoms below 300,000), with the decode loop fed whatever atob hands back for the text the encode loop produced. (1) `bytesFrom`\'s browser loop, `i < binary.length`: the extra turn at i === binary.length evaluates `out[i] = binary.charCodeAt(i)`, which is an out-of-range write of NaN onto a Uint8Array. Measured directly on the same day: such a write does not throw in strict mode (the module is an ES module, so strict is not optional) and stores nothing — `probe[3] = NaN` on a three-byte array leaves `probe[3]` undefined — and `"abc".charCodeAt(3)` is NaN, read from a string rather than a typed array so no bound is crossed there either. Zero disagreements in 3,019 comparisons of the returned bytes. (2) `base64From`\'s browser loop, `at < bytes.length`: widened, the loop runs one extra turn ONLY when `bytes.length` is an exact multiple of CHUNK (0xfffc), since `at` advances by CHUNK and the two bounds can only meet exactly. That turn appends `String.fromCharCode(...bytes.subarray(at, at))`, and `String.fromCharCode()` of an empty spread is "" — measured. Zero disagreements in 3,019 comparisons of the base64 text, five of those lengths being exact multiples of CHUNK (0, 65,532, 131,064, 196,596, 262,128), which is the only shape where the mutant differs at all. Both lines stay: they are the honest bounds of the arrays they walk, and killing either would mean asserting on a write that is discarded or a concatenation of the empty string.',
  },
  {
    file: "src/core/pptx/xml.ts",
    what: "boundary",
    was: "<",
    why: "`relationshipIdsIn`'s attribute loop bound, `i < (attrs?.length ?? 0)`, widened to `<=` so it runs one extra turn at `i === length`. Measured 2026-09-14 over 1,080 real XML parts — every part of template/library-16x9.pptx, template/library-4x3.pptx and template/validators.pptx, plus the committed catalogue markup — holding 469,288 elements and 571,356 attributes: `attributes.item(length)` answered null on all 469,288 elements and never anything else, so the `if (!attr) continue` on the line below absorbs the extra turn and nothing is read or added. Zero disagreements between original and mutant over all 1,080 documents. The same shape and the same DOM as the harvest.ts `relIdsIn` entry above; killing it would mean asserting on a DOM read that never happens.",
  },
  {
    file: "src/core/pptx/xml.ts",
    what: "fallback",
    was: "?? 0",
    why: "the same loop bound's `?? 0`, deleted so it reads `i < (attrs?.length)`. The fallback is reached only when `node.attributes` is nullish, and that happened 0 times in 469,288 elements (measured 2026-09-14, same corpus as the boundary entry). Unobservable even where it IS reached, which was measured rather than reasoned: against a synthetic node with `attributes` undefined, and again with it null, `i < undefined` is false and `i < 0` is false, so both loops run zero turns and both return the same empty set. So the mutation changes nothing on any input, reachable or not. The `?? 0` stays because it keeps the bound a number rather than resting on `i < undefined` being false.",
  },
  {
    file: "src/core/pptx/xml.ts",
    what: "off-by-one",
    was: "0",
    why: 'the `0` in that same `?? 0`, shifted to `?? 1`. Line 164 carries TWO `0` sites and only this one survives: `let i = 0` -> `let i = 1` skips the first attribute of every element and is killed by test/package-surface.test.ts\'s `expect([...relationshipIdsIn(bare)]).toEqual(["rId5"])` — measured 2026-09-14, it returns [] there, and disagrees with the original on 83 of the 1,080 real parts. `?? 1` disagrees on 0 of 1,080, because the fallback is only reached when `node.attributes` is nullish, which no part produces; and on a synthetic node with `attributes` undefined it still agrees, because the one turn it then runs reads `attrs?.item(0)` as undefined and the guard on line 166 continues past it. Keyed on file and operator, so this entry would also label a future `"0"` survivor at `let i = 0` — which is a kill today, so re-verdict rather than trust the label if that one ever appears.',
  },
  {
    file: "src/core/pptx/xml.ts",
    what: "guard",
    was: "continue",
    why: "`if (!attr) continue;` inside `relationshipIdsIn`, dropped. It can only fire when `attrs.item(i)` is falsy for an index the loop visits, which under the original `i < length` bound never happens: measured 2026-09-14, 0 falsy answers over 571,356 in-range attribute reads across 469,288 elements of the two library decks, validators.pptx and the committed catalogue, and @xmldom/xmldom's NamedNodeMap is densely indexed — its `item` IS `NodeList.prototype.item`, which answers null only outside 0..length-1. Zero disagreements over all 1,080 documents, and zero against a synthetic node whose `attributes` is undefined or null. The line is NOT deleted: it narrows `Attr | null | undefined` to `Attr` for the `attr.namespaceURI` and `attr.name` reads below it, so removing it takes `npm run typecheck` red, and it is also what makes the widened bound above harmless — the same reason the boxes.ts `continue` entry gives.",
  },
  {
    file: "src/core/pptx/theme.ts",
    what: "boundary",
    was: "<",
    why: "`coloursOf`'s `<p:clrMap>` attribute loop, `i < map.attributes.length`, widened to `<=`. The same shape as the `harvest.ts` and `boxes.ts` entries: one extra turn at `i === length`, and @xmldom/xmldom 0.9.12 answers null for `attributes.item(length)` — measured directly on 2026-09-14, as it does for `item(length + 5)` and `item(-1)` — so the `if (!attr) continue` on the next line absorbs it and no attribute is read. Zero disagreements over 712 theme x part pairs drawn from the three committed decks (every theme against every master, layout, slide and notes master in library-16x9, library-4x3 and validators) and over 20,000 synthetic `<p:clrMap>` elements carrying 0 to 8 attributes, names and values drawn from a vocabulary of the real ones plus empty strings, against three palettes. Killing it would mean asserting on a DOM read that never happens.",
  },
  {
    file: "src/core/pptx/theme.ts",
    what: "guard",
    was: "continue",
    why: '`if (!attr) continue` in that same loop, dropped. Under the loop\'s own bound `map.attributes.item(i)` is never null for 0 <= i < length, so the guard is never taken: zero disagreements over the same 712 real and 20,000 synthetic cases on 2026-09-14. It is NOT deleted, for the two reasons the `boxes.ts` guard entry gives — `item(i)` is typed `Attr | null`, so this line is what makes `attr.value` and `attr.name` honest without a non-null assertion, and it is the line that absorbs the boundary mutant above, so deleting it would make that mutant observable. CAUTION for the next reader: this entry is keyed on file and operator, so it would also label a future `continue` survivor in `relatedOfType`, which has three of them; all three are killed today by the "a relationship that does not lead to a theme" cases in test/colours.test.ts, so a `continue` survivor in theme.ts must be checked against its line number before this reasoning is applied to it.',
  },
  {
    file: "src/core/splice/colours.ts",
    what: "guard",
    was: "continue",
    why: "TWO lines in `pinSchemeColours`, both dropped guards and both unobservable, measured 2026-09-14. (1) `if (child.nodeType !== 1) continue` — a text node, a comment, a CDATA section and a processing instruction all answer `localName` null and `namespaceURI` null in @xmldom/xmldom 0.9.12, so a fallen-through node fails the `schemeClr` test and is handed to `walk`, where `Array.from(node.childNodes)` is a defined, empty list on every one of them; nothing throws, nothing is counted, nothing is replaced. A declared entity expands to a text node rather than an EntityReference, so those four and Element are the whole domain. (2) `const doc = el.ownerDocument; if (!doc) continue` — every Element in this DOM has an owner document, including a created-but-unattached one, and `walk` only ever receives Elements because `start` is `root.documentElement` when a Document is passed. Zero disagreements for either over 2,085 cases built from the committed catalogue (117 elements x 2 sizes plus every .xml part, against three palettes), run both through `pinColoursInXml` and as an Element handed straight to `pinSchemeColours`, and over 40,000 synthetic markup cases nesting every non-element node type three deep. Neither is deleted: the first is what makes the `as Element` cast below it honest, the second what lets `doc.createElementNS` be called without a non-null assertion.",
  },
  {
    file: "src/core/splice/shapes.ts",
    what: "boundary",
    was: "<",
    why: 'shapes.ts:57 today, `for (let i = 0; i < out.length; i++)` in everyElement. `<=` makes one extra pass in which `out[i]` is undefined; the `continue` on the next line skips it, i then exceeds out.length and the loop ends, so the array everyElement returns is byte-for-byte the same. The equivalence DEPENDS on that guard, which is itself type-killed and so cannot be removed. Measured 2026-09-14: 0 disagreements over both committed decks (234 elements), fingerprinting highestShapeId, renumber, relIdsIn, unionOf, unplaceholder, emptyBodyPlaceholders, slideShapes and five applyMove variants. NOTE: the ledger matches on file + operator + text, so this entry will also cover a future `<` survivor elsewhere in shapes.ts. The only other `<` in the file today is `shapes.length < 2` in groupable (line 289), and that one is KILLED by "refuses a single shape, which is already one thing to move".',
  },
  {
    file: "src/core/splice/shapes.ts",
    what: "fallback",
    was: "?? 0",
    why: 'shapes.ts:79 today, `Number(el.getAttribute("id") ?? 0)` in highestShapeId. getAttribute can only answer a string or null — measured against @xmldom/xmldom: a missing attribute gives null, an empty one gives "" — and Number(null) === 0 === Number(0), while a string makes the `??` unreachable. So the fallback cannot change the number for any DOM. Measured 2026-09-14: 0 disagreements over the same 234-element fingerprint. NOTE: matching is on file + operator + text, so this covers a future `?? 0` survivor anywhere in shapes.ts; there is no other `?? 0` in the file today. The file\'s other `??` fallbacks are `?? null` at line 341 (type-killed), `?? ""` at line 405, and `?? "body"` at line 398 (killed by "treats a placeholder that states no type as a body, which is what PowerPoint does").',
  },
  {
    file: "src/core/splice/shapes.ts",
    what: "boundary",
    was: ">",
    why: "shapes.ts:80 today, `if (Number.isSafeInteger(n) && n > max) max = n`. `>=` differs only when n === max, where the assignment writes the value already in max; NaN never reaches the comparison because Number.isSafeInteger rejects it first. Equivalent for every sequence of ids. Measured 2026-09-14: 0 disagreements over the same 234-element fingerprint. NOTE: matching is on file + operator + text, so this covers a future `>` survivor anywhere in shapes.ts; line 80 is the only `>` in the file today.",
  },
  {
    file: "src/core/splice/carry.ts",
    what: "fallback",
    was: "?? part",
    why: 'Line 105, `const name = to.pop() ?? part;` in `targetFrom`. `to` is `part.split("/")`, and `String.prototype.split` never answers a zero-length array — the empty string gives `[""]` — so `to.pop()` is always a string and the `?? part` arm is dead. Measured 2026-09-14: 240,100 owner/part pairs over a 490-path domain built from empty segments, leading and trailing slashes, dot-only and slash-only names gave 0 disagreements between the two versions, and a 200,000-string randomised sweep never saw `split("/")` return an empty array. This entry\'s text `?? part` is specific to this site; the file\'s other fallback, `rel.getAttribute("TargetMode") ?? ""` at line 247, has different text and is not covered by it.',
  },
  {
    file: "src/core/splice/carry.ts",
    what: "boundary",
    was: "<",
    why: 'Both `<` sites on line 107, `while (shared < from.length && shared < to.length && from[shared] === to[shared])` in `targetFrom`. Relaxing either one on its own is inert, because the OTHER conjunct still bounds the index it would have loosened: one side is then `undefined` and the other a string (both arrays come from `split("/")`, so they are dense arrays of strings), the equality is false, and the loop stops in the same place. Measured 2026-09-14: 0 disagreements over the same 240,100 owner/part pairs, for each mutation separately. CAUTION — this entry matches on file and operator and text, so it will also label a future `<` survivor anywhere else in carry.ts. The file\'s one other `<` is line 147, `i < bytes.length` in `fingerprint`, and that one is NOT equivalent: it reads one past the end and changes every carried picture\'s name. It is killed today by "names a picture by its bytes and its length, and that name does not drift between builds" in test/splice-carry.test.ts, so it should not reappear as a survivor; if it ever does, the ledger is wrong about it and the test is what to look at.',
  },
  {
    file: "src/core/splice/carry.ts",
    what: "boundary",
    was: ">",
    why: 'Line 194, `const extension = dot > part.lastIndexOf("/") ? part.slice(dot) : "";` in `mediaName`. `>` and `>=` differ only where the last dot and the last slash share an index, which two different characters can only do by both being absent (-1 and -1). `mediaName` is module-private and reached only through `shareable = typeof bytes !== "string" && isMedia(part)`, and `isMedia` demands the prefix `ppt/media/`, which contains a slash — so that input cannot arrive. Measured 2026-09-14: over 500,000 random strings every disagreement had both indices at -1, and 0 of the same 500,000 disagreed once prefixed with `ppt/media/`. CAUTION — this entry will also label a future `>` survivor elsewhere in carry.ts. The file\'s other two `>` sites are lines 128 and 129 in `freeName`, which is EXPORTED and takes any part name, including one with neither a dot nor a slash (`resolveFrom` produces one when a target climbs past the package root). Those are NOT equivalent and are killed today by "numbers a name with no directory and no extension at its end, not in its middle" in test/splice-carry.test.ts.',
  },
  {
    file: "src/core/splice/splice.ts",
    what: "guard",
    was: "continue",
    why: 'TWO of the file\'s three `continue` guards, and read the line number before trusting this entry, because the third is a KILL. (1) Line 248, `placeholderIn`\'s `if (node.nodeType !== 1) continue`. Dropped, the node is handed to `child(node as Element, P_NS, "nvPr")`, which walks `firstChild`/`nextSibling` — null on every non-element in @xmldom/xmldom — so it answers undefined and `if (nvPr && child(nvPr, P_NS, "ph")) return true` on the next line skips on, which is what the `continue` did. (2) Line 412, the `<p:grpSp>` test in the pass that stamps the shapes inside a group. Dropped, every added top-level shape descends into `slideShapes`, and nothing inside a non-group is taggable: `taggable` wants a `<p:nvPr>` exactly three levels below the top-level shape, and only a group has one there — a `<p:sp>` keeps its `<p:nvPr>` two levels down, a `<p:pic>`, `<p:cxnSp>` and `<p:graphicFrame>` likewise, and an `<mc:AlternateContent>` keeps a whole shape below its `Choice` and `Fallback`, further still. Measured 2026-09-14 by differential snapshot, not by reading the code: each mutant compiled as its own module and run beside the original over 680 splices — all 117 elements of BOTH committed libraries onto a three-slide fixture (grouped and loose, onto and as a new slide), plus 192 crafted cases pairing eight hand-written markups (a plain shape, a picture, a connector, a table in a graphicFrame, an `<mc:AlternateContent>` modern chart, a real group carrying a comment and a CDATA section, a shape whose children are interleaved with text, comments and processing instructions, and a shape holding a nested group) with six destination decks whose title placeholder carries a comment, a CDATA section, a processing instruction or only whitespace — comparing the whole output package, part by part, by sha256, and the SpliceReport beside it, with `Math.random` seeded so the clone\'s creation id is the same draw on both sides. 0 disagreements for line 248 in all 680. 0 for line 412 in 656 of them. THE 24 THAT DISAGREE, and the caveat this entry rests on: line 412\'s mutant differs on exactly one crafted markup, a `<p:sp>` carrying a `<p:grpSp>` among its children, where the nested group IS taggable and gets stamped. CT_Shape has no group child, so no package PowerPoint writes — and so no element the harvest can produce — contains one; the same shape of caveat the layout.ts and pkg.ts entries carry. NOT equivalent, and NOT covered by this entry: line 227, the `if (node.nodeType !== 1) continue` in `blank`\'s paragraph loop, where falling through REMOVES the node — 207 disagreements in the same 680-splice sweep — killed as of 2026-09-14 by test/splice.test.ts "empties the paragraph of its runs and leaves everything that is not an element alone". Neither equivalent line is deleted: each is what makes the `as Element` cast or the group assumption on the line below it honest.',
  },
  {
    file: "src/core/splice/splice.ts",
    what: "boolean",
    was: "||",
    why: 'Line 412\'s `if (shape.namespaceURI !== P_NS || shape.localName !== "grpSp") continue`, in the pass that stamps the shapes inside a group. With `&&` the guard only fires for a shape that is neither in the presentation namespace nor named `grpSp`, so the loop additionally descends into P_NS shapes that are not groups — and finds nothing taggable there, for the reason the guard/"continue" entry above sets out at length: `taggable` wants a `<p:nvPr>` exactly three levels below the top-level shape and only a group has one there. Same 2026-09-14 differential snapshot, same corpus: 0 disagreements over 656 splices, including all 234 elements of both committed libraries, of which 23 already ARE a single `<p:grpSp>` the owner drew. The same 24 crafted cases disagree, and only those — a `<p:sp>` carrying a nested `<p:grpSp>`, which is schema-invalid PresentationML and which the harvest cannot produce. Keyed on file and operator, so this entry would also label a future `||` survivor elsewhere in splice.ts; the file\'s other two are killed today — line 231\'s `pPr || endParaRPr`, by "keeps a placeholder but empties every paragraph in it", which asks for both by name, and line 271\'s slide-range check, by "refuses a slide number the deck does not have, by name".',
  },
  {
    file: "src/core/splice/remove.ts",
    what: "boundary",
    was: "<",
    why: "Line 99, `slidesHolding`'s `for (let i = 0; i < paths.length; i++)`. Widened to `<=` the loop runs one extra turn at `i === paths.length`, where `paths[i] as string` is undefined; `readShapeTags` answers [] for it at its own first line, `if (!pkg.has(slidePath)) return out`, so nothing is read and nothing is pushed. Both readings taken from the library on 2026-09-14 rather than reasoned about: `pkg.has(undefined)` is false (JSZip's `file(undefined)` is null) and `readShapeTags(pkg, undefined)` is []. Measured the same day against the original over 849 recorded observations with 0 disagreements — 8 decks (a plain three-slide deck, a deck after an insert, a deck carrying another add-in's shape tag, a two-element slide, and noise-laden copies of each) read at slide indices -2 through 4 and at 1.5, plus 400 randomised distributions of text, comment, CDATA and processing-instruction nodes through the shape tree, recording `removeElement`'s removed/left/deckSlides/slidePath, the whole rebuilt slide XML, and `slidesHolding` for three element ids. (`p14:creationId` is randomised per run and was normalised out.) Strictly not equivalent on a package holding a zip entry literally named `undefined`, which no OPC producer writes — the same caveat as the theme.ts guard entry. This entry is keyed on file and operator, so it would also label a future `<` survivor elsewhere in remove.ts; the file's other `<` site, `request.slide < 0` on line 148, is a kill today, so check the line before trusting the label.",
  },
  {
    file: "src/core/splice/remove.ts",
    what: "guard",
    was: "continue",
    why: 'BOTH nodeType guards this file has: line 127 in `idOf` and line 178 in `strip`, each `if (node.nodeType !== 1) continue;`. Established by measurement on 2026-09-14, in the same run as the boundary entry above: 849 observations per version, 0 disagreements for each mutant, over 8 decks and 400 randomised placements of text, comment, CDATA and processing-instruction nodes at random positions in the shape tree, inside plain shapes and inside the group the insert writes. The library says why, read directly the same day: for all four non-element kinds @xmldom/xmldom answers childNodes.length 0, firstChild null, localName null and namespaceURI null. So in `idOf` a fallen-through node gives `child(… , P_NS, "cNvPr")` undefined and the `if (id)` below skips it; in `strip` it gives `idOf` undefined and then fails `shape.namespaceURI !== P_NS` on the next line, which is the same skip — `removeChild` and `slideShapes` are never reached for it. Neither is deleted: each is what makes the `as Element` cast on the line below it honest, the same reason the tags.ts and boxes.ts guard entries give. Keyed on file and operator, so this entry would also label a future `continue` survivor elsewhere in remove.ts; the third one, `if (shape.namespaceURI !== P_NS || shape.localName !== "grpSp") continue;` on line 186, is NOT covered by it and is a kill today.',
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
 * about thirteen hours of first tier, at 62.2 s for each of the engine's 757
 * mutations. With it, the whole set of 1073 came in at 6 h 20 m — see the
 * measured total at the end of this comment, which replaced a much rosier
 * projection.
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
 * WHAT A WHOLE SWEEP ACTUALLY COSTS, measured rather than projected. On
 * 2026-09-13, on an idle machine, with the corrected operator and with kill
 * confirmation in place: 1020 mutations in 6 h 20 m, or about 22 s each. It was
 * cut short by a restart 53 short of the full 1073, so that is the figure for
 * 1020 and not for all of them.
 *
 * The paragraph above used to say "about an hour", from a projection of
 * 66 minutes of first tier over 998 mutations. It was wrong by roughly six
 * times, and it was wrong in the way this file keeps warning about: a number
 * derived once and then left standing as though it were a property of the tool.
 * Two things it left out. Of the 22 s, roughly 8 s per mutant is the survivor
 * re-check amortised — 144 survivors at 57.6 s each is 2 h 18 m of the 6 h 20 m,
 * and that is a DERIVED split, not a second measurement. And the projection
 * only ever covered `src/core`; `src/host` and `src/pane` have no rows in this
 * map at all, so their 316 mutations run against reaching sets that include
 * `test/pane-wiring.test.ts` at 15.2 s and `test/office-host.test.ts` at 15.7 s.
 * 227 of those 316 pay a ~15 s tier for what their own unit test answers in
 * ~1 s. Rows for them are the obvious next saving and are not written yet.
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
  if (/** @type {NodeJS.ErrnoException} */ (error)?.code === "ETIMEDOUT") return "hung";
  try {
    return failedNames(JSON.parse(readFileSync(out, "utf8"))).length > 0 ? "killed" : "inconclusive";
  } catch {
    return "inconclusive";
  }
}

/**
 * The per-file results of the report on disk, or nothing if it cannot be read.
 *
 * @param {string} out where the JSON report goes
 * @returns {{ name?: string, status?: string, assertionResults?: { status?: string }[] }[]}
 */
function readReport(out) {
  try {
    return /** @type {{ testResults?: unknown }} */ (JSON.parse(readFileSync(out, "utf8"))).testResults ?? [];
  } catch {
    return [];
  }
}

/**
 * Which TEST FILES a failed run blamed, relative to the tree being swept.
 *
 * `verdictOf` reads the report to learn WHETHER a test failed and throws away
 * WHICH, and that is the whole of what a kill is checked against. Naming the
 * file is what lets a kill be confirmed cheaply.
 *
 * Vitest writes an absolute path, and the sweep runs in a throwaway copy, so
 * the path is relativised — a run that is handed the copy's own absolute path
 * works, but the report and the console then disagree about what is being
 * talked about.
 *
 * @param {string} out where the JSON report goes
 * @returns {string[]} distinct test files, in the order the report lists them
 */
export function failedFilesOf(out) {
  const files = readReport(out);
  const blamed = files
    .filter((one) => one.status === "failed" || (one.assertionResults ?? []).some((test) => test.status === "failed"))
    .map((one) => relative(process.cwd(), one.name ?? ""))
    .filter(Boolean);
  return [...new Set(blamed)];
}

/**
 * Whether a kill survives being asked a second time, against the file that
 * reported the failure and nothing else.
 *
 * WHY A KILL NEEDS ASKING TWICE. Every SURVIVOR is re-run against the whole
 * suite before it is reported, and until now every KILL was believed on one
 * run. That asymmetry is backwards. A survivor wrongly reported costs a reader
 * an afternoon and says so out loud; a kill wrongly reported is a hole in the
 * suite that the report is SILENT about, which is the failure this whole script
 * exists to prevent.
 *
 * It is not hypothetical. On 2026-09-13 `src/core/catalogue/boxes.ts:148` was
 * reported killed by one sweep and alive by the next, on a suite that had only
 * grown and a line nothing about which had changed. The mutant is provably
 * behaviour-identical, so the kill was false: a test that failed for its own
 * reasons in a run sharing four cores with four other agents, and `verdictOf`
 * counts any named failure as a kill.
 *
 * The check is cheap because it is narrow. `--bail=1` means a killing run stops
 * at the first failure, and the confirmation re-runs ONE test file rather than
 * the tier. A real kill fails again; a flake does not.
 *
 * Anything short of a second failure resolves to "survived", including an
 * unclear answer — not because a survivor is the likelier truth, but because
 * "survived" costs one whole-suite re-check, which is the authoritative answer,
 * and "killed" costs silence. A test that only fails alongside others also
 * lands here and is then correctly killed by that re-check.
 *
 * @param {string[]} blamed the test files the first run reported failing
 * @param {(files: string[]) => "survived"|"killed"|"inconclusive"} rerun
 * @returns {"survived"|"killed"}
 */
export function confirmedKill(blamed, rerun) {
  // No file named is not evidence of a flake: a report can say a run failed
  // without attributing it. Nothing to re-run means nothing to re-check.
  if (!blamed.length) return "killed";
  return rerun(blamed) === "killed" ? "killed" : "survived";
}

/**
 * Whether the mutated source still typechecks.
 *
 * WHY THIS EXISTS. This file used to promise the opposite, in as many words: "a
 * mutation that will not compile is simply killed by the run like any other...
 * Nothing here can produce" a mutant reported ALIVE that never ran. That was
 * false. Vitest transpiles with esbuild, which STRIPS types rather than
 * checking them, so a mutation that breaks only the types runs perfectly and is
 * reported as a survivor — a hole in the suite that is not a hole at all,
 * because `npm run typecheck` is a CI step and rejects it.
 *
 * Proven 2026-09-14 on `src/core/splice/layout.ts:143`: neutering that guard
 * gives `tsc --noEmit` error TS2345 at 144,38 while the suite passes 54 of 54.
 * Measured over the whole report the same day, three of the 57 survivors then
 * open in `src/core/splice` were this and nothing else.
 *
 * A mutant `tsc` refuses is not a survivor and not an equivalent. It is caught,
 * by the other half of the gate.
 *
 * @param {string} file
 * @param {string} text the original
 * @param {string} mutated
 * @returns {boolean}
 */
function typechecks(file, text, mutated) {
  writeFileSync(file, mutated);
  try {
    execFileSync("npx", ["tsc", "--noEmit"], { stdio: "ignore", timeout: RUN_LIMIT_MS, killSignal: "SIGKILL" });
    return true;
  } catch {
    return false;
  } finally {
    writeFileSync(file, text);
  }
}

/**
 * The pids of vitest processes this sweep started and no longer owns.
 *
 * Separated from the killing so it can be held by a test: the dangerous part is
 * the FILTER, not the signal.
 *
 * @param {string} listing the output of `ps -eo pid,args`
 * @param {string} workspace the sweep's throwaway copy
 * @param {number} self this process's own pid
 * @returns {number[]}
 */
export function strayPids(listing, workspace, self) {
  return (
    listing
      .split("\n")
      .map((line) => /^\s*(\d+)\s+(.*)$/.exec(line))
      .filter((m) => m !== null)
      .filter((m) => Number(m[1]) !== self && (m[2] ?? "").includes(workspace))
      // A `ps` listing contains the command that ASKED for it, so a pattern
      // matching the workspace path matches the asker too. That is the mistake
      // `pkill -f` makes, and it cost this session two killed wrappers and an
      // orphaned sweep: the shell's own command line held the pattern. Matching
      // on the vitest entry point rather than on the path alone is what keeps
      // this from killing the hand that runs it.
      .filter((m) => (m[2] ?? "").includes("vitest"))
      .map((m) => Number(m[1]))
  );
}

/**
 * Kill the vitest workers a timed-out run left behind.
 *
 * WHY THIS EXISTS, measured 2026-09-14. `execFileSync`'s timeout signals the
 * vitest process it started. Vitest runs the tests in FORKS, and those are not
 * signalled: they are re-parented to init and keep running. When the mutation
 * is one that removes a loop's only exit, they keep running that loop — at full
 * speed, forever.
 *
 * Two mutants of this sweep did exactly that, and left FIFTEEN workers spinning
 * for two hours on a four-core machine: a load average of eighteen. Everything
 * measured beside it was wrong, and one mutant was reported KILLED that is not,
 * because a machine under that load fails timing-sensitive tests repeatedly —
 * which `confirmedKill` confirms rather than catches, since it re-runs the same
 * file.
 *
 * So the leak is not untidiness. It is how a sweep silently poisons its own
 * later answers, and the survivor list is what this script exists to produce.
 */
function killStrays() {
  try {
    const listing = execFileSync("ps", ["-eo", "pid,args"], { encoding: "utf8" });
    for (const pid of strayPids(listing, process.cwd(), process.pid)) {
      try {
        process.kill(pid, "SIGKILL");
      } catch {
        // Already gone between the listing and the signal, which is the common case.
      }
    }
  } catch {
    // No `ps`, or it answered nothing. A sweep that cannot tidy up still runs.
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
    let verdict = verdictOf(tests, out);
    // A hang and an unclear answer are both asked twice, for opposite reasons.
    // An unclear answer is usually noise and usually goes away. A hang usually
    // does NOT: a mutation that removes a loop's only exit hangs every time,
    // and the second ask is what separates that from a wedged machine.
    if (verdict === "inconclusive" || verdict === "hung") {
      if (verdict === "hung") killStrays();
      const again = verdictOf(tests, out);
      verdict = verdict === "hung" && again === "hung" ? "hung" : again;
      if (verdict === "hung") killStrays();
    }
    if (verdict !== "killed") return verdict;
    // The mutation is still in place — the restore is in the `finally` below —
    // so the confirmation asks the same question of the same code.
    return confirmedKill(failedFilesOf(out), (files) => verdictOf(files, out));
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
  /** Mutants that stop the code terminating. The suite notices; it just never finishes. */
  const hung = [];
  /** Mutants `tsc` refuses. CI catches them; the suite cannot, so they are not survivors. */
  const typeKilled = [];
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
      if (verdict === "hung") {
        hung.push(where);
        console.log(`mutants: DID NOT TERMINATE  ${where}`);
        continue;
      }
      if (verdict === "killed") {
        if (done % 10 === 0) console.log(`mutants: ${done}/${total}, ${survivors.length} surviving so far`);
        continue;
      }
      // A survivor of the fast tier is not a survivor yet. Re-run it against
      // everything before it goes in the report.
      if (tryMutation(file, text, mutated, null, out) === "survived") {
        // The suite is not the only thing that can reject a mutation, and until
        // today this script said otherwise. `tsc` is asked LAST, only about a
        // mutant that has already survived everything else, so it costs a few
        // seconds on the handful rather than on all of them.
        if (!typechecks(file, text, mutated)) {
          typeKilled.push(where);
          console.log(`mutants: TYPE-KILLED  ${where}`);
          continue;
        }
        survivors.push(where);
        // Written and printed as it is found, not only in the summary. A sweep
        // of the whole set is over an hour, and a run that is interrupted
        // should still have told the reader everything it knew at the time.
        console.log(`mutants: SURVIVED  ${where}`);
        appendFileSync(report, `${where}\n`);
      }
    }
  }
  if (hung.length) {
    console.log(`\nmutants: ${hung.length} DID NOT TERMINATE — the suite noticed by never finishing:`);
    for (const one of hung) console.log(`  ${one}`);
    console.log("Each removes the only way out of a loop. NOT survivors, and not equivalents either.");
  }
  if (typeKilled.length) {
    console.log(`\nmutants: ${typeKilled.length} TYPE-KILLED — the suite passed, \`tsc\` did not:`);
    for (const one of typeKilled) console.log(`  ${one}`);
    console.log("Caught by `npm run typecheck` in CI. Not a hole in the suite, and nothing to write.");
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
