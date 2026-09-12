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
 * Measured 2026-09-12: the whole suite is 166 s, and `test/splice.test.ts`
 * alone is 82 of them, because it harvests both 1.5 MB library decks and sweeps
 * 117 elements. The files a mutation of `src/host` can possibly affect are all
 * in the fast half. So each mutant runs against the fast tier, and then every
 * SURVIVOR is re-run against the whole suite before it is reported — because a
 * survivor of a partial run is not a survivor, it is an untested guess.
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
];

/**
 * Mutations that CANNOT be killed, with the proof for each.
 *
 * A record, NOT a suppression list. Every one of these still appears in the
 * report; they are only labelled, so the next reader does not spend an afternoon
 * re-deriving what this afternoon already settled. `docs/SIBLING.md` keeps a
 * borrowed finding dated for the same reason.
 *
 * All eleven were established on 2026-09-12, by making the mutation and looking
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
    why: "the line routes an id with a '#' past position 0 against one with none; compared original against mutant over 1,003,578 id pairs with zero disagreements. The read-back that the borrowed setSelectedSlides measurement rests on is line 97 and the empty-anchor guard, and both are pinned.",
  },
  {
    file: "src/host/jump.ts",
    what: "off-by-one",
    was: "0",
    why: "same line, same proof. A mutation of it in the other direction IS killed by test/jump.test.ts, so the routing is held; these two mutants simply cannot be observed.",
  },
  {
    file: "src/host/jump.ts",
    what: "off-by-one",
    was: "1",
    why: "the same line's `indexOf(\"#\") > 0` shifted to `> 1`: an id whose only '#' sits at position 1 is not a shape this add-in writes, so no input distinguishes them. Covered by the same 1,003,578-pair comparison, which is exhaustive over the id shapes the harvest produces.",
  },
  {
    file: "src/host/probe.ts",
    what: "boundary",
    was: ">",
    why: "line 85 is only reached after the line above returned for `landed === o.expected`, so `>` and `>=` agree on every remaining input, NaN and -0 included. The branch itself is live and covered.",
  },
  {
    file: "src/host/probe.ts",
    what: "fallback",
    was: "?? []",
    why: "the sweep's mask blanks the outer template literal but not the one nested in it, so the mutation lands on the INNER `?? []`, which is only evaluated once the outer guard has proved the array non-empty. It cannot be deleted either: tsc reports TS2532 without it. Mutating the OUTER one is killed.",
  },
  {
    file: "src/pane/search.ts",
    what: "guard",
    was: "return true",
    why: "an early-out that returns exactly what the code below it would return. It runs on every keystroke, so it is a speed shortcut rather than the dead defensive branch CLAUDE.md deletes on sight; removing it would cost time in the browser and change no answer.",
  },
  {
    file: "src/pane/search.ts",
    what: "guard",
    was: "return 0",
    why: 'the same shape inside `distance`: an early `return 0` for two equal strings, which the loops below reach anyway and answer 0 for. This function already lost one such shortcut deliberately — its docstring records deleting a guard for "doing nothing the loops below do not already do" — so whether this one goes too is a question about browser time, not correctness.',
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
];

/**
 * A survivor annotated with the reason it cannot be killed, when there is one.
 *
 * @param {string} where a survivor line as the report prints it
 * @returns {{ known: boolean, why: string }}
 */
export function judgeSurvivor(where) {
  const match = EQUIVALENT.find(
    (one) =>
      where.startsWith(`${one.file}:`) && where.includes(`  ${one.what}  `) && where.includes(JSON.stringify(one.was)),
  );
  return match ? { known: true, why: match.why } : { known: false, why: "" };
}

/**
 * Which recorded entries matched nothing this run, and so need re-verdicting.
 *
 * @param {string[]} survivors
 * @returns {string[]}
 */
export function staleEquivalents(survivors) {
  return EQUIVALENT.filter(
    (one) =>
      !survivors.some(
        (where) =>
          where.startsWith(`${one.file}:`) &&
          where.includes(`  ${one.what}  `) &&
          where.includes(JSON.stringify(one.was)),
      ),
  ).map((one) => `${one.file}  ${one.what}  ${JSON.stringify(one.was)}`);
}

/**
 * Which test files can reach a source file, by following imports.
 *
 * The first version of this ran every mutant against a "fast tier" — the whole
 * suite minus the nine files that harvest a real library deck. Measured
 * 2026-09-12: that tier is 22 seconds, and 349 mutations of it is over two
 * hours. Running only the tests that can actually SEE the mutated file is 1 to
 * 3 seconds, and it is also the more honest question: a mutation of
 * `src/host/insert.ts` that `test/splice.test.ts` fails to notice tells nobody
 * anything.
 *
 * The safety is not in this map. It is in the whole suite being re-run against
 * every survivor before it is reported, so an import this misses can cost a
 * wasted minute and cannot produce a false survivor.
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

  // A comparison boundary. `=>`, `<=`/`>=` as part of `<<=`, and the arrow of a
  // generic are all left alone by matching the operator with what is around it.
  for (const m of mask.matchAll(/(^|[^=<>!])(<=|>=|<|>)(?!=)/g)) {
    const at = (m.index ?? 0) + (m[1] ?? "").length;
    const op = m[2] ?? "";
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
 * @param {string[]|null} files null for the whole suite
 * @param {string} out where the JSON report goes
 * @returns {"survived"|"killed"|"inconclusive"}
 */
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
    execFileSync(process.execPath, args, { stdio: "ignore" });
    return "survived";
  } catch {
    try {
      return failedNames(JSON.parse(readFileSync(out, "utf8"))).length > 0 ? "killed" : "inconclusive";
    } catch {
      return "inconclusive";
    }
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
    return { file, text, mutations: mutationsOf(text), tests: testsReaching(file) };
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
    console.log(`  ${one}${judged.known ? "   [known equivalent, 2026-09-12]" : ""}`);
    if (judged.known) console.log(`      ${judged.why}`);
  }
  const stale = staleEquivalents(survivors);
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
