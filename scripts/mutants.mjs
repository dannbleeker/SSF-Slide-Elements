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
import { appendFileSync, existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { isMain } from "./is-main.mjs";

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
 * Run some of the suite, or all of it. True when it passed — which, for a
 * mutant, means it SURVIVED.
 *
 * @param {string[]|null} files null for the whole suite
 * @returns {boolean}
 */
function suitePasses(files) {
  const args = [join("node_modules", "vitest", "vitest.mjs"), "run", "--reporter=dot", "--bail=1"];
  if (files) args.push(...files);
  try {
    execFileSync(process.execPath, args, { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

/**
 * Put the mutation in place, run the tests, and ALWAYS put the original back.
 *
 * The restore is in a `finally` because a half-written file left behind by an
 * interrupted sweep is the one way this script can do damage, and a `git
 * checkout` to recover it would take the reader's own work with it.
 *
 * @param {string} file
 * @param {string} text the original
 * @param {string} mutated
 * @param {string[]|null} tests null for the whole suite
 * @returns {boolean} true when the suite passed, which for a mutant is survival
 */
function tryMutation(file, text, mutated, tests) {
  writeFileSync(file, mutated);
  try {
    return suitePasses(tests);
  } finally {
    writeFileSync(file, text);
  }
}

function main() {
  const argv = process.argv.slice(2);
  const onlyAt = argv.indexOf("--only");
  const only = onlyAt === -1 ? "" : (argv[onlyAt + 1] ?? "");
  const listing = argv.includes("--list");
  const files = TARGETS.filter((f) => f.includes(only));

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
  let done = 0;
  for (const { file, text, mutations, tests } of plan) {
    if (!tests.length) console.log(`\nmutants: NOTHING IMPORTS ${file} — every mutation of it will survive`);
    for (const m of mutations) {
      const mutated = text.slice(0, m.at) + m.now + text.slice(endOf(m));
      const lived = tryMutation(file, text, mutated, tests);
      done += 1;
      const where = `${file}:${lineOf(text, m.at)}  ${m.what}  ${JSON.stringify(m.was)} -> ${JSON.stringify(m.now)}`;
      if (!lived) {
        if (done % 10 === 0) console.log(`mutants: ${done}/${total}, ${survivors.length} surviving so far`);
        continue;
      }
      // A survivor of the fast tier is not a survivor yet. Re-run it against
      // everything before it goes in the report.
      if (tryMutation(file, text, mutated, null)) {
        survivors.push(where);
        // Written and printed as it is found, not only in the summary. A sweep
        // of the whole set is over an hour, and a run that is interrupted
        // should still have told the reader everything it knew at the time.
        console.log(`mutants: SURVIVED  ${where}`);
        appendFileSync(report, `${where}\n`);
      }
    }
  }
  if (!survivors.length) {
    console.log("mutants: no survivors — every mutation the suite could see, it saw");
    return;
  }
  console.log(`mutants: ${survivors.length} survived the WHOLE suite —`);
  for (const one of survivors) console.log(`  ${one}`);
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
