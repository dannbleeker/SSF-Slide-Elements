#!/usr/bin/env node
/**
 * Hold a floor under the number of tests.
 *
 * Moving tests between files is where suites lose cases silently: the run stays
 * green because the deleted tests are simply not there to fail. A sibling
 * project lost 43 that way and found out much later. Comparing the total
 * against a recorded number costs one file and catches it in the diff.
 *
 * The floor rises on its own when the suite grows, because a floor that only
 * moves by hand drifts behind and stops catching a partial deletion. A
 * deliberate DROP is re-recorded with `--update`, so it lands in a commit where
 * a reviewer sees it rather than being absorbed silently.
 *
 * **Two numbers, not one, and the second is why.** This used to count only
 * tests that RAN — `numTotalTests` minus the pending ones — because counting
 * pending tests is what let 23 of them be switched off (`it(` to `it.skip(`)
 * with every gate green: the whole merge-plan decision engine, silent, floor
 * untouched.
 *
 * That was right about the danger and wrong about the measure. A test skipped
 * because the MACHINE cannot run it is not a test somebody switched off, and
 * counting the two the same way made the floor platform-dependent: this repo's
 * `is-main.test.ts` needs a symlink, Windows refuses that without elevation, so
 * the same commit counts one lower on Windows than on Linux. (The numbers this
 * paragraph used to quote, 1475 and 1476, were a SIBLING's; this repo's floor
 * has never been near them. It is in `test/fixtures/test-count.json`.) Either
 * number committed breaks the other machine — and the first way out taken was worse
 * than the problem, because it left the whole suite failing locally, which
 * teaches everyone to scroll past a red run.
 *
 * So the floor now counts tests that EXIST, which is the same number on every
 * machine, and a second recorded number caps how many may be skipped. Switching
 * off 23 tests leaves the total alone and blows the cap; deleting 23 drops the
 * total. Both still caught, and neither number depends on who ran it.
 */
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { isMain } from "./is-main.mjs";

const RECORD = "test/fixtures/test-count.json";

/**
 * What to do about a run, decided from the numbers alone.
 *
 * Separated from running vitest so it can be tested without running the suite
 * inside itself. Every branch below has cost a sibling repository something, and
 * none of them was provable while this was a column of `if`s wrapped around an
 * `execFileSync`.
 */
export function verdict({ defined, skipped, record, update = false }) {
  const min = Number.isInteger(record?.min) ? record.min : 0;
  // Absent in a record written before this change. Zero rather than Infinity: a
  // repository that has never recorded a cap has never agreed to a skip.
  const maxSkipped = Number.isInteger(record?.maxSkipped) ? record.maxSkipped : 0;

  if (!Number.isInteger(defined) || defined <= 0) {
    return { ok: false, message: `the run reported ${defined} tests, which is not a count. Refusing to compare.` };
  }
  if (!Number.isInteger(skipped) || skipped < 0) {
    return { ok: false, message: `the run reported ${skipped} skipped, which is not a count. Refusing to compare.` };
  }

  if (defined < min && !update) {
    return {
      ok: false,
      message:
        `the suite DEFINES ${defined} tests, down from ${min}.\n` +
        `If you deliberately removed tests, re-record it with:  node scripts/test-count.mjs --update`,
    };
  }
  if (skipped > maxSkipped && !update) {
    return {
      ok: false,
      message:
        `${skipped} tests are skipped and the record allows ${maxSkipped}.\n` +
        `A test skipped because this machine cannot run it is fine, and is recorded with:  node scripts/test-count.mjs --update\n` +
        `A test skipped to turn a red run green is the thing this gate exists for.`,
    };
  }

  // The floor rises on its own; the skip cap does NOT. A suite growing is
  // ordinary. A new skip is a decision somebody should be seen making.
  const raise = defined > min;
  return {
    ok: true,
    write:
      update || raise
        ? { min: update ? defined : Math.max(defined, min), maxSkipped: update ? skipped : maxSkipped }
        : null,
    message: raise
      ? `floor raised to ${defined}. Commit ${RECORD}.`
      : update
        ? `re-recorded: ${defined} tests, ${skipped} skipped, deliberately.`
        : `${defined} tests defined, floor ${min}.${skipped > 0 ? ` ${skipped} skipped, cap ${maxSkipped}.` : ""}`,
  };
}

/**
 * A scratch test file: `test/zz…`, gitignored, for probing something before it
 * becomes a case.
 *
 * Vitest runs one like any other file, so a scratch file open on the machine
 * that RECORDS the floor raises it by however many cases it holds — and the
 * next CI run fails on a number no commit contains. That is not a story: 989
 * was recorded here against a 988 CI, and the difference was one scratch file
 * I had left in the directory.
 *
 * The floor is a number about the repository, so it counts the repository's
 * files and nothing else.
 */
const SCRATCH = /(^|\/)zz[^/]*\.test\.[cm]?ts$/;

/** Whether a result file is scratch rather than committed. */
const isScratch = (name) => SCRATCH.test(String(name ?? "").replace(/\\/g, "/"));

/**
 * Every test result the repository owns, scratch left out.
 *
 * @param {unknown} report
 * @returns {{ status?: string, fullName?: string, title?: string }[]}
 */
function results(report) {
  const files = /** @type {{ name?: string, assertionResults?: { status?: string }[] }[]} */ (
    /** @type {{ testResults?: unknown }} */ (report ?? {}).testResults ?? []
  );
  return files.filter((file) => !isScratch(file.name)).flatMap((file) => file.assertionResults ?? []);
}

/**
 * How many tests the committed suite defines, and how many of those are
 * skipped.
 *
 * Counted from the per-file results rather than from the run's own totals,
 * because the totals include scratch and the floor must not. A number that
 * depends on what happens to be lying in the directory is the kind of gate that
 * fails on somebody else's machine for no reason anybody can act on.
 *
 * @param {unknown} report
 * @returns {{ defined: number, skipped: number }}
 */
export function counts(report) {
  const all = results(report);
  return {
    defined: all.length,
    skipped: all.filter((test) => test.status === "pending" || test.status === "skipped").length,
  };
}

/**
 * The names of the skipped tests, so a reviewer sees WHICH ones in the log.
 *
 * A cap says how many; a reviewer needs to know which, because "1 skipped" is
 * the same line whether it is the symlink test this machine cannot run or a
 * case somebody switched off to get a green run.
 *
 * @param {unknown} report
 * @returns {string[]}
 */
export function skippedNames(report) {
  return results(report)
    .filter((test) => test.status === "pending" || test.status === "skipped")
    .map((test) => test.fullName ?? test.title ?? "?");
}

/**
 * The names of the tests that FAILED, so a run that goes red says which.
 *
 * `--reporter=json` writes everything to a FILE and nothing to the console, and
 * `execFileSync` throws on a non-zero exit — so before this, a failing run
 * printed a Node stack trace about `execFileSync` and not one word about the
 * suite. Measured on 2026-09-12: `main` went red exactly that way after #73
 * merged, and the log gave no way at all to tell which test had failed. The
 * coverage step immediately above it had passed all 1196.
 *
 * Scratch files are NOT filtered out here, unlike in `counts` above, and the
 * difference is deliberate: the floor is a number about the repository, but the
 * exit code is a fact about the run. A failure in a scratch file is why the
 * command failed, and hiding it would leave the same silence this exists to
 * end.
 *
 * A file that failed before defining any test — an import that threw, a worker
 * that died — has no failed assertion to name, and is precisely the case a bare
 * stack trace leaves the reader guessing about. It is named by its own message,
 * bounded, because a report is not a place to paste a stack.
 *
 * @param {unknown} report
 * @returns {string[]}
 */
export function failedNames(report) {
  const files =
    /** @type {{ name?: string, status?: string, message?: string, assertionResults?: { status?: string, fullName?: string, title?: string }[] }[]} */ (
      /** @type {{ testResults?: unknown }} */ (report ?? {}).testResults ?? []
    );
  return files.flatMap((file) => {
    const where = file.name ?? "?";
    const cases = (file.assertionResults ?? []).filter((test) => test.status === "failed");
    if (cases.length) return cases.map((test) => `${where} > ${test.fullName ?? test.title ?? "?"}`);
    if (file.status !== "failed") return [];
    const why = (file.message ?? "no message").replace(/\s+/g, " ").slice(0, 200);
    return [`${where} (no test ran: ${why})`];
  });
}

/**
 * What to print when the run exited non-zero, from the report it left behind.
 *
 * The second sentence is for a non-zero exit with NOTHING failed in the report.
 * A worker the kernel killed looks exactly like that: the tests it had already
 * finished are in the file, the ones it never reached are simply absent, and
 * vitest exits 1 with no failure to point at. Saying how much of the suite the
 * report accounts for is what separates that from a run that failed for a
 * reason outside the tests, and it is the only thing the log can offer.
 *
 * @param {unknown} report
 * @returns {string}
 */
export function failureMessage(report) {
  const names = failedNames(report);
  if (names.length) return `the suite failed —\n  ${names.join("\n  ")}`;
  const { defined, skipped } = counts(report);
  return (
    `the run exited non-zero with no failed test in its report — ${defined} tests accounted for, ` +
    `${skipped} skipped. A worker that died or ran out of memory reads like this.`
  );
}

/**
 * Say what failed. False only when there is no report to read at all.
 *
 * @param {string} out
 * @returns {boolean}
 */
function reportFailures(out) {
  let report;
  try {
    report = JSON.parse(readFileSync(out, "utf8"));
  } catch {
    return false;
  }
  console.error(`test-count: ${failureMessage(report)}`);
  return true;
}

function main() {
  const update = process.argv.includes("--update");
  const out = join(mkdtempSync(join(tmpdir(), "ssf-slide-elements-")), "results.json");
  // Vitest's own entry point through the running Node, not `npx`. `execFileSync`
  // does not go through a shell, and on Windows the executable is `npx.cmd`, so
  // spawning `npx` there is `ENOENT` before any test runs — this gate could not
  // be run at all on the owner's machine while CI on ubuntu passed it. Naming
  // the installed entry also guarantees the pinned vitest rather than whatever
  // `npx` would resolve, which is what a floor under the suite wants anyway.
  try {
    execFileSync(
      process.execPath,
      [join("node_modules", "vitest", "vitest.mjs"), "run", "--reporter=json", `--outputFile=${out}`],
      { stdio: "inherit" },
    );
  } catch (error) {
    // The run failed. Everything it knows is in `out` and nothing is on the
    // console, so read it back rather than letting the throw surface as a
    // stack trace about `execFileSync`. See `failedNames` for the run that
    // forced this.
    if (!reportFailures(out)) {
      console.error(`test-count: the suite failed and left no readable report: ${String(error)}`);
    }
    process.exit(1);
  }

  const report = JSON.parse(readFileSync(out, "utf8"));
  const record = JSON.parse(readFileSync(RECORD, "utf8"));
  // From the per-file results, never `numTotalTests`: the run's own totals
  // count a scratch `test/zz…` file, and the floor is a number about the
  // repository.
  const { defined, skipped } = counts(report);
  const answer = verdict({ defined, skipped, record, update });

  const names = skippedNames(report);
  if (names.length) console.log(`test-count: skipped — ${names.join("; ")}`);

  if (!answer.ok) {
    console.error(`test-count: ${answer.message}`);
    process.exit(1);
  }
  if (answer.write) writeFileSync(RECORD, `${JSON.stringify(answer.write, null, 2)}\n`);
  console.log(`test-count: ${answer.message}`);
}

if (isMain(import.meta.url)) main();
