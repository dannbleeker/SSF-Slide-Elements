import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types, shared with the scripts.
import * as tool from "../scripts/test-count.mjs";

interface Record_ {
  min: number;
  maxSkipped?: number;
}
interface Answer {
  ok: boolean;
  message: string;
  write?: { min: number; maxSkipped: number } | null;
}
const verdict = tool.verdict as (o: { defined: number; skipped: number; record: Record_; update?: boolean }) => Answer;
const skippedNames = tool.skippedNames as (report: unknown) => string[];
const counts = tool.counts as (report: unknown) => { defined: number; skipped: number };
const failedNames = tool.failedNames as (report: unknown) => string[];
const failureMessage = tool.failureMessage as (report: unknown) => string;

/**
 * The gate that holds a floor under the suite, which had no test of its own.
 *
 * It is the one gate whose failure is invisible: every other check fails loudly
 * when it breaks, and this one just stops noticing. In a sibling it was wrong twice —
 * once counting pending tests, so 23 could be switched off in silence, and once
 * counting only tests that RAN, which made the number depend on the machine.
 */
describe("the floor under the suite", () => {
  const record = { min: 100, maxSkipped: 1 };

  it("passes a run that matches", () => {
    expect(verdict({ defined: 100, skipped: 0, record }).ok).toBe(true);
  });

  it("raises the floor on its own when the suite grows", () => {
    const answer = verdict({ defined: 120, skipped: 0, record });
    expect(answer.ok).toBe(true);
    expect(answer.write).toEqual({ min: 120, maxSkipped: 1 });
  });

  it("refuses a suite that has lost tests", () => {
    const answer = verdict({ defined: 77, skipped: 0, record });
    expect(answer.ok).toBe(false);
    expect(answer.message).toContain("down from 100");
  });

  it("refuses 23 tests switched off, which is what it is for", () => {
    // `it(` to `it.skip(` leaves the total alone, so the floor cannot see it.
    // The cap is what does, and this is the case that motivated the whole gate:
    // the merge plan's decision engine, silent, every check green.
    const answer = verdict({ defined: 100, skipped: 23, record });
    expect(answer.ok).toBe(false);
    expect(answer.message).toContain("23 tests are skipped");
  });

  it("is the same number on a machine that must skip one", () => {
    /**
     * The reason it counts tests that EXIST rather than tests that ran.
     *
     * `is-main.test.ts` needs a symlink and Windows refuses that without
     * elevation, so counting only what RAN gives one fewer on Windows than in
     * CI, and either number committed breaks the other machine. The figures
     * below are a sibling's, where this was found; this repo's own floor is in
     * `test/fixtures/test-count.json` and has never been near them. What the
     * case is about is the mechanism, not the size.
     */
    const ci = verdict({ defined: 1476, skipped: 0, record: { min: 1476, maxSkipped: 1 } });
    const windows = verdict({ defined: 1476, skipped: 1, record: { min: 1476, maxSkipped: 1 } });
    expect(ci.ok, "CI runs every test").toBe(true);
    expect(windows.ok, "and the machine that cannot make a symlink agrees").toBe(true);
    expect(windows.write, "neither rewrites the record").toBeNull();
    expect(ci.write).toBeNull();
  });

  it("does not raise the skip cap by itself", () => {
    // A suite growing is ordinary; a new skip is a decision somebody should be
    // seen making. Only `--update` moves this one.
    const answer = verdict({ defined: 130, skipped: 0, record });
    expect(answer.write).toEqual({ min: 130, maxSkipped: 1 });
  });

  it("records a deliberate drop, and a deliberate skip, only with --update", () => {
    const answer = verdict({ defined: 90, skipped: 4, record, update: true });
    expect(answer.ok).toBe(true);
    expect(answer.write).toEqual({ min: 90, maxSkipped: 4 });
  });

  it("treats a record with no cap as allowing none", () => {
    // Written before the cap existed. Infinity would silently accept whatever
    // was skipped on the day somebody upgraded.
    const answer = verdict({ defined: 100, skipped: 1, record: { min: 100 } });
    expect(answer.ok).toBe(false);
    expect(answer.message).toContain("allows 0");
  });

  it("refuses a count that is not a count rather than comparing it", () => {
    for (const defined of [NaN, 0, -1, 1.5, undefined as unknown as number]) {
      expect(verdict({ defined, skipped: 0, record }).ok, `defined ${defined}`).toBe(false);
    }
    for (const skipped of [NaN, -1, 1.5, undefined as unknown as number]) {
      expect(verdict({ defined: 100, skipped, record }).ok, `skipped ${skipped}`).toBe(false);
    }
  });
});

describe("naming what was skipped", () => {
  it("lists them, so the log says which rather than how many", () => {
    const report = {
      testResults: [
        {
          assertionResults: [
            { status: "passed", fullName: "a passing one" },
            { status: "pending", fullName: "whether a module is the entry point > through a SYMLINK" },
          ],
        },
        { assertionResults: [{ status: "skipped", fullName: "another" }] },
      ],
    };
    expect(skippedNames(report)).toEqual(["whether a module is the entry point > through a SYMLINK", "another"]);
  });

  it("answers nothing for a report with no results", () => {
    expect(skippedNames({})).toEqual([]);
    expect(skippedNames(undefined)).toEqual([]);
  });
});

describe("a scratch file in the directory", () => {
  /**
   * `test/zz…` is gitignored, and vitest runs it like any other file. So the
   * machine that RECORDS the floor counts it, and the next CI run — which has
   * no such file — fails on a number no commit contains.
   *
   * It is not hypothetical. 989 was recorded against a CI that defined 988, and
   * the difference was one scratch file left in the directory after a probe.
   */
  const report = {
    testResults: [
      {
        // A Windows path, backslashes and all: vitest reports the file the way
        // the platform spells it, and a rule written for one separator only is
        // a rule that works on CI and not on the machine recording the floor.
        name: "C:\\repo\\test\\splice.test.ts",
        assertionResults: [
          { status: "passed", fullName: "a committed one" },
          { status: "pending", fullName: "the symlink one" },
        ],
      },
      {
        name: "/repo/test/zz-dbg.test.ts",
        assertionResults: [
          { status: "passed", fullName: "a scratch one" },
          { status: "skipped", fullName: "another scratch one" },
        ],
      },
    ],
  };

  it("is left out of the count the floor is compared against", () => {
    // The committed file's two, and neither of the scratch file's two. Counting
    // the lot answers `{ defined: 4, skipped: 2 }`, which is the number that
    // gets recorded and the number CI then cannot reach.
    expect(counts(report)).toEqual({ defined: 2, skipped: 1 });
  });

  it("is left out of the skipped names too, so the log names what CI will see", () => {
    expect(skippedNames(report)).toEqual(["the symlink one"]);
  });

  it("counts a committed file whose name merely CONTAINS zz", () => {
    // The convention is a file whose NAME starts with `zz`, matching
    // `.gitignore`'s `test/zz*`. A committed `test/fuzz.test.ts` is not scratch,
    // and a rule that matched it anywhere in the path would quietly stop
    // counting a real file — which is the same defect in the other direction.
    const committed = {
      testResults: [
        { name: "/repo/test/fuzz.test.ts", assertionResults: [{ status: "passed", fullName: "fuzzing" }] },
        { name: "/repo/zz-elsewhere/test/real.test.ts", assertionResults: [{ status: "passed", fullName: "real" }] },
      ],
    };
    expect(counts(committed)).toEqual({ defined: 2, skipped: 0 });
  });
});

describe("what a failing run says", () => {
  /**
   * The gate ran the suite with `--reporter=json`, which writes everything to a
   * FILE and nothing to the console, and `execFileSync` throws on a non-zero
   * exit — so a red run printed a Node stack trace about `execFileSync` and not
   * one word about the suite.
   *
   * Measured on 2026-09-12: `main` went red exactly that way, and the log gave
   * no way at all to tell which test had failed. The coverage step immediately
   * above it, running the same suite, had passed all 1196.
   */
  it("names the failed test and the file it is in", () => {
    const report = {
      testResults: [
        {
          name: "/repo/test/host-insert.test.ts",
          assertionResults: [
            { status: "passed", fullName: "an insert that worked > says so" },
            { status: "failed", fullName: "an insert that raised > reads the delta" },
          ],
        },
      ],
    };
    expect(failedNames(report)).toEqual(["/repo/test/host-insert.test.ts > an insert that raised > reads the delta"]);
  });

  it("says nothing about a run that passed", () => {
    const report = {
      testResults: [{ name: "/repo/test/a.test.ts", assertionResults: [{ status: "passed", fullName: "fine" }] }],
    };
    expect(failedNames(report)).toEqual([]);
  });

  it("names a file that failed before defining a single test", () => {
    // An import that threw or a worker that died leaves no failed assertion to
    // name, and is exactly the case a bare stack trace leaves the reader
    // guessing about.
    const report = {
      testResults: [
        {
          name: "/repo/test/broken.test.ts",
          status: "failed",
          message: "Error: Cannot find module\n    at load (node:internal)",
          assertionResults: [],
        },
      ],
    };
    expect(failedNames(report)).toEqual([
      "/repo/test/broken.test.ts (no test ran: Error: Cannot find module at load (node:internal))",
    ]);
  });

  it("bounds the message, because a report is not a place to paste a stack", () => {
    const report = {
      testResults: [
        { name: "/repo/test/broken.test.ts", status: "failed", message: "x".repeat(5000), assertionResults: [] },
      ],
    };
    const named = failedNames(report)[0] ?? "";
    // Named first, then bounded. `?? ""` is a length of zero, so a version that
    // reported nothing at all would satisfy the bound and this case would pass
    // over the very silence it exists to prevent — measured, on the break that
    // removed the file-level branch.
    expect(named, "the file is still named").toContain("/repo/test/broken.test.ts");
    expect(named.length, "and the stack is not pasted into the log").toBeLessThan(300);
  });

  it("names a scratch file's failure, unlike the count the floor uses", () => {
    // `counts` leaves `test/zz…` out on purpose: the floor is a number about
    // the repository. The exit code is not — it is a fact about the RUN, and a
    // scratch file that fails is why the command failed. Leaving it out here
    // would recreate the same silence in the one case a developer is most
    // likely to hit.
    const report = {
      testResults: [
        { name: "/repo/test/zz-dbg.test.ts", assertionResults: [{ status: "failed", fullName: "scratch" }] },
      ],
    };
    expect(counts(report)).toEqual({ defined: 0, skipped: 0 });
    expect(failedNames(report)).toEqual(["/repo/test/zz-dbg.test.ts > scratch"]);
  });

  it("says how much of the suite the report accounts for when nothing in it failed", () => {
    // A non-zero exit with no failed test is what a worker the kernel killed
    // looks like: the tests it finished are in the file, the ones it never
    // reached are simply absent, and vitest exits 1 with nothing to point at.
    // Measured on 2026-09-12: `main` went red twice this way, and the report
    // was the only thing that could have said which of the two it was.
    const report = {
      testResults: [{ name: "/repo/test/a.test.ts", assertionResults: [{ status: "passed", fullName: "fine" }] }],
    };
    const said = failureMessage(report);
    expect(said, "it does not claim a test failed").not.toContain("the suite failed");
    expect(said, "and it says how far the report got").toContain("1 tests accounted for");
  });

  it("names the failures rather than counting, when there are any", () => {
    const report = {
      testResults: [{ name: "/repo/test/a.test.ts", assertionResults: [{ status: "failed", fullName: "broke" }] }],
    };
    const said = failureMessage(report);
    expect(said).toContain("/repo/test/a.test.ts > broke");
    expect(said, "the count sentence is the OTHER case, not both").not.toContain("accounted for");
  });
});
