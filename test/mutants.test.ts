import { writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types, shared with the scripts.
import * as tool from "../scripts/mutants.mjs";

const codeMask = tool.codeMask as (text: string) => string;
const mutationsOf = tool.mutationsOf as (text: string) => { at: number; was: string; now: string; what: string }[];
const judgeSurvivor = tool.judgeSurvivor as (where: string) => { known: boolean; why: string };
const staleEquivalents = tool.staleEquivalents as (survivors: string[], swept: string[]) => string[];
const EQUIVALENT = tool.EQUIVALENT as { file: string; what: string; was: string; why: string }[];
const FAST = tool.FAST as Record<string, string[]>;
const TARGETS = tool.TARGETS as string[];
const fastTests = tool.fastTests as (file: string) => string[];
const testsReaching = tool.testsReaching as (file: string) => string[];
const verdictOfFailure = tool.verdictOfFailure as (error: unknown, out: string) => "killed" | "inconclusive";
const failedFilesOf = tool.failedFilesOf as (out: string) => string[];
const strayPids = tool.strayPids as (listing: string, workspace: string, self: number) => number[];
const confirmedKill = tool.confirmedKill as (
  blamed: string[],
  rerun: (files: string[]) => "survived" | "killed" | "inconclusive",
) => "survived" | "killed";

/**
 * The mutation sweep, which is a tool rather than a gate and therefore has to be
 * held by something other than CI running it.
 *
 * The two functions worth testing are the two that decide WHAT gets mutated. A
 * mask that blanks the wrong span silently mutates a comment — which always
 * survives, and is reported as a hole in the suite that is really a hole in the
 * tool. That exact class of error already cost this sweep one wrong report.
 */
describe("the mask that says which characters are code", () => {
  it("keeps every offset, because a mutation is found in the mask and applied to the original", () => {
    const text = 'const a = "hello"; // note\nconst b = 1;\n';
    expect(codeMask(text)).toHaveLength(text.length);
  });

  it("blanks a string's contents and leaves the code around it", () => {
    const masked = codeMask('if (x < "a < b") return 1;');
    expect(masked).toContain("if (x < ");
    expect(masked, "the comparison inside the literal is not code").not.toContain("a < b");
  });

  it("blanks a line comment and a block comment", () => {
    expect(codeMask("const a = 1; // x < y\n")).not.toContain("x < y");
    expect(codeMask("/* x < y */\nconst a = 1;\n")).not.toContain("x < y");
  });

  it("blanks a regex literal, whose contents are not operators either", () => {
    // A regex is recognised by what precedes it. Getting this wrong is how a
    // sweep comes to mutate a character class and call the result a survivor.
    expect(codeMask("const r = /a<b/;\n")).not.toContain("a<b");
  });
});

describe("which mutations it finds", () => {
  it("flips a comparison in code and ignores the same characters in prose", () => {
    const found = mutationsOf('if (n < 3) return "n < 3";\n');
    const boundaries = found.filter((m) => m.what === "boundary");
    expect(boundaries, "one comparison, not two").toHaveLength(1);
    expect(boundaries[0]?.now).toBe("<=");
  });

  it("does not mistake an arrow or a generic for a comparison", () => {
    const found = mutationsOf("const f = (n: number) => n;\n");
    expect(found.filter((m) => m.what === "boundary")).toEqual([]);
  });

  // The case above says "generic" and contains none: `=> n` has no angle
  // bracket at all, so it passed while the sweep was reporting ten survivors in
  // a file of nothing but `Record<string, string>`. These are the generics.
  it("leaves a type argument alone, however many angle brackets it has", () => {
    const source = [
      "const m = new Map<string, number>();\n",
      "type Held = Record<string, Promise<Uint8Array | string>>;\n",
      "function f<T>(x: Promise<T>): Promise<T[]> {\n  return x.then((v) => [v]);\n}\n",
    ].join("");
    expect(mutationsOf(source).filter((m) => m.what === "boundary")).toEqual([]);
  });

  it("leaves a shift alone, which is not a comparison either", () => {
    const source = "const a = (hash >>> 0).toString(16);\nconst b = (n >> 16) & 255;\n";
    expect(mutationsOf(source).filter((m) => m.what === "boundary")).toEqual([]);
  });

  it("still finds the comparison standing next to a generic", () => {
    const source = 'const seen = new Set<string>();\nif (seen.size > 0) return "some";\n';
    const boundaries = mutationsOf(source).filter((m) => m.what === "boundary");
    expect(boundaries.map((m) => `${m.was} -> ${m.now}`)).toEqual(["> -> >="]);
  });

  it("finds a number to move, a boolean to flip and a fallback to delete", () => {
    const kinds = new Set(mutationsOf("const a = b ?? 0;\nif (c && d) return 7;\n").map((m) => m.what));
    expect(kinds).toContain("off-by-one");
    expect(kinds).toContain("boolean");
    expect(kinds).toContain("fallback");
  });
});

describe("a comparison read backwards", () => {
  /**
   * The operator that moves a comparison's SENSE rather than its edge.
   *
   * The boundary operator above cannot stand in for it: `<` to `<=` moves where
   * the edge falls, and a test that pins the edge kills it. Swapping the
   * operands asks a different question — whether anything distinguishes "the
   * element fits in the room" from "the room fits in the element" — and in this
   * family that is the question five overlap defects in SSF-Charts turned out to
   * be.
   */
  const swaps = (source: string) =>
    mutationsOf(source)
      .filter((m) => m.what === "operands")
      .map((m) => `${m.was} -> ${m.now}`);

  it("swaps a comparison of two names", () => {
    expect(swaps("if (rect.cx <= room.x) return true;\n")).toEqual(["rect.cx <= room.x -> room.x <= rect.cx"]);
  });

  it("swaps a loop bound, where the swap stops the loop running at all", () => {
    expect(swaps("for (let i = 0; i < paths.length; i++) {\n}\n")).toEqual(["i < paths.length -> paths.length < i"]);
  });

  it("swaps a name against a number, which is a sense flip and not a boundary", () => {
    expect(swaps("if (slide < 1) return;\n")).toEqual(["slide < 1 -> 1 < slide"]);
  });

  it("leaves a generic, an arrow and a shift alone, like the boundary operator", () => {
    const source = [
      "const m = new Map<string, number>();\n",
      "type Held = Record<string, Promise<Uint8Array>>;\n",
      "const f = (n: number) => n;\n",
      "const a = (hash >>> 0).toString(16);\n",
      "const b = (n >> 16) & 255;\n",
    ].join("");
    expect(swaps(source)).toEqual([]);
  });

  it("leaves an operand it cannot read alone rather than rearranging it wrongly", () => {
    // A call, an index and an expression each need a parser to move safely. The
    // cost of skipping one is a mutant nobody tried; the cost of moving it
    // wrongly is a mutant that means something else, reported as a survivor.
    const source = ["if (widthOf(a) < b) return;\n", "if (a.slides[0] < b) return;\n", "if (a + 1 < b) return;\n"].join(
      "",
    );
    expect(swaps(source)).toEqual([]);
  });

  it("skips a comparison straight after a keyword, which is the price of the rule above", () => {
    // `return a < b`: the character before the left operand is the `n` of
    // `return`, and admitting word characters there is exactly what would let a
    // match start inside a name and swap a fragment. So this site is skipped,
    // knowingly. It is recorded here so the next reader finds a decision rather
    // than a hole — and so that widening the rule has to change a test.
    expect(swaps("return a < b;\n")).toEqual([]);
    // The same comparison inside an `if` IS swept, which is how it is written
    // nearly everywhere in this repository.
    expect(swaps("if (a < b) return;\n")).toEqual(["a < b -> b < a"]);
  });

  it("says nothing about a comparison of one thing with itself", () => {
    // `n > n` swapped is `n > n`. A mutant identical to its original cannot be
    // killed by any test ever written, so emitting it would put a permanent
    // survivor in the report — the one failure mode that makes a sweep's clean
    // section worthless.
    expect(swaps("if (n > n) return;\n")).toEqual([]);
  });

  it("ignores the same characters in prose, as every operator here must", () => {
    expect(swaps('// a < b here\nconst s = "x < y";\n')).toEqual([]);
  });

  it("is a different mutation from the boundary at the same place", () => {
    // Both operators fire on one comparison, and they must stay distinguishable:
    // the equivalence ledger is keyed on the operator and the text, so two
    // operators reporting the same `was` under the same name would let one
    // entry silently excuse the other.
    const found = mutationsOf("if (shared < to.length) return;\n");
    const kinds = found.filter((m) => m.what === "boundary" || m.what === "operands");
    expect(kinds.map((m) => m.what).sort()).toEqual(["boundary", "operands"]);
    expect(kinds.find((m) => m.what === "boundary")?.was).toBe("<");
    expect(kinds.find((m) => m.what === "operands")?.was).toBe("shared < to.length");
  });

  it("produces source that still parses, because a swap is a rewrite and not a substitution", () => {
    // Every other operator replaces one token. This one moves two spans, so the
    // text it writes is worth checking rather than assuming.
    const source = "if (rect.cy <= into.cy) return true;\n";
    const one = mutationsOf(source).find((m) => m.what === "operands");
    if (!one) throw new Error("no swap found");
    const mutated = source.slice(0, one.at) + one.now + source.slice(one.at + one.was.length);
    expect(mutated).toBe("if (into.cy <= rect.cy) return true;\n");
  });
});

describe("the record of mutations that cannot be killed", () => {
  it("labels a survivor it has a reason for", () => {
    const judged = judgeSurvivor('src/host/jump.ts:49  boundary  "<" -> "<="');
    expect(judged.known).toBe(true);
    expect(judged.why, "and the reason is a measurement, not a shrug").toContain("1,003,578");
  });

  it("does not label a survivor in the same file with a different operator", () => {
    expect(judgeSurvivor('src/host/jump.ts:49  boolean  "&&" -> "||"').known).toBe(false);
  });

  it("does not label a survivor in a different file with the same operator", () => {
    expect(judgeSurvivor('src/host/insert.ts:94  boundary  "<" -> "<="').known).toBe(false);
  });

  // The entry recorded for `<` is about widening it to `<=`. A report line for
  // the opposite mutation, `"<=" -> "<"`, CONTAINS the text `"<"`, and matching
  // on that substring labelled it as the same known-unkillable line.
  it("does not label the opposite mutation, whose report line contains the same text", () => {
    expect(judgeSurvivor('src/host/jump.ts:49  boundary  "<=" -> "<"').known).toBe(false);
  });

  it("names a recorded entry that matched nothing, rather than letting it rot", () => {
    // The ledger's own guard. An entry that stops matching means the code moved
    // or a test now kills it, and either way the reasoning needs reading again —
    // which nobody will do if the report stays quiet about it.
    const all = [...new Set(EQUIVALENT.map((one) => one.file))];
    expect(staleEquivalents([], all), "every entry is stale against an empty run").toHaveLength(EQUIVALENT.length);
    expect(staleEquivalents(['src/pane/storage.ts:40  boundary  ">" -> ">="'], all)).toHaveLength(
      EQUIVALENT.length - 1,
    );
  });

  it("says nothing about a file the run did not sweep", () => {
    // Found by using it. `--only host/jump,host/probe,pane/search` on 2026-09-13
    // reported `src/pane/storage.ts` as gone stale — nonsense, because storage
    // was never swept, so of course nothing of its matched. A gate that cries
    // wolf on every partial run teaches the reader to skim past it, which is the
    // same failure as a gate that cannot fail at all.
    const swept = ["src/host/jump.ts"];
    for (const one of staleEquivalents([], swept)) {
      expect(one, "only a swept file may be judged").toContain("src/host/jump.ts");
    }
    expect(staleEquivalents([], []), "a run that swept nothing judges nothing").toEqual([]);
  });

  it("gives every entry a reason long enough to be one", () => {
    for (const one of EQUIVALENT) {
      expect(one.why.length, `${one.file} ${one.what} ${one.was}`).toBeGreaterThan(80);
    }
  });
});

describe("what a run that failed is allowed to mean", () => {
  const report = join(tmpdir(), "ssf-mutants-test-report.json");

  // A real kill: the report on disk names a test that failed.
  beforeEach(() => {
    writeFileSync(
      report,
      JSON.stringify({
        testResults: [{ name: "test/x.test.ts", assertionResults: [{ status: "failed", fullName: "x" }] }],
      }),
    );
  });

  it("reads a named failure as a kill", () => {
    expect(verdictOfFailure(new Error("exit 1"), report)).toBe("killed");
  });

  // A run that never finished says nothing a REPORT can answer, and the report
  // beside it belongs to the PREVIOUS mutant — so reading it would answer this
  // one with the last one's result. It is a hang, not a kill, and the stale
  // report here names a failed test precisely so that reading it would say so.
  it("reads a timed-out run as a hang, whatever the stale report says", () => {
    const timedOut = Object.assign(new Error("ETIMEDOUT"), { code: "ETIMEDOUT" });
    expect(verdictOfFailure(timedOut, report)).toBe("hung");
  });

  it("reads a report naming no failure as inconclusive rather than a kill", () => {
    writeFileSync(report, JSON.stringify({ testResults: [] }));
    expect(verdictOfFailure(new Error("exit 1"), report)).toBe("inconclusive");
  });
});

describe("a kill asked a second time", () => {
  const report = join(tmpdir(), "ssf-mutants-test-blame.json");
  const here = (name: string) => join(process.cwd(), name);

  it("names the test files a failed run blamed, and only those", () => {
    writeFileSync(
      report,
      JSON.stringify({
        testResults: [
          { name: here("test/passed.test.ts"), status: "passed", assertionResults: [{ status: "passed" }] },
          { name: here("test/blamed.test.ts"), status: "failed", assertionResults: [{ status: "failed" }] },
        ],
      }),
    );
    expect(failedFilesOf(report)).toEqual(["test/blamed.test.ts"]);
  });

  // A file the runner could not even start reports no assertions at all. It is
  // still the file to re-run.
  it("names a file that failed without running a test", () => {
    writeFileSync(
      report,
      JSON.stringify({ testResults: [{ name: here("test/broken.test.ts"), status: "failed", assertionResults: [] }] }),
    );
    expect(failedFilesOf(report)).toEqual(["test/broken.test.ts"]);
  });

  it("names nothing when the report cannot be read", () => {
    expect(failedFilesOf(join(tmpdir(), "ssf-mutants-no-such-report.json"))).toEqual([]);
  });

  it("keeps a kill the blamed file reports again", () => {
    expect(confirmedKill(["test/blamed.test.ts"], () => "killed")).toBe("killed");
  });

  // The whole point. A first run that failed for its own reasons — a flake on a
  // loaded machine — passes on its own, and the mutant goes back to being a
  // survivor so the whole-suite re-check gets the last word.
  it("turns a kill the blamed file does not repeat back into a survivor", () => {
    expect(confirmedKill(["test/blamed.test.ts"], () => "survived")).toBe("survived");
  });

  it("treats an unclear second answer as a survivor too, because that answer gets re-checked", () => {
    expect(confirmedKill(["test/blamed.test.ts"], () => "inconclusive")).toBe("survived");
  });

  it("keeps a kill that blamed no file, having nothing to re-run", () => {
    expect(
      confirmedKill([], () => {
        throw new Error("must not re-run when no file was named");
      }),
    ).toBe("killed");
  });
});

describe("the vitest workers a timed-out run leaves behind", () => {
  // `execFileSync`'s timeout signals the vitest process it started. Vitest runs
  // the tests in FORKS, and those are not signalled — they are re-parented to
  // init and keep going. When the mutation removed a loop's only exit, they
  // keep going forever: fifteen of them span for two hours on 2026-09-14 and
  // took a four-core machine to a load average of eighteen.
  const ws = "/tmp/ssf-mutants-abc123";
  const listing = [
    "  101 /opt/node22/bin/node " + ws + "/node_modules/vitest/dist/workers/forks.js",
    "  102 /opt/node22/bin/node --require " + ws + "/node_modules/vitest/suppress-warnings.cjs",
    "  103 /opt/node22/bin/node node_modules/vitest/vitest.mjs run --outputFile=/tmp/other/report.json",
    "  104 /usr/bin/some-daemon --unrelated",
  ].join("\n");

  it("names the workers of this workspace and leaves other runs alone", () => {
    expect(strayPids(listing, ws, 999)).toEqual([101, 102]);
  });

  // The trap this replaces: `pkill -f <pattern>` matches the shell that is
  // RUNNING it, because the pattern is in that shell's own command line. It
  // cost two killed wrappers and one orphaned sweep in a single session.
  it("never names the process doing the asking, even when its own command line matches", () => {
    const asker = "  777 /bin/bash -c ps -eo pid,args | grep " + ws + "/node_modules/vitest";
    expect(strayPids(listing + "\n" + asker, ws, 777)).toEqual([101, 102]);
  });

  it("ignores a process of this workspace that is not vitest at all", () => {
    const editor = "  555 /usr/bin/vim " + ws + "/src/core/pptx/pkg.ts";
    expect(strayPids(listing + "\n" + editor, ws, 999)).toEqual([101, 102]);
  });
});

describe("what a run that failed is allowed to mean, when it never finished", () => {
  const report = join(tmpdir(), "ssf-mutants-hung-report.json");

  it("still reads a named failure as a kill", () => {
    writeFileSync(
      report,
      JSON.stringify({
        testResults: [{ name: "test/x.test.ts", assertionResults: [{ status: "failed", fullName: "x" }] }],
      }),
    );
    expect(verdictOfFailure(new Error("exit 1"), report)).toBe("killed");
  });
});

describe("the first-tier test map", () => {
  /**
   * `FAST` is an optimisation, and the only thing that makes it safe is the
   * whole-suite re-check every survivor still gets. So these cases hold the two
   * things that would make an entry USELESS rather than unsafe: naming a test
   * that cannot see the file, and naming a file that is not swept at all.
   */
  it("names only files the sweep actually mutates", () => {
    for (const file of Object.keys(FAST)) {
      expect(TARGETS, `${file} has a FAST entry but is not a target`).toContain(file);
    }
  });

  it("names only tests that can actually reach the file", () => {
    // A test outside the reaching set cannot observe a mutation of this file, so
    // an entry naming one is dead weight in every run. It is not a safety
    // problem — the whole-suite re-check catches what a narrow tier misses — but
    // a map nobody can trust to mean something stops being read.
    for (const [file, tests] of Object.entries(FAST)) {
      const reaching = testsReaching(file);
      expect(tests, `${file} has an empty FAST entry`).not.toHaveLength(0);
      for (const one of tests) {
        expect(reaching, `${one} does not reach ${file}`).toContain(one);
      }
    }
  });

  it("falls back to the whole reaching set for a file with no entry", () => {
    const without = TARGETS.find((f) => !(f in FAST));
    expect(without, "this case needs at least one file without an entry").toBeDefined();
    if (without !== undefined) expect(fastTests(without)).toEqual(testsReaching(without));
  });
});
