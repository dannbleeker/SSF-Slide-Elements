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

  // A run that never finished says nothing about the mutant, and the report
  // beside it belongs to the PREVIOUS mutant — so reading it would answer this
  // one with the last one's result.
  it("reads a timed-out run as inconclusive, whatever the stale report says", () => {
    const timedOut = Object.assign(new Error("ETIMEDOUT"), { code: "ETIMEDOUT" });
    expect(verdictOfFailure(timedOut, report)).toBe("inconclusive");
  });

  it("reads a report naming no failure as inconclusive rather than a kill", () => {
    writeFileSync(report, JSON.stringify({ testResults: [] }));
    expect(verdictOfFailure(new Error("exit 1"), report)).toBe("inconclusive");
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
