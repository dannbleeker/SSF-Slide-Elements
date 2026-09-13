import { describe, expect, it } from "vitest";
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
