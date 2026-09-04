/**
 * The layering, held in both directions.
 *
 * `src/core` is pure and `src/host` decides; `src/office` calls and imports
 * every judgement from `src/host`. Both halves matter: an Office.js import in
 * the pure layers makes a rule untestable, and a rule reimplemented inline in
 * `src/office` looks tidier and rots quietly. This is SSF-Merge's split, and it
 * is why anything in this repo can be checked without a PowerPoint at all.
 */
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutProse } from "../scripts/without-prose.mjs";

function filesUnder(dir: string): string[] {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) out.push(...filesUnder(path));
    else if (name.endsWith(".ts")) out.push(path);
  }
  return out;
}

/**
 * Naming one of the globals in CODE.
 *
 * Read against the file with its comments and strings removed. Every file here
 * talks about Office.js in prose — "hand the result back to PowerPoint" — and a
 * guard that counts those reports the best-documented files as the worst
 * offenders. The first version of this test did exactly that, on five files
 * that are entirely clean.
 */
const OFFICE = /\b(Office|PowerPoint)\s*\./;

function code(file: string): string {
  return withoutProse(readFileSync(file, "utf8")) as string;
}

describe("src/core is pure", () => {
  it("imports nothing from Office.js and names neither global", () => {
    const offenders = filesUnder(join(process.cwd(), "src", "core")).filter((f) => OFFICE.test(code(f)));
    expect(offenders).toEqual([]);
  });
});

describe("src/host decides", () => {
  it("makes no Office.js calls of its own", () => {
    const offenders = filesUnder(join(process.cwd(), "src", "host")).filter((f) => OFFICE.test(code(f)));
    expect(offenders).toEqual([]);
  });
});

describe("src/office only calls", () => {
  /**
   * The direction that rots quietly. A judgement inlined here — a version
   * string compared, a count checked, a message capped — is a rule the suite
   * cannot reach, and it looks perfectly reasonable in review.
   */
  it("imports its judgements from src/host", () => {
    for (const f of filesUnder(join(process.cwd(), "src", "office"))) {
      expect(readFileSync(f, "utf8")).toMatch(/from "\.\.\/host\//);
    }
  });
});

describe("src/pane", () => {
  it("touches Office.js in main.ts and nowhere else", () => {
    const offenders = filesUnder(join(process.cwd(), "src", "pane"))
      .filter((f) => !f.endsWith("main.ts"))
      .filter((f) => OFFICE.test(code(f)));
    expect(offenders).toEqual([]);
  });
});

describe("the guard itself", () => {
  /**
   * A gate that cannot fail is decoration. `withoutProse` is the whole reason
   * these checks are green, so it has to be shown removing the right thing and
   * keeping the right thing — otherwise a stripper that returned "" would pass
   * every test above.
   */
  it("strips prose that names a global", () => {
    const prose = `// hand it back to PowerPoint.\n/* Office echoes it back. */\nconst s = "PowerPoint.run";\n`;
    expect(OFFICE.test(withoutProse(prose) as string)).toBe(false);
  });

  it("keeps a real call", () => {
    const real = `// a comment\nconst v = Office.context.requirements;\n`;
    expect(OFFICE.test(withoutProse(real) as string)).toBe(true);
  });

  it("does not join identifiers when it removes a string", () => {
    expect((withoutProse(`const a = x + "mid" + y;`) as string).replace(/\s+/g, " ")).toBe(`const a = x + "" + y;`);
  });
});
