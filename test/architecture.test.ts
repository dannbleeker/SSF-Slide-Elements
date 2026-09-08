import { readdirSync, readFileSync, statSync } from "node:fs";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { MEASURED, NOT_MEASURED } from "../scripts/coverage-scope.mjs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutTsProse } from "../scripts/without-prose.mjs";

function filesUnder(dir: string): string[] {
  return readdirSync(dir).flatMap((entry) => {
    const path = join(dir, entry);
    return statSync(path).isDirectory() ? filesUnder(path) : path.endsWith(".ts") ? [path] : [];
  });
}

/**
 * The file with its prose removed, so a guard reads what the code DOES.
 *
 * A sibling's first version of this matched the words "Office.js" and
 * "PowerPoint.run" in the comments that explain WHY the engine avoids them, so
 * it failed on four files that were entirely correct. A guard that goes red
 * for the wrong reason teaches the next reader to widen it until it goes green.
 */
function codeOf(file: string): string {
  return withoutTsProse(readFileSync(file, "utf8")) as string;
}

/** The file as written, for a rule whose subject is a string literal. */
function rawOf(file: string): string {
  return readFileSync(file, "utf8");
}

/**
 * An import of the host's own typings or shim, in any spelling.
 *
 * Read against the RAW source, so it must not match the word where it appears
 * in prose — hence the `import`/`require`/`from` context rather than the bare
 * package name.
 */
const IMPORTS_OFFICE = /(?:^|\n)\s*(?:import\b[^\n]*|const[^\n]*=\s*require\s*\()\s*["'`][^"'`\n]*office-js/;

const HOST_GLOBAL = /\bOffice\.|\bPowerPoint\./;

describe("the layering, held in both directions", () => {
  it("keeps Office.js out of src/core", () => {
    // The seam the whole design rests on. The engine takes bytes and answers
    // bytes, which is what lets it run in the pane, in a script and in this
    // suite with no PowerPoint anywhere. TWO reads, because the two halves of
    // this rule live in different places: an identifier is code and must be
    // read with the prose stripped; an IMPORT SPECIFIER is a string literal,
    // which `withoutTsProse` blanks, so it is read raw.
    const files = filesUnder("src/core");
    expect(files.length, "src/core has no source to hold").toBeGreaterThan(0);
    const usesHost = files.filter((f) => HOST_GLOBAL.test(codeOf(f)));
    const importsHost = files.filter((f) => IMPORTS_OFFICE.test(rawOf(f)));
    expect([...new Set([...usesHost, ...importsHost])]).toEqual([]);
  });

  it("keeps Office.js out of src/host too, so every decision stays testable", () => {
    // src/host is the DECISIONS: which version floor the host clears, what a
    // raise means on screen. src/office is the calls. One Office.js import
    // here and the same rules become untestable again.
    const files = filesUnder("src/host");
    expect(files.length, "src/host has no source to hold").toBeGreaterThan(0);
    const usesHost = files.filter((f) => HOST_GLOBAL.test(codeOf(f)));
    const importsHost = files.filter((f) => IMPORTS_OFFICE.test(rawOf(f)));
    expect([...new Set([...usesHost, ...importsHost])]).toEqual([]);
  });

  it("keeps the decisions out of src/office, so none of them hide in a callback", () => {
    // The other direction, and the one that rots quietly: a rule reimplemented
    // inline next to the call it guards looks tidier and is untestable. Every
    // file here imports its judgements from src/host or src/core, or it has
    // started deciding for itself.
    const office = filesUnder("src/office");
    expect(office.length).toBeGreaterThan(0);
    for (const file of office) {
      // The RAW source: an import specifier IS a string literal, and codeOf
      // strips those so the Office.js check above cannot trip over prose.
      expect(rawOf(file), `${file} decides nothing on its own`).toMatch(/from "\.\.\/(host|core)\//);
    }
  });

  it("lets only the pane's entry point touch Office.js", () => {
    // steps.ts and render.ts are the pane's decisions and its screen, both
    // checked in the suite without a PowerPoint. The moment one of them reads
    // Office.context the labels a user acts on become untestable.
    const offenders = filesUnder("src/pane")
      .filter((f) => !f.endsWith("main.ts"))
      .filter((f) => HOST_GLOBAL.test(codeOf(f)));
    expect(offenders).toEqual([]);
  });

  it("proves the guard reads code, not prose", () => {
    // A gate that cannot fail is decoration. `withoutTsProse` is the whole
    // reason the checks above are green on files whose comments name the
    // host, so it has to be shown removing the right thing and keeping the
    // right thing.
    const prose = `// hand it back to PowerPoint.\n/* Office echoes it back. */\nconst s = "PowerPoint.run";\n`;
    expect(HOST_GLOBAL.test(withoutTsProse(prose) as string)).toBe(false);
    const real = `// a comment\nconst v = Office.context.requirements;\n`;
    expect(HOST_GLOBAL.test(withoutTsProse(real) as string)).toBe(true);
  });
});

describe("what coverage measures", () => {
  /**
   * The coverage `include` is a list of directories, and a new top-level
   * directory under `src/` would be measured by nothing — an uncounted module
   * is how a threshold quietly stops meaning anything. So every directory is
   * either measured or named with the reason it is not.
   */
  const dirs = readdirSync("src").filter((d) => statSync(join("src", d)).isDirectory());
  const measured = MEASURED as string[];
  const notMeasured = NOT_MEASURED as Record<string, string>;

  it("names every directory under src, one way or the other", () => {
    expect(dirs.length).toBeGreaterThan(2);
    for (const dir of dirs) {
      const said = measured.includes(dir) || Object.hasOwn(notMeasured, dir);
      expect(said, `src/${dir} is neither measured nor excused in scripts/coverage-scope.mjs`).toBe(true);
    }
  });

  it("measures only directories that exist", () => {
    for (const dir of measured) expect(dirs, `MEASURED names src/${dir}, which is not there`).toContain(dir);
    for (const dir of Object.keys(notMeasured)) {
      expect(dirs, `NOT_MEASURED names src/${dir}, which is not there`).toContain(dir);
      expect(notMeasured[dir]?.length ?? 0, `src/${dir} is excused with no reason`).toBeGreaterThan(10);
    }
  });

  it("does not measure the Office.js calls", () => {
    expect(measured).not.toContain("office");
    expect(Object.keys(notMeasured)).toContain("office");
  });
});

describe("one resolver for relationship targets", () => {
  /**
   * A relationship target is resolved against the part that owns it, honouring
   * a leading `/` and any number of `..`. That is fiddly enough to get subtly
   * wrong, and in SSF-Merge it WAS: three files carried their own copy, and
   * when the root-part case was fixed on 2026-08-29 only one of the three got
   * the fix. The other two were latent rather than broken, which is exactly
   * how a copy survives: it costs nothing until it does.
   *
   * Anchored on the one line no other function has, and deliberately on a
   * fragment carrying NO string literal: `codeOf` blanks every literal so a
   * guard cannot match prose. A guard that finds zero offenders looks exactly
   * like a guard that is satisfied, so this names the one carrier it expects
   * and goes red if the anchor moves rather than quiet.
   */
  const SIGNATURE = "ownerPart.slice(0, slash)";

  it("lives in exactly one file", () => {
    const carriers = filesUnder("src")
      .filter((f) => codeOf(f).includes(SIGNATURE))
      .map((f) => f.replaceAll("\\", "/"));
    expect(carriers, "a second copy of resolveTarget has appeared").toEqual(["src/core/pptx/pkg.ts"]);
  });
});

describe("what a slide's relationships are called", () => {
  /**
   * These strings decide which parts a clone copies, which it drops, and which
   * a removal deletes. SSF-Merge had them written out in six files, three of
   * them twice; the copies agreed and nothing had gone wrong, but one copy of
   * the comment list said what a clone drops and the other what a removal
   * deletes. PowerPoint has already added a second spelling of comments once,
   * and adding a third to one copy and not the other leaves a clone carrying a
   * comment part the removal will not clean up. So `src/core/pptx/parts.ts` is
   * the one place a relationship type is spelled out.
   */
  it("is spelled out in one file", () => {
    // RAW source, because `codeOf` masks string literals and these ARE string
    // literals — the check would pass by finding nothing.
    const offenders = filesUnder("src")
      .filter((f) => !f.endsWith("parts.ts"))
      .filter((f) => /["'`][^"'`]*\/relationships\/[a-zA-Z]/.test(rawOf(f)));
    expect(offenders, "writes a relationship type out instead of naming it").toEqual([]);
  });
});
