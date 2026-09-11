import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import * as sweep from "../scripts/dead-exports.mjs";

/** One export the sweep reports, and how far it is reached from. */
interface Dead {
  file: string;
  name: string;
  kind: string;
  reach: "tests" | "nothing";
}

const deadExports = sweep.deadExports as (root?: string) => Dead[];
const ALLOWED = sweep.ALLOWED as Record<string, string>;
const keyOf = sweep.keyOf as (row: Dead) => string;

/**
 * Nothing in `src/` is written for its test alone.
 *
 * An export the product never calls is a comment that compiles: it type-checks,
 * its own test is green, and the shipped add-in runs none of it — so whatever
 * it was written to do is still not done. That shape has cost this family four
 * separate sessions, most recently a `shapesOf` the splice was documented to
 * adopt shapes by while the splice called `topLevel` directly, with a test that
 * compared the alias with the thing it aliased and could not fail.
 *
 * Remembering is what failed those four times, so this is a sweep.
 * `scripts/dead-exports.mjs` holds the rule and `ALLOWED` holds the deliberate
 * exceptions, each with its reason.
 */
describe("every export in src is reached by something that ships", () => {
  it("finds nothing the product never calls, apart from what is recorded", () => {
    const unexcused = deadExports().filter((row) => !(keyOf(row) in ALLOWED));
    expect(
      unexcused.map((row) => `${row.file} :: ${row.name} (reached by ${row.reach})`),
      "an export nothing in src/ or scripts/ calls. Wire it up, delete it, or record it in ALLOWED with the reason",
    ).toEqual([]);
  });

  it("keeps no excuse for an export that no longer exists", () => {
    // An allow-list is a second place a name lives, and a second place is a
    // place that goes stale. An entry naming a deleted export is not harmless:
    // it is a line of prose claiming a thing is deliberate when the thing is
    // gone, and the next person reads it as current.
    for (const key of Object.keys(ALLOWED)) {
      const [file, name] = key.split("::");
      const source = readFileSync(file as string, "utf8");
      expect(source, `${key} is recorded as a deliberate exception, but ${file} no longer exports it`).toMatch(
        new RegExp(`^export (?:async )?(?:function|const) ${name}\\b`, "m"),
      );
    }
  });

  it("proves the sweep reads calls rather than mentions", () => {
    /**
     * Two ways this sweep could be a gate that cannot fail, and both have
     * happened to it.
     *
     * A name written only in the paragraph explaining it is not a use — if
     * prose counted, every documented export would look alive and the sweep
     * would report nothing, ever. And a name written inside `${…}` IS a use:
     * the first version of this read `withoutTsProse`, which blanks a template
     * literal whole, and called `nameOfRatio` dead while the pane ran it on
     * every borrowed deck.
     *
     * `nameOfRatio` is reached from exactly one place, from inside a template
     * literal, and `carriedTypes` from the pane in the ordinary way. Neither
     * may be reported, and if the file below ever stops reaching them that way
     * this case has stopped proving anything — so it checks the shape of the
     * reach as well as the verdict.
     */
    const catalogue = readFileSync("src/pane/catalogue.ts", "utf8");
    expect(catalogue, "nameOfRatio is no longer called from inside a template literal").toContain(
      "${nameOfRatio(ratio)}",
    );
    const reported = deadExports().map((row) => keyOf(row));
    expect(reported).not.toContain("src/pane/catalogue.ts::nameOfRatio");
    expect(reported).not.toContain("src/pane/catalogue.ts::carriedTypes");
  });
});
