import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * What node this repo says it needs, against what its dependencies say.
 *
 * `package.json` declared `>=22` and nothing held it, so nothing noticed when
 * the dev-tooling bump of 2026-09-23 took `jsdom` to 30.1.1 — which, with
 * `@asamuzakjp/css-color`, `@asamuzakjp/dom-selector` and `w3c-xmlserializer`,
 * declares `^22.22.2 || ^24.15.0 || >=26.0.0`. On node 22.10 the install
 * succeeds (there is no `.npmrc`, so `engine-strict` is off and npm only
 * warns) and the declaration is simply false: the repo says it runs somewhere
 * it does not.
 *
 * Read from `package-lock.json`, which is COMMITTED and carries `engines` for
 * 166 of its packages, so this needs no `node_modules` and answers the same on
 * a fresh checkout as it does here. Every workflow takes its node from
 * `.nvmrc`, which is why the declaration could rot without CI noticing — CI
 * happens to run a node that satisfies it.
 *
 * The check is deliberately narrow: it does not try to intersect semver
 * ranges, it asserts that the repo's own floor is not the bare `>=22` that
 * three of its dependencies refuse, and that `.nvmrc`'s line is inside every
 * range the tree asks for. Anything cleverer would be a semver library in a
 * test.
 */
const lock = JSON.parse(readFileSync("package-lock.json", "utf8")) as {
  packages: Record<string, { engines?: { node?: string } }>;
};
const root = JSON.parse(readFileSync("package.json", "utf8")) as { engines?: { node?: string } };

/**
 * The lowest node-22 release a range admits, as a sortable number, or undefined
 * when it admits none.
 *
 * BOTH spellings, which the first version of this did not do: it matched
 * `^22.x.y` alone, and the tree carries `>=22.19.0` and `>=22.13.0 || >=24` as
 * well. Those are all below today's strictest, so the check happened to give
 * the right answer — by luck, not because it was looking. A future
 * `>=22.23.0` would have sailed past it, which is the shape `CLAUDE.md` calls
 * a gate that cannot fail.
 *
 * A range is a UNION, so its floor is the LOWEST of its alternatives: an
 * alternative naming no 22 at all cannot hold a node 22 developer up, and one
 * that does sets the bar for that branch. `^20.19.0 || >=22.12.0` therefore
 * floors at 22.12.0 — node 22.0 satisfies neither side.
 *
 * Anchored so the major is exactly 22: `>=12.22.7` and `^12.22.0` name no node
 * 22 and must not be read as though they did.
 */
function floor22(range: string): number | undefined {
  const floors: number[] = [];
  for (const alternative of range.split("||")) {
    const caret = /(?:^|\s)\^v?22(?:\.(\d+))?(?:\.(\d+))?/.exec(alternative);
    const atLeast = /(?:^|\s)>=\s*v?22(?:\.(\d+))?(?:\.(\d+))?/.exec(alternative);
    const bare = /^\s*v?22\s*$/.test(alternative);
    const hit = caret ?? atLeast;
    if (hit) floors.push(Number(hit[1] ?? 0) * 1000 + Number(hit[2] ?? 0));
    else if (bare) floors.push(0);
  }
  return floors.length === 0 ? undefined : Math.min(...floors);
}

describe("reading a node range's floor on the 22 line", () => {
  /**
   * The parsing, held directly, because the check above can only ever be as
   * good as this — and the first version of it matched `^22.x.y` alone while
   * the tree carries `>=22.19.0` and `>=22.13.0 || >=24` too. Those are all
   * BELOW today's strictest, so the check gave the right answer without
   * looking at them.
   *
   * Every string here is one this repository's own lockfile carries today,
   * plus the one that would have slipped through.
   */
  it("reads both spellings, and takes the lowest branch of a union", () => {
    expect(floor22("^22.22.2 || ^24.15.0 || >=26.0.0"), "the caret form, today's strictest").toBe(22002);
    expect(floor22(">=22.19.0"), "the >= form the first version missed entirely").toBe(19000);
    expect(floor22(">=22.13.0 || >=24"), "the >= form inside a union").toBe(13000);
    // A union is permissive, so its floor is its LOWEST branch: node 22.0
    // satisfies neither side of this one, and 22.12 satisfies the second.
    expect(floor22("^20.19.0 || >=22.12.0")).toBe(12000);
    expect(floor22(">=22"), "no minor named at all").toBe(0);
    expect(floor22("18 || 20 || >=22")).toBe(0);
  });

  it("does not read a 22 that is somebody else's minor", () => {
    // Anchored on the MAJOR. `>=12.22.7` names node 12, and reading it as 22
    // would invent a floor out of another major's patch number.
    expect(floor22(">=12.22")).toBeUndefined();
    expect(floor22(">=v12.22.7")).toBeUndefined();
    expect(floor22("^12.22.0 || ^14.17.0 || >=16.0.0")).toBeUndefined();
    expect(floor22("^18.0.0"), "a range naming no 22 cannot hold a node 22 developer up").toBeUndefined();
  });

  it("would catch a stricter dependency arriving in the form it used to miss", () => {
    // The whole point. A future `>=22.23.0` is ABOVE the declared floor, and
    // the first version of this function returned undefined for it — so the
    // guard would have passed a tree the repo could no longer run on.
    expect(floor22(">=22.23.0")).toBe(23000);
    expect(floor22(">=22.23.0")).toBeGreaterThan(floor22("^22.22.2 || ^24.15.0 || >=26.0.0") ?? 0);
  });
});

describe("the node this repo says it needs", () => {
  it("is not looser than what its own dependencies demand", () => {
    const demands = Object.entries(lock.packages)
      .filter(([name]) => name !== "")
      .map(([name, meta]) => [name, meta.engines?.node] as const)
      .filter((pair): pair is readonly [string, string] => typeof pair[1] === "string");
    expect(demands.length, "the lockfile carries no engines at all — has its shape changed?").toBeGreaterThan(50);

    const strictest = demands
      .map(([name, range]) => [name, floor22(range)] as const)
      .filter((pair): pair is readonly [string, number] => pair[1] !== undefined)
      .sort((a, b) => b[1] - a[1])[0];
    if (strictest === undefined) return;

    const declared = root.engines?.node ?? "";
    const ours = floor22(declared);
    expect(
      ours,
      `${strictest[0]} demands ${String(
        Object.entries(lock.packages).find(([n]) => n === strictest[0])?.[1].engines?.node,
      )} and package.json declares "${declared}", which admits node 22 releases it refuses`,
    ).not.toBeUndefined();
    expect(ours ?? -1).toBeGreaterThanOrEqual(strictest[1]);
  });

  it("is a range .nvmrc's own line can satisfy", () => {
    // `.nvmrc` is what every workflow resolves its node from. A floor the
    // pinned major cannot reach would make CI unable to run the repo at all.
    const nvmrc = readFileSync(".nvmrc", "utf8").trim();
    const declared = root.engines?.node ?? "";
    expect(declared, "package.json declares no node engine").not.toBe("");
    expect(declared, `.nvmrc pins node ${nvmrc}, which the declared range does not mention`).toContain(nvmrc);
  });
});
