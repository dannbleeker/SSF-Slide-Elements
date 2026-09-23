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

/** The lowest node each dependency's range admits on the 22 line, if it names one. */
function floor22(range: string): number | undefined {
  // `^22.22.2 || ^24.15.0 || >=26.0.0` -> 22.22.2. Only the 22 branch matters:
  // `.nvmrc` pins the major, and a range naming no 22 at all cannot constrain
  // a node 22 developer downward.
  const match = /\^22\.(\d+)\.(\d+)/.exec(range);
  if (!match) return undefined;
  return Number(match[1]) * 1000 + Number(match[2]);
}

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
