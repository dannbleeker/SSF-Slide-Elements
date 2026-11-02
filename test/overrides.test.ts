import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * Every `overrides` entry in `package.json` must still be doing something.
 *
 * An override forces a transitive dependency to a version its parents have
 * not declared, which is a claim about compatibility the parents did not
 * make. The one written on 2026-11-02 forces `minimatch` to 11 under `eslint`,
 * `@eslint/config-array` and `@typescript-eslint/typescript-estree`, all of
 * which still declare `^10.2.x`, because the only fix for GHSA-g3hv-75wg-m249
 * is a major and none of the three had taken it (`docs/DEPENDENCY-ALERTS.md`).
 *
 * npm never tells anyone when an override stops being needed. Once every
 * parent admits the version on its own, the override is dead weight that
 * keeps pinning a major by hand, and the day a parent moves to 12 it silently
 * holds the tree back. So this reads the COMMITTED lockfile — which carries
 * every package's declared `dependencies` — and fails the moment no parent
 * still caps the overridden package below the override: that is the signal
 * to delete the entry, not to keep it.
 *
 * Deliberately narrow: it compares caret MAJORS, the only shape this tree's
 * declarations take, rather than carrying a semver library. An override whose
 * range or a parent's range is not a plain `^M.m.p` fails loudly so the
 * comparison is widened on purpose.
 */
type Lock = { packages: Record<string, { version?: string; dependencies?: Record<string, string> }> };
const lock = JSON.parse(readFileSync("package-lock.json", "utf8")) as Lock;
const root = JSON.parse(readFileSync("package.json", "utf8")) as { overrides?: Record<string, string> };

/** The major of a plain caret range, or undefined for any other shape. */
export function caretMajor(range: string): number | undefined {
  const hit = /^\^(\d+)\.\d+\.\d+$/.exec(range.trim());
  return hit ? Number(hit[1]) : undefined;
}

/** Every package in the lockfile that declares `name`, with the range it declares. */
export function declarersOf(packages: Lock["packages"], name: string): Array<readonly [string, string]> {
  return Object.entries(packages)
    .filter(([path]) => path !== "")
    .flatMap(([path, meta]) => {
      const range = meta.dependencies?.[name];
      return typeof range === "string" ? [[path.replace(/^node_modules\//, ""), range] as const] : [];
    });
}

/**
 * Whether an override is still forcing something: true while at least one
 * declarer caps the package at a major below the override's.
 */
export function stillNeeded(overrideRange: string, declared: ReadonlyArray<readonly [string, string]>): boolean {
  const target = caretMajor(overrideRange);
  if (target === undefined) throw new Error(`override range "${overrideRange}" is not a plain caret; widen the check`);
  return declared.some(([name, range]) => {
    const major = caretMajor(range);
    if (major === undefined)
      throw new Error(`${name} declares "${range}", which is not a plain caret; widen the check`);
    return major < target;
  });
}

describe("reading a caret range's major", () => {
  it("reads the shapes this lockfile carries and refuses the rest", () => {
    expect(caretMajor("^10.2.4")).toBe(10);
    expect(caretMajor("^11.0.1")).toBe(11);
    expect(caretMajor(">=10.2.4"), "not a caret").toBeUndefined();
    expect(caretMajor("^10.2"), "not a full version").toBeUndefined();
    expect(caretMajor("10.2.4"), "pinned, not a caret").toBeUndefined();
  });

  it("knows a dead override from a live one", () => {
    expect(stillNeeded("^11.0.1", [["eslint", "^10.2.5"]])).toBe(true);
    expect(stillNeeded("^11.0.1", [["eslint", "^11.0.0"]]), "the parent caught up").toBe(false);
    expect(stillNeeded("^11.0.1", []), "nobody declares it at all").toBe(false);
    expect(() => stillNeeded("~11.0.1", [])).toThrow(/not a plain caret/);
  });
});

describe("the overrides in package.json", () => {
  it("are each still forcing a version some parent would not take on its own", () => {
    const overrides = Object.entries(root.overrides ?? {});
    for (const [name, range] of overrides) {
      const declared = declarersOf(lock.packages, name);
      expect(declared.length, `${name} is overridden but nothing in the lockfile depends on it`).toBeGreaterThan(0);
      expect(
        stillNeeded(range, declared),
        `every parent of ${name} now admits ${range} on its own (${declared
          .map(([p, r]) => `${p} → ${r}`)
          .join(", ")}); the override is dead — delete it from package.json and let the parents' own ranges resolve`,
      ).toBe(true);
    }
  });

  it("have landed: the lockfile resolves each overridden package inside the override", () => {
    for (const [name, range] of Object.entries(root.overrides ?? {})) {
      const target = caretMajor(range);
      const resolved = lock.packages[`node_modules/${name}`]?.version ?? "";
      expect(Number(resolved.split(".")[0]), `${name} resolved to ${resolved}, outside the override ${range}`).toBe(
        target,
      );
    }
  });
});
