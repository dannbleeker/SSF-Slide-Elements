import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the harvest script.
import { publish, stage } from "../scripts/catalogue-out.mjs";

/**
 * `npm run harvest` used to open by deleting `public/catalogue` outright,
 * before a deck had been read — and it has two early exits after that point,
 * both of which then left the tree with the COMMITTED `catalogue.json` gone and
 * nothing in its place.
 *
 * The committed index is not a build artifact: five test files read it,
 * `npm run previews` reads it, and `.gitignore` covers only the two per-size
 * directories. So an owner who added a slide, ran the harvest and got the
 * ordinary "has no English name in the names file" error was then looking at a
 * suite failing for a reason unrelated to their deck, and at a staged deletion
 * of the index. `git checkout public/catalogue/catalogue.json` is the recovery
 * and was written down nowhere.
 *
 * The staging directory is the fix, and these hold both halves of it.
 */
const made: string[] = [];
const temp = (): string => {
  const dir = mkdtempSync(join(tmpdir(), "ssf-harvest-"));
  made.push(dir);
  return dir;
};
afterEach(() => {
  for (const dir of made.splice(0)) rmSync(dir, { recursive: true, force: true });
});

describe("where the harvest writes", () => {
  it("clears what an earlier run left in the staging directory", () => {
    // Merged into rather than cleared, a stale element file from a run against
    // an older deck would be published beside this run's with the index not
    // naming it — the same shape of bug, moved one step along.
    const root = temp();
    const work = join(root, "staging");
    mkdirSync(join(work, "16x9", "elements"), { recursive: true });
    writeFileSync(join(work, "16x9", "elements", "gone.json"), "{}");

    stage(work);

    expect(existsSync(join(work, "16x9", "elements", "gone.json")), "a stale file survived into this run").toBe(false);
    expect(existsSync(work), "the staging directory itself is gone").toBe(true);
  });

  it("leaves the committed catalogue alone when the run does not reach the end", () => {
    // The whole point. Everything before `publish` is a write into the staging
    // directory, so an exit anywhere in the harvest leaves the tree as it was.
    const root = temp();
    const out = join(root, "catalogue");
    const work = join(root, "staging");
    mkdirSync(out, { recursive: true });
    writeFileSync(join(out, "catalogue.json"), '{"committed":true}');

    stage(work);
    // ...and the run exits here, the way a HarvestError or a key mismatch does.

    expect(JSON.parse(readFileSync(join(out, "catalogue.json"), "utf8"))).toEqual({ committed: true });
  });

  it("refuses to publish a staging directory with no index in it", () => {
    // The one file every reader needs. Publishing without it would put the tree
    // into exactly the state this module exists to prevent, so it raises rather
    // than answering a flag a caller could ignore.
    const root = temp();
    const out = join(root, "catalogue");
    const work = join(root, "staging");
    mkdirSync(out, { recursive: true });
    writeFileSync(join(out, "catalogue.json"), '{"committed":true}');
    stage(work);
    writeFileSync(join(work, "something-else.json"), "{}");

    expect(() => (publish as (a: string, b: string) => string)(work, out)).toThrow(/no catalogue\.json/);
    expect(JSON.parse(readFileSync(join(out, "catalogue.json"), "utf8")), "it published anyway").toEqual({
      committed: true,
    });
  });

  it("replaces the committed catalogue once there is a whole one to replace it with", () => {
    const root = temp();
    const out = join(root, "catalogue");
    const work = join(root, "staging");
    mkdirSync(out, { recursive: true });
    writeFileSync(join(out, "catalogue.json"), '{"committed":true}');
    // A file the new run does NOT produce, to prove the old tree is replaced
    // rather than written over: a removed element must not survive a harvest.
    mkdirSync(join(out, "16x9", "elements"), { recursive: true });
    writeFileSync(join(out, "16x9", "elements", "removed.json"), "{}");

    stage(work);
    writeFileSync(join(work, "catalogue.json"), '{"fresh":true}');

    expect(publish(work, out)).toBe(out);
    expect(JSON.parse(readFileSync(join(out, "catalogue.json"), "utf8"))).toEqual({ fresh: true });
    expect(existsSync(join(out, "16x9", "elements", "removed.json")), "an element the new run dropped came back").toBe(
      false,
    );
    expect(existsSync(work), "the staging directory was left behind").toBe(false);
  });
});
