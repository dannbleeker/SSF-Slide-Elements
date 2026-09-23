/**
 * Where `npm run harvest` writes, and when the committed catalogue is replaced.
 *
 * The harvest used to open with `rmSync("public/catalogue", {recursive: true,
 * force: true})`, before a single deck had been read. Two exits follow it — a
 * `HarvestError` (an element with no English name in `template/names.en.json`
 * is the ordinary one, since that file is hand-kept) and the two decks not
 * carrying the same keys — and both left the working tree with the COMMITTED
 * `public/catalogue/catalogue.json` deleted and nothing in its place.
 *
 * What that costs is not the harvest. `catalogue.json` is a committed file that
 * five test files read, `npm run previews` reads it too, and `.gitignore` covers
 * only `public/catalogue/16x9/` and `public/catalogue/4x3/` — so the owner who
 * added a slide and ran the harvest got a deck error, and then a suite failing
 * for a reason that has nothing to do with the deck, and a staged deletion of
 * the index waiting for a distracted `git add -A`. Recovery is
 * `git checkout public/catalogue/catalogue.json`, which was written down
 * nowhere. CI was never harmed: it runs on an ephemeral checkout and the exit
 * fails the step before the diff, which is the correct outcome.
 *
 * So the harvest writes to a staging directory and the committed one is
 * replaced only once everything has succeeded. A run that exits early leaves
 * the tree exactly as it found it, plus a staging directory that the next run
 * clears.
 */
import { existsSync, mkdirSync, rmSync, renameSync } from "node:fs";
import { join } from "node:path";

/** The committed catalogue, and the directory a run builds its replacement in. */
export const OUT = "public/catalogue";
export const STAGE = "public/.catalogue-staging";

/**
 * An empty staging directory, whatever a previous run left behind.
 *
 * Cleared rather than merged into: a stale element file from a run against an
 * older deck would otherwise be published beside this run's, and the index
 * would not name it — which is exactly the shape of the bug the staging
 * directory exists to prevent, moved one step along.
 */
export function stage(root = STAGE) {
  rmSync(root, { recursive: true, force: true });
  mkdirSync(root, { recursive: true });
  return root;
}

/**
 * The staged catalogue in place of the committed one.
 *
 * Refuses a staging directory with no `catalogue.json` in it, because that is
 * the one file every reader needs and publishing without it would leave the
 * tree in the state this module exists to prevent. Raises rather than
 * returning a flag: there is nothing sensible for a caller to do with the
 * answer except stop.
 */
export function publish(root = STAGE, out = OUT) {
  if (!existsSync(join(root, "catalogue.json"))) {
    throw new Error(`refusing to publish ${root} over ${out}: it has no catalogue.json`);
  }
  rmSync(out, { recursive: true, force: true });
  renameSync(root, out);
  return out;
}
