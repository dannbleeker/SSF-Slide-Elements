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
/** Where the committed catalogue is parked for the moment the swap takes. */
export const ASIDE = "public/.catalogue-previous";

/**
 * An empty staging directory, whatever a previous run left behind.
 *
 * Cleared rather than merged into: a stale element file from a run against an
 * older deck would otherwise be published beside this run's, and the index
 * would not name it — which is exactly the shape of the bug the staging
 * directory exists to prevent, moved one step along.
 */
export function stage(root = STAGE, out = OUT, aside = ASIDE) {
  // RECOVER FIRST. `publish` moves the committed catalogue aside and then moves
  // the staged one into place; a run killed between those two renames leaves
  // `out` missing and everything it held under `aside`. Putting it back is the
  // whole of the repair, and doing it here means the next harvest does it
  // without anybody knowing the aside directory exists.
  if (!existsSync(out) && existsSync(aside)) renameSync(aside, out);
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
export function publish(root = STAGE, out = OUT, aside = ASIDE) {
  if (!existsSync(join(root, "catalogue.json"))) {
    throw new Error(`refusing to publish ${root} over ${out}: it has no catalogue.json`);
  }
  // TWO RENAMES, then the delete — not a delete and then a rename.
  //
  // The first version of this did `rmSync(out, {recursive: true})` and then
  // renamed the staged tree in. That is a walk over 427 files, and a run
  // interrupted inside it — Ctrl-C, a crash, a full disk — left the committed
  // catalogue deleted with no replacement: exactly the state this module exists
  // to prevent, moved from an early exit to the publish itself.
  //
  // A rename is one metadata operation. Parking the old tree under `aside`
  // first and moving the staged one in second narrows the window where nothing
  // is at `out` to the gap between two of them, and `stage` puts `aside` back
  // if a run ever dies in it. The slow delete happens last, when it can cost
  // nothing but disk.
  rmSync(aside, { recursive: true, force: true });
  if (existsSync(out)) renameSync(out, aside);
  renameSync(root, out);
  rmSync(aside, { recursive: true, force: true });
  return out;
}
