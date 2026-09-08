/**
 * The Office.js calls, and nothing else.
 *
 * Every judgement is imported from `src/host`, where it is a pure function the
 * suite can check. Nothing here decides anything, because a decision inside a
 * `PowerPoint.run` callback is a decision nobody can test — and the host this
 * will run on is documented, by both sibling projects and at length, to lie
 * about ids, to accept calls it does not perform, and to answer differently on
 * two runs of the same build. `CLAUDE.md` ("Host rules, learned the expensive way") is the ledger,
 * borrowed from the siblings.
 *
 * Today this file asks one question. The insert arrives with the host layer
 * (`docs/BACKLOG.md`), built from the two calls the siblings have proven
 * against a real PowerPoint: `getFileAsync(Compressed)` to read the deck and
 * `insertSlidesFromBase64` WITH a `targetSlideId`.
 */
import { checkFloor, type Readiness, type Supports } from "../host/capability.js";

/**
 * What the host says it supports.
 *
 * Answers false rather than throwing when it cannot ask: `Office.context` being
 * absent is not a supported host reporting an absent set, it is a host that
 * cannot answer — and every caller wants the same thing from both. The guard is
 * at the root because this is called before the pane has decided anything, and
 * a raise here leaves the user with a blank pane on precisely the host that
 * needed the explanation.
 */
export const hostSupports: Supports = (version) => {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
};

/** Whether this host clears the floor, as a sentence the pane can show. */
export function ready(): Readiness {
  return checkFloor(hostSupports);
}
