/**
 * What an undo may take back, and what it may not.
 *
 * An insert here REPLACES a slide: the spliced copy goes in after the original
 * and the original is removed. Undo therefore has to put a slide back that no
 * longer exists, which it cannot — so what this actually does is remove the
 * slide the insert ADDED, leaving the user where they started minus their
 * original slide's identity. That distinction is why the pane says "remove the
 * element" rather than "undo".
 *
 * The rules are the sibling's, learned against a real host that accepts calls
 * it does not perform:
 *
 * - **Positional, never by id.** Slide ids on this host look like
 *   `256#3561048925` and a subset export handed `["4", "5", "6"]` was accepted
 *   and did the wrong thing. A position is checked against a count.
 * - **Clamped.** A deck that changed under the add-in — the user deleted
 *   slides while the pane was open — must not have an unrelated slide removed
 *   because an index went stale.
 * - **The deck DELTA is the evidence, never the absence of an error.** An
 *   insert has timed out having landed everything it was asked for.
 */

export interface Sweep {
  /** Zero-based positions to remove, ascending. Empty when it is not safe. */
  positions: number[];
  /** Why, in a sentence the pane can show. */
  why: string;
}

/**
 * Which slide an insert's own output occupies, given what the deck looked like
 * before and after.
 *
 * `expected` is how many slides the insert should have added — always one here,
 * passed rather than assumed so a caller cannot silently disagree.
 */
export function sweepPlan(before: number, after: number, insertedAfterIndex: number, expected = 1): Sweep {
  const gained = after - before;
  if (gained <= 0) {
    return { positions: [], why: "nothing was added, so there is nothing to take back" };
  }
  if (gained !== expected) {
    // Not a refusal to act on principle: a deck that gained a different number
    // of slides than the insert asked for has had something else happen to it,
    // and removing by position would be removing somebody's work.
    return {
      positions: [],
      why: `the deck gained ${gained} slide${gained === 1 ? "" : "s"} where ${expected} was expected, so which one to remove is not certain`,
    };
  }
  const at = insertedAfterIndex + 1;
  if (at < 0 || at >= after) {
    return { positions: [], why: "the slide the insert added is no longer where it was put" };
  }
  return { positions: [at], why: `removing the slide added after position ${insertedAfterIndex + 1}` };
}

/** A queued delete that raised nothing has not necessarily happened. */
export function removalProven(before: number, after: number, asked: number): boolean {
  return before - after === asked;
}
