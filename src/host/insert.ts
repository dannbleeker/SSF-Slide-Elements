/**
 * What an insert DID, read from the deck's size rather than from what the host
 * said about it.
 *
 * `CLAUDE.md` has the two rules this file exists for, and they pull in opposite
 * directions: **a queued call that raises nothing has not necessarily
 * happened**, and **a call can raise and still have done the work** — SSF-Merge
 * measured an insert that timed out with both its slides landed. So neither the
 * absence of an error nor its presence decides anything here. The deck is
 * counted before, after the insert, and after the removal, and the DELTA is the
 * evidence.
 *
 * Every sentence below is the one `docs/DESIGN.md` sections 6 and 10 specify,
 * word for word, because the pane's wording is design rather than
 * implementation. A test holds them to the document.
 */

/** Which of the two things the user asked for. `docs/DESIGN.md` section 7. */
export type Target = "onto" | "new";

/** The counts an insert produced, each measured in its own read. */
export interface Attempt {
  target: Target;
  /**
   * The slide the user was on, counting from ONE, because it goes into a
   * sentence a user reads.
   *
   * `undoPlan` below counts the same slide from ZERO, because its numbers go
   * into API calls. Two bases in one file is a trap, and it is the trap the
   * undo defect fell into, so each is named at the field rather than left to
   * the reader.
   */
  slide: number;
  /** Slides in the deck before anything was asked of the host. */
  before: number;
  /** Slides after `insertSlidesFromBase64`. */
  inserted: number;
  /**
   * Slides after the replaced slide was removed by position.
   *
   * Absent for a new slide, where nothing is removed — and absent, too, when
   * the insert did not land, because the removal is not attempted then. A
   * reader must not treat "absent" as "zero".
   */
  removed?: number;
  /** Whatever the host raised, already made readable and bounded by `errors.ts`. */
  error?: string;
}

export interface Outcome {
  /** Whether the deck now holds what the user asked for. */
  ok: boolean;
  /** The footer's sentence. */
  detail: string;
  /**
   * Whether the deck is in a state only the user can put right.
   *
   * One case reaches this: the insert landed and the removal did not, so there
   * is a slide too many. The pane must say which one rather than trying again,
   * because a second attempt at a positional delete on a deck whose shape it
   * has already misread is how the wrong slide goes.
   */
  byHand: boolean;
}

/**
 * The measured delta, as the sentence the footer shows.
 *
 * The shape of the sentence is `docs/DESIGN.md` section 6: "the footer reports
 * the deck's slide count before and after, `12 → 13 slides` for a new slide,
 * `12 → 13 → 12 slides, slide 4 replaced` for an insert onto the slide, because
 * in the file that is insert a rebuilt copy after the slide and remove the
 * original".
 */
export function outcomeOf(attempt: Attempt): Outcome {
  const { before, inserted, target, slide } = attempt;
  const landed = inserted - before;

  if (landed <= 0) {
    // Nothing arrived. Which of the two sentences depends on whether the host
    // said anything, and BOTH are honest: a refusal names the reason, and a
    // silent no-op says plainly that the deck is untouched, because "it did
    // not work" without a count leaves the user wondering what to undo.
    if (attempt.error !== undefined) {
      return { ok: false, detail: `The insert was refused: ${attempt.error}`, byHand: false };
    }
    return {
      ok: false,
      detail: `The insert did not confirm: the deck still has ${before} slides, nothing was changed.`,
      byHand: false,
    };
  }

  if (landed > 1) {
    // More than the one slide the package listed. Nothing in the design
    // predicts it, so it is named rather than folded into a failure — and the
    // deck is left exactly as it is, because a sweep over slides this code
    // cannot account for is how the wrong one goes.
    return {
      ok: false,
      detail: `The deck grew by ${landed} slides where one was expected: ${before} → ${inserted} slides. Nothing was removed; check the end of the deck.`,
      byHand: true,
    };
  }

  if (target === "new") {
    return { ok: true, detail: `${before} → ${inserted} slides.`, byHand: false };
  }

  const removed = attempt.removed;
  if (removed === undefined || removed !== before) {
    return {
      ok: false,
      detail: `The deck grew by one but the copy could not be removed: delete slide ${slide} by hand.`,
      byHand: true,
    };
  }
  return {
    ok: true,
    detail: `${before} → ${inserted} → ${removed} slides, slide ${slide} replaced.`,
    byHand: false,
  };
}

/**
 * Whether a positional delete may be attempted at all.
 *
 * The insert is confirmed first, and this is that rule as a function rather
 * than an `if` inside a callback nobody can test. `CLAUDE.md`: undo and removal
 * are **positional with clamps, never by id**, because a slide the run just
 * added does not resolve by id on the web — so the index has to be one this
 * code computed and can defend, and the deck has to be the size that index was
 * computed against.
 */
export function mayRemove(attempt: Pick<Attempt, "before" | "inserted">): boolean {
  return attempt.inserted === attempt.before + 1;
}

/** Which two slides an undo has to touch, and in which order. */
export interface UndoPlan {
  /**
   * The index of the slide to insert the restored copy AFTER, or undefined when
   * the undo only removes something.
   */
  after?: number;
  /** The index to remove, once the deck is the size the plan expects. */
  remove: number;
  /** What the deck should hold after the insert half, when there is one. */
  grownTo: (before: number) => number;
}

/**
 * What taking back an insert actually means, as indices.
 *
 * This was three lines inside a callback and it was WRONG, in the way that is
 * worst: it reported success and changed nothing. Undoing an insert that landed
 * onto slide N means putting the user's original slide N back and taking the
 * rebuilt one away — and the rebuilt one is at index N, so the restored copy
 * lands at N+1 and it is **N** that must go. The code removed N+1, which is the
 * copy it had just restored, so the deck came back to its old size, the count
 * check passed, and the pane said "Undone" over a slide that had not moved.
 * Caught by pressing Undo in PowerPoint and looking at the slide.
 *
 * The insert is aimed at the REBUILT slide rather than at whatever the user has
 * selected, which is the other half of the same lesson: an undo that depends on
 * where the selection happens to be is an undo that puts a slide somewhere else
 * the moment the user clicks away before pressing it.
 *
 * A new slide is the simple case. It sits immediately after the slide it was
 * inserted against, and taking it back is one removal.
 *
 * `index` counts from ZERO, unlike `Attempt.slide` above, which counts from one
 * because it is read aloud. The field is named for its base rather than for the
 * thing it points at, so a call site cannot quietly hand over the other one.
 */
export function undoPlan(entry: { target: Target; index: number }): UndoPlan {
  if (entry.target === "new") {
    return { remove: entry.index + 1, grownTo: (before) => before };
  }
  return { after: entry.index, remove: entry.index, grownTo: (before) => before + 1 };
}

/**
 * What the pane announces while an insert runs.
 *
 * `docs/DESIGN.md` section 6: one insert at a time, the pane locks, and the
 * tile says "Inserting…". Two inserts 0.4 s apart killed a sibling's tab; the
 * lock is that rule made visible.
 */
export const INSERTING = "Inserting…";

/** The live region's sentence for an outcome, which is the footer's without the arrows. */
export function announcement(outcome: Outcome, name: string): string {
  return outcome.ok ? `${name} inserted. ${outcome.detail}` : outcome.detail;
}
