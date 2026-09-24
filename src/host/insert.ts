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

import { sameSlideId, slideIdSuffix } from "./jump.js";

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
  /**
   * True when the removal was NOT attempted because the deck moved under it —
   * the slide at the computed index is no longer the one the insert was aimed
   * at. `removed` is absent then, like every other case where nothing was
   * removed, and this says WHY, because the two want different sentences: one
   * asks the user to delete a slide, and this one must not name a number at
   * all.
   */
  moved?: boolean;
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
   * Three cases reach it. The insert landed and the removal did not, so there
   * is a slide too many — the pane must say which one rather than trying again,
   * because a second attempt at a positional delete on a deck whose shape it
   * has already misread is how the wrong slide goes. The deck grew by more than
   * the one slide the package listed. And the deck SHRANK, which no step here
   * can do and which the pane therefore cannot describe further than the two
   * counts it took.
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
    // The SHRINK is asked about first, because it is the more serious of the
    // two facts and the raise does not make it untrue. This tested `error`
    // first, so an insert that both raised and left the deck smaller came out
    // as "The insert was refused: …" with `byHand: false` — the pane's mildest
    // sentence over a deck that had lost a slide, with nothing telling the user
    // to look. Both halves are reachable in one go: `insertPackage` answers a
    // reason instead of throwing, and `countReaching` then answers whatever it
    // last saw, so an insert that ran out of budget while the user deleted a
    // slide in the same window supplies exactly this.
    if (landed < 0) {
      // A deck that SHRANK. This used to fall into the sentence below, which
      // then stated a count the deck does not have — "the deck still has 12
      // slides" over a deck holding 11 — and claimed nothing had changed from a
      // delta that is itself the evidence something did. Both halves false, in
      // the one function whose whole purpose is to never say more than the
      // count supports, and a test pinned it that way.
      //
      // Reachable without any host misbehaving: the pane locks ITSELF, not
      // PowerPoint, and on the web an insert plus its confirming count takes
      // seconds (`CLAUDE.md`'s 2.8 second lag), so a user deleting a slide in
      // that window produces exactly this. What the pane cannot know is whether
      // the insert also landed, so it says what it measured and stops.
      return {
        ok: false,
        detail:
          attempt.error === undefined
            ? `The insert did not confirm: the deck has ${inserted} slides where it had ${before}. Check the deck before inserting again.`
            : `The insert was refused: ${attempt.error} — and the deck has ${inserted} slides where it had ${before}. Check the deck before inserting again.`,
        byHand: true,
      };
    }
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
  if (removed !== undefined && removed < before) {
    // The removal took the copy AND the deck lost something else. The sentence
    // below is written for `removed === before + 1`, where the copy is still
    // there — and it fired here too, where the copy is gone and slide `slide`
    // is the one the element LANDED on (`landedOn` for "onto" is the same
    // number). A user who followed it deleted their own content, which is the
    // exact failure `src/pane/main.ts` says this file exists to prevent.
    return {
      ok: false,
      detail: `The insert landed, but the deck now has ${removed} slides where it had ${before}. Check the deck before inserting again.`,
      byHand: true,
    };
  }
  if (attempt.moved === true) {
    // The deck was reordered while the insert ran, so the index this code
    // computed no longer names the slide it computed it for and the removal was
    // not attempted. It deliberately names NO slide number: the positions this
    // function was given are the ones that just went stale, and the sentence
    // below — which does name one — is what sent a user to delete their own
    // content the last time this file guessed.
    return {
      ok: false,
      detail:
        `The insert landed, but the deck was reordered while it ran, so the copy was left in place: ` +
        `${before} → ${inserted} slides. Both your slide and the copy are there; delete whichever you do not want.`,
      byHand: true,
    };
  }
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
 * Whether the slide at the index the insert computed is still the slide it was
 * computed FOR.
 *
 * `mayRemove` below asks only about the COUNT, and a count cannot see a
 * reorder. The index is read before the host calls and used after them, with
 * the whole splice and an `insertSlidesFromBase64` in between — up to
 * `BUDGET.insert`, which is 180 seconds. The pane locks ITSELF, not PowerPoint,
 * so the user can drag a slide in the thumbnail strip in that window, and
 * dragging changes no count at all.
 *
 * The two halves of the cycle then disagree about which slide they mean:
 * `insertPackage` aims by `targetSlideId`, which survives a reorder, and
 * `removeSlideAt` aims by POSITION, which does not. So the delete took whatever
 * had been dragged into that slot, the delta was exactly what success looks
 * like, and the pane reported it.
 *
 * Compared with `sameSlideId` rather than `===`, because a selection id can
 * lack the `#suffix` the deck's own list carries — `jump.ts` owns that rule and
 * this is the same comparison.
 *
 * Answers false when either id is missing. A read that did not answer is not
 * permission to delete: `CLAUDE.md`'s rule is that the failure mode should be a
 * duplicate the user can delete rather than a slide they have lost, and that is
 * the whole reason the removal comes after the insert.
 */
export function stillThere(expected: string | undefined, atIndex: string | undefined): boolean {
  if (expected === undefined || atIndex === undefined) return false;
  return sameSlideId(expected, atIndex);
}

/**
 * Where a slide IS now, given the deck's ids in order.
 *
 * `stillThere` above answers whether a slide has moved. This answers where it
 * moved TO, and the difference is the whole of what a run of several cycles
 * needs: a run that walks the INDEXES it was handed before it started is
 * walking a map of a deck that no longer exists the moment the user drags
 * anything.
 *
 * Measured on Windows on 2026-09-23, on a deck whose every slide was labelled.
 * A stamp over 39 selected slides with one slide dragged from position 30 to
 * position 5 mid-run: the guard caught the move that landed INSIDE a cycle and
 * stranded a copy honestly. An earlier run the same day, where the drag landed
 * BETWEEN two cycles, was not caught at all — nothing moves during a cycle, so
 * there is nothing for `stillThere` to compare — and the run reported "Stamped
 * 59 slides" where 58 slides had gained one. Nothing was lost, and one selected
 * slide was silently skipped.
 *
 * So the runs now carry slide IDS and ask this where each one currently sits,
 * every cycle. A slide that moved is followed; a slide that has GONE answers
 * `undefined`, which is a skip the caller must count rather than a position it
 * may delete.
 *
 * `sameSlideId` rather than `===`, for the reason `stillThere` gives: an id
 * from a selection read can lack the `#suffix` the deck's own list carries.
 */
export function indexOfSlide(ids: readonly (string | undefined)[], slide: string | undefined): number | undefined {
  if (slide === undefined) return undefined;
  for (let at = 0; at < ids.length; at += 1) {
    const here = ids[at];
    if (here !== undefined && sameSlideId(slide, here)) return at;
  }
  return undefined;
}

/**
 * Which slides a PART lands on, given what the host says is selected.
 *
 * `docs/DESIGN.md` section 5: "A part ignores the insert target: it always
 * lands on the slide the user is on. A stamp or a label with several slides
 * selected lands on every selected slide." The record asserted this as built
 * from the start and nothing implemented it — `currentSlide` kept only
 * `selected.items[0]` and the insert drove exactly one cycle, so a user who
 * selected slides 2, 5 and 9 and clicked the Confidential stamp got it on
 * slide 2 and plain success reported over the two that were untouched.
 *
 * The answer is a list of INDICES, counting from zero, sorted and with
 * duplicates gone — the host's own order is the deck's, but nothing in the
 * record says the selection is sorted, and the cycles below rely on it being
 * so to reason about what a positional delete shifts.
 *
 * `current` is the fallback and not a member: when the host answers nothing,
 * or answers one slide, the part lands where it always did, and the caller runs
 * the ordinary single-slide path rather than a loop of one. Answering an empty
 * list is what says "there is nothing here the loop should take over".
 */
export function stampTargets(selected: number[] | undefined): number[] {
  if (selected === undefined || selected.length < 2) return [];
  return [...new Set(selected)].sort((a, b) => a - b);
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
 * Why an undo must not run, when the deck's size says it has changed since the
 * insert it would take back — or undefined when it may run.
 *
 * `undoPlan` is positions, read at the insert, and nothing disarms the undo
 * when the user edits the deck. The count checks inside the undo start from a
 * count read at the PRESS, so they cannot see what happened before it. Two
 * paths through that gap delete one of the user's own slides and then report
 * "Undone.", derived from the code on 2026-09-24 and not yet run on a host:
 *
 * - "As a new slide", then PowerPoint's own Ctrl+Z (which `docs/MANUAL.md`
 *   recommends), then the pane's Undo. Ctrl+Z has already taken the new slide
 *   out, and the undo deletes whatever now sits at `index + 1` — the user's
 *   next slide — while the count drops by exactly the one it expects.
 * - Any slide added or deleted, which shifts every position after it.
 *
 * Comparing the count at the press with the count the insert left closes both,
 * and costs no host call: the undo reads the count first anyway. It is a
 * COUNT, so it sees only what changes one:
 *
 * - a drag changes none; `undoAim` below closes that one, and the pane asks it
 *   immediately after this;
 * - a count that moved and moved back passes — Ctrl+Z on a new slide and then
 *   a slide added in its place is the likely one — exactly as every undo did
 *   before this;
 * - an "onto this slide" insert taken back with Ctrl+Z TWICE, as the manual
 *   says to, leaves the count where the insert left it; `undoAlreadyReverted`
 *   below catches that one by the id of the user's own slide;
 * - on the web the count has been measured lagging a change the add-in made by
 *   2.8 seconds (2026-09-11), so a press straight after a Ctrl+Z may read the
 *   old count and pass. Whether a user's Ctrl+Z lags the same way is
 *   unmeasured, and a second read would cost every web undo that long.
 *
 * A refusal is the safe direction of wrong: it leaves the deck as it is.
 */
export function undoRefusal(expected: number, now: number): string | undefined {
  if (now === expected) return undefined;
  return (
    `The deck has changed since the insert — it has ${now} slides where the insert left ${expected} — ` +
    `so Undo could take back the wrong slide. Nothing was changed. ` +
    `Check the deck, and use PowerPoint's own Ctrl+Z if the insert is still there.`
  );
}

/**
 * Why an "onto this slide" undo must not run when the user's own slide is back
 * where the insert put its copy — the Ctrl+Z case `undoRefusal` cannot see.
 *
 * The manual tells the user to press Ctrl+Z twice to take back an insert onto a
 * slide, and after that the deck is the size the insert left it, so the count
 * check passes. The undo then put the pre-insert snapshot back beside the
 * user's slide and deleted the slide — their own, with anything they had done
 * to it since — and said "Undone.".
 *
 * The undo already reads the id at that position to aim the restore, and the
 * id of the user's own slide was known at the insert. That id is a SETTLED
 * one, the kind measured following a drag on Windows on 2026-09-23 (`bd91527`)
 * — not the id of a slide the run had just added, which is what the borrowed
 * rule warns about. If PowerPoint's Ctrl+Z gives the slide back under a new id
 * (unmeasured), this does not fire and the undo behaves as it did before; it
 * can never fire on the rebuilt slide, which is a different slide.
 */
export function undoAlreadyReverted(slide: number): string {
  return (
    `Your slide ${slide} is already back as it was — PowerPoint's Ctrl+Z has taken the insert back — ` +
    `so there is nothing for Undo to do. Nothing was changed.`
  );
}

/**
 * What `undoAim` decided.
 *
 * Three outcomes rather than a boolean, because "I could not tell" and "no"
 * must not be the same answer. A boolean forced one of them to borrow the
 * other's behaviour, and either borrowing is a defect: reading "could not
 * tell" as "no" breaks the Undo on every host that does not mark its ids,
 * and reading it as "yes" is a positional delete with a check that never ran
 * standing behind it.
 */
export type Aim =
  | { kind: "ok" }
  /**
   * Nothing to check WITH, so the caller does exactly what it did before this
   * check existed. Not an error and not a refusal — see `undoAim`.
   */
  | { kind: "unmarked"; why: "no-creation-id" | "no-listing" | "host-marks-nothing" }
  | { kind: "refuse"; detail: string };

/**
 * Whether the slide the insert ADDED is still where the insert left it, asked
 * by the creation id this add-in's own engine wrote into it.
 *
 * `undoRefusal` above asks the same question of the deck's SIZE and therefore
 * cannot see a drag, which changes no count. That was the last hole in the
 * Undo, and `docs/BACKLOG.md` names the route out of it: at the press, proceed
 * only when exactly one listed slide carries the rebuilt slide's creation id
 * and it sits where the insert left it.
 *
 * The id compared is the `#suffix` and ONLY the suffix, via `slideIdSuffix`
 * rather than `sameSlideId`. The difference matters and it is the whole
 * correctness of this function:
 *
 * - The PREFIX is `<p:sldId id>`, which the deck assigns and reuses. Two
 *   different slides can carry the same prefix at different times, so a check
 *   that fell back to it would answer "still there" about a slide that is not.
 * - The SUFFIX is `<p14:creationId val>`, drawn by `cloneSlide` and kept
 *   unique within the package by `freshCreationId`. It names the slide this
 *   add-in built and nothing else.
 * - `sameSlideId` deliberately treats "one side has no suffix" as a match,
 *   because a SELECTION id can lack one (office-js#2474). That tolerance is
 *   right where it lives and wrong here: it would turn "the host told me
 *   nothing about this slide's identity" into "yes, delete it".
 *
 * **Unmarked is not a refusal.** Three things leave nothing to check with, and
 * each falls back to the Undo exactly as it behaved before this function
 * existed — count-checked, positional, with the drag still open:
 *
 * - the entry carries no creation id, which is every undo armed before this
 *   shipped and any package whose slide carried none;
 * - the listing did not answer inside its budget, which is a read that failed
 *   and not a deck that changed — refusing there would trade a hole that opens
 *   on a drag for one that opens on a slow read, on the host whose reads are
 *   slow;
 * - the listing answered and NO id in it carries a suffix at all, which is a
 *   host that does not mark ids. Mac and iPad have never been measured
 *   (`docs/DESIGN.md` section 15 lists every host fact there as assumed), and
 *   an AppSource validator meeting a refusal on its first Undo is the failure
 *   this branch exists to avoid.
 *
 * Everything the listing CAN answer is answered strictly: gone, moved, or in
 * the deck twice all refuse, and the sentence says which was seen, because a
 * refusal that does not say what it saw is one the user cannot act on.
 *
 * `index` counts from ZERO, like `undoPlan`'s, and is the position the delete
 * is about to take — `plan.remove`, which is the rebuilt slide in both targets.
 */
export function undoAim(listed: readonly string[] | undefined, creationId: number | undefined, index: number): Aim {
  if (creationId === undefined) return { kind: "unmarked", why: "no-creation-id" };
  if (listed === undefined) return { kind: "unmarked", why: "no-listing" };

  const wanted = String(creationId);
  const at: number[] = [];
  let anyMarked = false;
  for (let i = 0; i < listed.length; i += 1) {
    const suffix = slideIdSuffix(listed[i]);
    if (suffix === undefined) continue;
    anyMarked = true;
    if (suffix === wanted) at.push(i);
  }

  if (!anyMarked) return { kind: "unmarked", why: "host-marks-nothing" };

  if (at.length > 1) {
    // BOUNDED, the way `readable` bounds a host's error. Naming every position
    // reads fine for the two a duplicated slide produces and becomes a footer
    // full of slide numbers if the premise fails badly — and this branch exists
    // precisely for a premise failing, so it is the branch least entitled to
    // assume a small number.
    const where =
      at.length === 2
        ? `as slide ${(at[0] ?? 0) + 1} and slide ${(at[1] ?? 0) + 1}`
        : `${at.length} times, the first as slide ${(at[0] ?? 0) + 1}`;
    return {
      kind: "refuse",
      detail:
        `The slide this add-in inserted is in the deck more than once — ${where} — ` +
        `so Undo cannot tell which copy to take back. ` +
        `Nothing was changed. Delete whichever copy you do not want.`,
    };
  }
  const only = at[0];
  if (only === undefined) {
    return {
      kind: "refuse",
      detail:
        `The slide this add-in inserted is no longer in the deck, so there is nothing for Undo to take back. ` +
        `Nothing was changed.`,
    };
  }
  if (only !== index) {
    return {
      kind: "refuse",
      detail:
        `The deck has been reordered since the insert — the slide this add-in inserted is now slide ` +
        `${only + 1}, where the insert left it at slide ${index + 1} — so Undo could take back the wrong slide. ` +
        `Nothing was changed. Check the deck, and use PowerPoint's own Ctrl+Z if the insert is still there.`,
    };
  }
  return { kind: "ok" };
}

/**
 * Which slide the element ended up on, counting from ONE.
 *
 * `undoPlan` from the other end, and the reason it lives beside it: the two are
 * the same arithmetic read in opposite directions, and the bug `undoPlan`'s
 * docstring records — a success reported over a slide that had not moved — was
 * an off-by-one between exactly these two. Keeping them apart is how they drift.
 *
 * "Onto this slide" rebuilt the slide the user was on and took the original
 * away, so the element is on that slide: `index + 1`. "As a new slide" put one
 * after it: `index + 2`.
 *
 * `index` counts from zero, like `undoPlan`'s and unlike `Attempt.slide`. The
 * answer counts from one because everything downstream of it — the "Used in
 * this deck" row, the card's grey boxes, the undo entry — is a slide number a
 * user reads.
 */
export function landedOn(entry: { target: Target; index: number }): number {
  return entry.index + (entry.target === "new" ? 2 : 1);
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
