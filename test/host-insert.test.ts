import { describe, expect, it, vi } from "vitest";
import {
  INSERTING,
  announcement,
  landedOn,
  mayRemove,
  outcomeOf,
  stampTargets,
  indexOfSlide,
  stillThere,
  type Attempt,
  undoPlan,
  copyLanded,
  undoAim,
  undoRefusal,
  undoAlreadyReverted,
} from "../src/host/insert.js";
import { BUDGET, withTimeout } from "../src/host/timeout.js";

/**
 * What an insert DID, read from the deck's size.
 *
 * These are the sentences a user sees after clicking a tile, and they are
 * design rather than implementation: `docs/DESIGN.md` sections 6 and 10 write
 * them out word for word, so the cases below quote them rather than matching a
 * pattern. A sentence that drifts from the record is a sentence somebody
 * changed without changing the record.
 *
 * The rule underneath every one of them is that the COUNT decides, never the
 * raise. `CLAUDE.md` has both halves: a queued call that raises nothing has not
 * necessarily happened, and a call can raise and still have done the work.
 */

const base: Attempt = { target: "onto", slide: 2, before: 12, inserted: 13, removed: 12 };

describe("an insert that worked", () => {
  it("reports the deck through all three counts when it replaced a slide", () => {
    // The shape `docs/DESIGN.md` section 6 specifies: in the file, an insert
    // onto a slide is an insert plus a removal, and the footer shows both.
    expect(outcomeOf(base)).toEqual({
      ok: true,
      detail: "12 → 13 → 12 slides, slide 2 replaced.",
      byHand: false,
    });
  });

  it("reports two counts for a new slide, because nothing was removed", () => {
    expect(outcomeOf({ target: "new", slide: 2, before: 12, inserted: 13 })).toEqual({
      ok: true,
      detail: "12 → 13 slides.",
      byHand: false,
    });
  });
});

describe("an insert that raised and landed anyway", () => {
  /**
   * The host rule this whole function exists for, and the one case it had none
   * of — found by auditing `CLAUDE.md`'s host rules against the code on
   * 2026-09-12.
   *
   * "A call can raise and still have done the work. SSF-Merge's insert timed
   * out with both slides landed. Read the delta." Every `error` case in this
   * file paired the raise with a deck that did NOT grow, so the delta and the
   * error always agreed and nothing distinguished them. An early
   * `if (attempt.error) return failure` at the top of `outcomeOf` would have
   * kept the entire suite green — and would have told a user their insert
   * failed while the slide was on their screen, and left them no Undo for it.
   */
  it("reports success when the host raised and the deck grew regardless", () => {
    const out = outcomeOf({ ...base, inserted: 13, removed: 12, error: "RichApi.Error: timeout" });
    expect(out.ok, "the delta is the evidence, not the raise").toBe(true);
    expect(out.detail, "and the sentence does not mention a refusal").not.toMatch(/refused/i);
  });

  it("still reports the refusal when the raise came with no growth", () => {
    // The pair, so the case above cannot be read as "errors are ignored".
    const out = outcomeOf({ ...base, inserted: 12, removed: undefined, error: "RichApi.Error: timeout" });
    expect(out.ok).toBe(false);
    expect(out.detail).toMatch(/refused/i);
  });
});

describe("an insert that did not land", () => {
  it("names the host's reason when there was one", () => {
    const out = outcomeOf({ ...base, inserted: 12, removed: undefined, error: "InvalidArgument" });
    expect(out.ok).toBe(false);
    expect(out.detail).toBe("The insert was refused: InvalidArgument");
    expect(out.byHand).toBe(false);
  });

  it("says the deck is untouched when the host said nothing at all", () => {
    // The silent no-op: no error, no growth. Without the count in the sentence
    // the user is left wondering what to undo.
    const out = outcomeOf({ ...base, inserted: 12, removed: undefined });
    expect(out.detail).toBe("The insert did not confirm: the deck still has 12 slides, nothing was changed.");
    expect(out.ok).toBe(false);
  });

  it("says what a deck that SHRANK actually holds, rather than the count it used to", () => {
    // This case used to assert `toContain("nothing was changed")`, which is the
    // sentence above — and that sentence names `before`. Over a deck of 11 it
    // read "the deck still has 12 slides, nothing was changed": a count the
    // deck does not have, and a claim the delta refutes, in the one function
    // whose purpose is to say no more than the count supports.
    //
    // Reachable without any host misbehaving. The pane locks itself, not
    // PowerPoint, and on the web the insert plus its confirming count takes
    // seconds, so a user deleting a slide in that window produces it.
    const out = outcomeOf({ ...base, inserted: 11, removed: undefined });
    expect(out.ok).toBe(false);
    expect(out.detail).toBe(
      "The insert did not confirm: the deck has 11 slides where it had 12. Check the deck before inserting again.",
    );
    expect(out.detail, "the count the deck no longer has").not.toContain("12 slides,");
    expect(out.byHand, "only the user can say what went").toBe(true);
  });

  it("says the deck shrank even when the host ALSO raised", () => {
    // The raise was tested first, so an insert that both raised and left the
    // deck smaller came out as "The insert was refused: …" with
    // `byHand: false` — the mildest sentence the pane has, over a deck that had
    // lost a slide, with nothing telling the user to look. The raise does not
    // make the delta untrue, and the delta is the more serious of the two
    // facts, so it is asked about first and the reason is carried along.
    const out = outcomeOf({ ...base, inserted: 11, removed: undefined, error: "RichApi.Error: timeout" });
    expect(out.detail).toBe(
      "The insert was refused: RichApi.Error: timeout — and the deck has 11 slides where it had 12. " +
        "Check the deck before inserting again.",
    );
    expect(out.byHand, "the pane said nothing was worth looking at").toBe(true);
  });

  it("keeps the plain refusal when the deck is exactly the size it was", () => {
    // The pair: a raise with no delta behind it is still just a refusal, and
    // that sentence is quoted word for word in `docs/DESIGN.md` section 10.
    const out = outcomeOf({ ...base, inserted: 12, removed: undefined, error: "InvalidArgument" });
    expect(out.detail).toBe("The insert was refused: InvalidArgument");
    expect(out.byHand).toBe(false);
  });

  it("keeps the no-op sentence for a deck that is exactly the size it was", () => {
    // The pair. `landed === 0` is the silent no-op and its sentence is quoted
    // word for word in `docs/DESIGN.md` section 10; only the negative delta
    // moved.
    const out = outcomeOf({ ...base, inserted: 12, removed: undefined });
    expect(out.detail).toBe("The insert did not confirm: the deck still has 12 slides, nothing was changed.");
    expect(out.byHand).toBe(false);
  });
});

describe("an insert that landed and could not be tidied", () => {
  it("does not name a slide to delete when the deck came back SHORTER", () => {
    // The removal half's guard, the mirror of the insert half's. That branch
    // is written for `removed === before + 1`, where the copy is still on the
    // slide — and it fired for `removed < before` too, where the copy is gone
    // and the deck has lost something else. `slide` is then the slide the
    // element LANDED on, so the sentence told the user to delete their own
    // content: `landedOn({target: "onto", index})` is the same number.
    //
    // Same route as the insert half, which `docs/DESIGN.md` section 10 already
    // accepts: `countReaching` answers whatever it last saw after its pauses,
    // and the pane locks itself and not PowerPoint.
    const out = outcomeOf({ ...base, inserted: 13, removed: 11 });
    expect(out.detail, "it named the slide the element is on").not.toContain("delete slide");
    expect(out.detail).toBe(
      "The insert landed, but the deck now has 11 slides where it had 12. Check the deck before inserting again.",
    );
    expect(out.byHand).toBe(true);
    expect(out.ok).toBe(false);
  });

  it("names the slide the user has to delete", () => {
    // `docs/DESIGN.md` section 10, word for word. The pane does not try again:
    // a second positional delete on a deck whose shape it has already misread
    // is how the wrong slide goes.
    const out = outcomeOf({ ...base, removed: 13 });
    expect(out).toEqual({
      ok: false,
      detail: "The deck grew by one but the copy could not be removed: delete slide 2 by hand.",
      byHand: true,
    });
  });

  it("says the same when the removal was never attempted", () => {
    expect(outcomeOf({ ...base, removed: undefined }).byHand).toBe(true);
  });

  it("names an insert that brought MORE than the one slide, and touches nothing", () => {
    // Nothing in the design predicts it. It is named rather than folded into a
    // failure, and the deck is left alone, because a sweep over slides this
    // code cannot account for is how the wrong one goes.
    const out = outcomeOf({ ...base, inserted: 15, removed: undefined });
    expect(out.ok).toBe(false);
    expect(out.byHand).toBe(true);
    expect(out.detail).toContain("grew by 3 slides where one was expected");
    expect(out.detail).toContain("Nothing was removed");
  });

  it("names TWO as more than one, which is the boundary of that rule", () => {
    // The case above is a deck that grew by three, and it was the only one, so
    // the boundary itself was untested. Found by the mutation sweep on
    // 2026-09-12: `if (landed > 1)` changed to `if (landed > 2)` and the whole
    // suite stayed green.
    //
    // What that costs, measured rather than guessed: a deck that grew by two
    // then falls through to the removal branch and the pane says "The deck grew
    // by ONE but the copy could not be removed: delete slide 2 by hand" — a
    // sentence that is wrong about what happened and points the user at a
    // specific slide to delete. `mayRemove` is a separate rule and still
    // refuses, so nothing is deleted by the code; the damage is entirely in
    // what the user is told, which is the half of this function that is design.
    //
    // Two is also the likeliest way this ever happens: an insert carrying one
    // slide more than the package listed.
    const out = outcomeOf({ ...base, inserted: 14, removed: undefined });
    expect(out.ok).toBe(false);
    expect(out.byHand, "and the deck is left alone, not swept").toBe(true);
    expect(out.detail).toContain("grew by 2 slides where one was expected");
  });
});

describe("whether the replaced slide may be removed at all", () => {
  it("is true only for a deck that grew by exactly one", () => {
    expect(mayRemove({ before: 12, inserted: 13 })).toBe(true);
    expect(mayRemove({ before: 12, inserted: 12 })).toBe(false);
    expect(mayRemove({ before: 12, inserted: 14 })).toBe(false);
    expect(mayRemove({ before: 12, inserted: 11 })).toBe(false);
  });
});

describe("what the live region says", () => {
  it("names the element on success and repeats the failure otherwise", () => {
    expect(announcement(outcomeOf(base), "One box")).toBe("One box inserted. 12 → 13 → 12 slides, slide 2 replaced.");
    const failed = outcomeOf({ ...base, inserted: 12, removed: undefined, error: "no" });
    expect(announcement(failed, "One box")).toBe("The insert was refused: no");
  });

  it("has a word for the tile while the insert runs", () => {
    expect(INSERTING).toBe("Inserting…");
  });
});

describe("every call to the host is bounded", () => {
  it("answers the work when it settles in time", async () => {
    await expect(withTimeout(Promise.resolve(7), 1000, "counting")).resolves.toBe(7);
  });

  it("passes a rejection through as itself, so the reason survives", async () => {
    await expect(withTimeout(Promise.reject(new Error("refused")), 1000, "counting")).rejects.toThrow("refused");
  });

  it("wraps a thrown non-Error, because a host raises those too", async () => {
    // Office.js routinely rejects with a plain object rather than an Error, and
    // the wrapper has to turn that into something a caller can read a message
    // off. Built rather than written inline, because the lint rule that forbids
    // rejecting with a non-Error is right about production code and this is the
    // one place that has to reproduce it.
    const raw: Promise<never> = new Promise((_, reject) => {
      (reject as (reason: unknown) => void)("just a string");
    });
    await expect(withTimeout(raw, 1000, "counting")).rejects.toThrow("just a string");
  });

  it("gives up by name, saying what it was waiting for and for how long", async () => {
    vi.useFakeTimers();
    try {
      const work = withTimeout(new Promise(() => undefined), 20_000, "inserting the rebuilt slide");
      const caught = expect(work).rejects.toThrow(/gave up waiting for inserting the rebuilt slide after 20 seconds/);
      await vi.advanceTimersByTimeAsync(20_000);
      await caught;
    } finally {
      vi.useRealTimers();
    }
  });

  it("clears its timer when the work settles, so a slow budget does not hold the pane open", async () => {
    vi.useFakeTimers();
    try {
      await expect(withTimeout(Promise.resolve("done"), BUDGET.deck, "reading")).resolves.toBe("done");
      // Nothing is left pending: if the timer had survived, this would count it.
      expect(vi.getTimerCount()).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("gives the deck read a bigger budget than a count, because the two differ by orders of magnitude", () => {
    // Measured on the web on 2026-09-10: a count is a batch with nothing in it,
    // and `getFileAsync` took 40 seconds for 40 KB on a degraded session.
    expect(BUDGET.read).toBeLessThan(BUDGET.deck);
    expect(BUDGET.remove).toBeLessThan(BUDGET.insert);
  });
});

describe("what taking an insert back means, as indices", () => {
  /**
   * This is the arithmetic an undo performs, and it was wrong in the way that
   * is worst: it reported success and changed nothing.
   *
   * Undoing an insert that landed ONTO slide N means putting the user's own
   * slide N back and taking the rebuilt one away. The rebuilt one is at index N,
   * so the restored copy lands at N+1 — and it is N that must go. The code
   * removed N+1, which is the copy it had just restored, so the deck came back
   * to its old size, the count check passed, and the pane said "Undone" over a
   * slide that had not moved. Found by pressing Undo in PowerPoint and looking.
   */
  it("puts the original back beside the rebuilt slide and removes the REBUILT one", () => {
    const plan = undoPlan({ target: "onto", index: 0 });
    expect(plan.after, "the insert must aim at the rebuilt slide").toBe(0);
    // NOT plan.after + 1. That is the restored copy, and removing it is the
    // defect this case exists for.
    expect(plan.remove).toBe(0);
    expect(plan.grownTo(5)).toBe(6);
  });

  it("keeps the two indices apart on a slide further down the deck", () => {
    const plan = undoPlan({ target: "onto", index: 7 });
    expect(plan).toMatchObject({ after: 7, remove: 7 });
    expect(plan.grownTo(12)).toBe(13);
  });

  /**
   * The two numbers for the SAME slide, side by side, counted from different
   * ends.
   *
   * `outcomeOf` is read aloud, so its slide counts from one; `undoPlan` is fed
   * to the API, so its index counts from zero. The defect above was this
   * confusion by another name, and the fields are called `slide` and `index`
   * now so a call site cannot hand over the wrong one without the compiler
   * saying so. This case is the reason that naming may not be tidied away.
   */
  it("counts from zero where the sentence counts from one", () => {
    const at = 3; // the fourth slide, as PowerPoint's API numbers it
    const sentence = outcomeOf({ target: "onto", slide: at + 1, before: 9, inserted: 10, removed: 9 });
    expect(sentence.detail).toBe("9 → 10 → 9 slides, slide 4 replaced.");
    const plan = undoPlan({ target: "onto", index: at });
    expect(plan.remove, "the plan addresses the same slide, from zero").toBe(3);
  });

  it("takes a new slide back with one removal, from the position after its target", () => {
    const plan = undoPlan({ target: "new", index: 3 });
    expect(plan.after, "nothing is inserted to undo a new slide").toBeUndefined();
    expect(plan.remove).toBe(4);
    // The deck does not grow on the way: there is no insert half.
    expect(plan.grownTo(9)).toBe(9);
  });
});

/**
 * Where the element ended up, which is `undoPlan` read from the other end.
 *
 * These two are one piece of arithmetic and the reason they are tested together
 * is the defect in `undoPlan`'s own docstring: an off-by-one BETWEEN them
 * reported "Undone" over a slide that had not moved. Until 2026-09-12 this half
 * was a ternary inside `insert()` in `src/pane/main.ts`, which the coverage
 * floor exempts — so the pair that must agree had one half measured and one
 * half not.
 *
 * Three things downstream read the answer as a slide number a user sees: the
 * "Used in this deck" row, the preview card's grey boxes, and the undo entry
 * that puts both back.
 */
describe("which slide the element landed on", () => {
  it("names the slide the user was on when the insert went onto it", () => {
    // Index 0 is the first slide, and "onto" rebuilt it in place.
    expect(landedOn({ target: "onto", index: 0 })).toBe(1);
    expect(landedOn({ target: "onto", index: 7 })).toBe(8);
  });

  it("names the slide after it when the insert made a new one", () => {
    expect(landedOn({ target: "new", index: 0 })).toBe(2);
    expect(landedOn({ target: "new", index: 7 })).toBe(9);
  });

  it("agrees with the plan that takes the same insert back", () => {
    // The pair, held against each other at one index. `undoPlan` removes by
    // ZERO-based index and this answers a ONE-based slide number, so the slide
    // the undo removes is this answer: an insert onto slide 4 (index 3) put the
    // element on slide 4 and the undo takes index 3 away.
    const at = 3;
    expect(landedOn({ target: "onto", index: at })).toBe(undoPlan({ target: "onto", index: at }).remove + 1);
    // And for a new slide, the one removal is the slide the element is on.
    expect(landedOn({ target: "new", index: at })).toBe(undoPlan({ target: "new", index: at }).remove + 1);
  });
});

describe("stillThere", () => {
  /**
   * `mayRemove` asks only about the COUNT, and a count cannot see a reorder.
   * The index is read before the host calls and used after them, with the whole
   * splice and an insertSlidesFromBase64 between — up to BUDGET.insert, 180
   * seconds. The pane locks itself, not PowerPoint, so a user can drag a slide
   * in the thumbnail strip in that window, and dragging changes no count.
   *
   * The two halves of the cycle then disagree: the insert aims by
   * `targetSlideId` and survives a reorder, the removal aims by position and
   * does not. So the delete took whatever had been dragged into the slot, the
   * delta was exactly what success looks like, and the pane said it worked.
   */
  it("is true only when the slide at the index is the one the insert was aimed at", () => {
    expect(stillThere("256", "256")).toBe(true);
    expect(stillThere("256", "257"), "a slide was dragged into the slot").toBe(false);
  });

  it("accepts the suffixed spelling, because a selection id may lack it", () => {
    // `jump.ts` owns this rule — office-js#2474, a selection id without the
    // `#suffix` the deck's own list carries. Compared with `===` this would
    // refuse a slide that had not moved at all, and the pane would leave a
    // duplicate on every insert.
    expect(stillThere("256", "256#424201")).toBe(true);
    expect(stillThere("256#424201", "256")).toBe(true);
    expect(stillThere("256#424201", "257#424202")).toBe(false);
  });

  it("refuses when either id is missing, because a read that did not answer is not permission", () => {
    expect(stillThere(undefined, "256")).toBe(false);
    expect(stillThere("256", undefined)).toBe(false);
    expect(stillThere(undefined, undefined)).toBe(false);
  });
});

describe("indexOfSlide", () => {
  /**
   * `stillThere` answers whether a slide moved; this answers where it moved TO.
   *
   * A run of several cycles used to walk the INDEXES it was handed before it
   * started. Measured on Windows on 2026-09-23: a drag that landed INSIDE a
   * cycle was caught and stranded a copy honestly, and a drag that landed
   * BETWEEN two cycles was not caught at all — nothing moves during a cycle, so
   * `stillThere` has nothing to compare — and the run reported "Stamped 59
   * slides" where 58 had gained one.
   */
  it("finds a slide that has moved", () => {
    expect(indexOfSlide(["256", "257", "258"], "258")).toBe(2);
    // The drag the rounds actually did: the last slide to the front.
    expect(indexOfSlide(["258", "256", "257"], "258")).toBe(0);
  });

  it("answers undefined for a slide that has GONE, which is a skip and not a position", () => {
    // The dangerous shape: returning 0 here would delete the first slide of the
    // deck. `undefined` is the only safe answer, and the caller must count it.
    expect(indexOfSlide(["256", "257"], "999")).toBeUndefined();
    expect(indexOfSlide([], "256")).toBeUndefined();
    expect(indexOfSlide(["256"], undefined)).toBeUndefined();
  });

  it("accepts the suffixed spelling, the same rule stillThere follows", () => {
    expect(indexOfSlide(["256", "257#424202"], "257")).toBe(1);
    expect(indexOfSlide(["256", "257"], "257#424202")).toBe(1);
  });

  it("steps over a slot the host could not name rather than matching it", () => {
    // A collection load can answer short (`CLAUDE.md`), so a slot can be
    // undefined. Matching `undefined` to `undefined` would hand back the index
    // of a slide nobody identified.
    expect(indexOfSlide([undefined, "257"], "257")).toBe(1);
    expect(indexOfSlide([undefined, undefined], undefined)).toBeUndefined();
  });

  it("answers the FIRST match, so a duplicated id cannot widen what is deleted", () => {
    expect(indexOfSlide(["256", "257", "257"], "257")).toBe(1);
  });
});

describe("outcomeOf, when the deck moved under the insert", () => {
  it("names no slide number, because the positions it was given are the stale ones", () => {
    const out = outcomeOf({ target: "onto", slide: 3, before: 8, inserted: 9, moved: true });
    expect(out.ok).toBe(false);
    expect(out.byHand, "the user is not told to look").toBe(true);
    expect(out.detail).toContain("reordered while it ran");
    expect(out.detail).toContain("8 → 9 slides");
    // The sentence this replaces named one, and that is what sent a user to
    // delete their own content the last time this file guessed from a position.
    expect(out.detail, "named a slide to delete off a stale index").not.toContain("delete slide 3");
  });

  it("still asks for the slide by number when the deck did NOT move", () => {
    // The pair: a removal that simply failed is a different fact, and its
    // sentence is still the right one. A fix that answered "reordered" for
    // every un-removed copy would lose it.
    const out = outcomeOf({ target: "onto", slide: 3, before: 8, inserted: 9 });
    expect(out.detail).toContain("delete slide 3 by hand");
  });
});

describe("which slides a stamp lands on", () => {
  /**
   * `docs/DESIGN.md` section 5 has asserted "a stamp or a label with several
   * slides selected lands on every selected slide" since the record was
   * written, and nothing implemented it: `currentSlide` kept only
   * `selected.items[0]`, so selecting slides 2, 5 and 9 and clicking the
   * Confidential stamp put it on slide 2 and reported plain success over the
   * two that were untouched.
   */
  it("takes over only when the host names more than one slide", () => {
    // Under two, the ordinary single-slide path runs — answering a list of one
    // would be a loop of one, with a different footer sentence and the pane's
    // Undo disarmed for no reason.
    expect(stampTargets(undefined), "no selection read").toEqual([]);
    expect(stampTargets([]), "an empty selection").toEqual([]);
    expect(stampTargets([4]), "one slide").toEqual([]);
    expect(stampTargets([4, 7])).toEqual([4, 7]);
  });

  it("sorts what the host hands back, and drops a slide named twice", () => {
    // SORTED, because each cycle is net zero on the slide count — the copy
    // lands after the original and the original is taken away — which keeps
    // later indices where this code computed them ONLY if they are worked
    // through in order. Nothing in the record says the host's selection is
    // sorted; the probe measured that its ORDER matches the file's, which is
    // not the same promise.
    expect(stampTargets([9, 1, 4])).toEqual([1, 4, 9]);
    // A duplicate would stamp one slide twice and count it as two.
    expect(stampTargets([3, 3, 1])).toEqual([1, 3]);
    expect(stampTargets([3, 3]), "one slide named twice is still one slide").toEqual([3]);
  });

  it("keeps the zeroth slide, which is a position and not a missing value", () => {
    expect(stampTargets([2, 0])).toEqual([0, 2]);
  });
});

describe("an undo against a deck that has changed since its insert", () => {
  it("lets the undo run when the deck is the size the insert left it", () => {
    expect(undoRefusal(6, 6)).toBeUndefined();
  });

  it("refuses when the count has moved either way, and says what it saw", () => {
    // Ctrl+Z took the new slide back: one fewer. A slide added: one more.
    for (const now of [5, 7]) {
      const said = undoRefusal(6, now);
      expect(said, `a deck of ${now} passed an undo planned for 6`).toBeDefined();
      expect(said).toContain(`${now} slides where the insert left 6`);
      expect(said).toContain("Nothing was changed");
    }
  });

  it("says the user's slide is already back, and that nothing was changed", () => {
    const said = undoAlreadyReverted(3);
    expect(said).toContain("slide 3 is already back");
    expect(said).toContain("Nothing was changed");
  });
});

/**
 * The check that closes the drag.
 *
 * `undoRefusal` above reads the deck's SIZE, and a drag changes no size: the
 * user reorders the strip, the count agrees, and the undo deletes whatever has
 * been dragged into the slot its plan names. `undoAim` asks the deck's listing
 * for the slide by the creation id the engine wrote into it instead.
 *
 * The ids below are the shape both hosts answered on 2026-09-24 for probe
 * question 8 — `<p:sldId id>#<p14:creationId val>` — and the prefix is varied
 * deliberately where the suffix is what decides, because a check that fell
 * back to the prefix would pass these and be wrong: the prefix is the deck's
 * and the deck reuses it.
 */
describe("whether the slide an insert added is still where the insert left it", () => {
  const marked = ["256#111", "257#222", "258#333"];

  it("lets the undo run when the slide is at the position the delete will take", () => {
    expect(undoAim(marked, 222, 1)).toEqual({ kind: "ok" });
  });

  it("refuses when a drag has moved it, and names both slides", () => {
    // The count is untouched — this is the whole case — so nothing above this
    // function sees it.
    const said = undoAim(["257#222", "256#111", "258#333"], 222, 1);
    expect(said.kind).toBe("refuse");
    if (said.kind !== "refuse") throw new Error("unreachable");
    expect(said.detail).toContain("now slide 1");
    expect(said.detail).toContain("insert left it at slide 2");
    expect(said.detail).toContain("Nothing was changed");
  });

  it("refuses when the slide is gone, rather than deleting whatever took its place", () => {
    const said = undoAim(["256#111", "258#333"], 222, 1);
    expect(said.kind).toBe("refuse");
    if (said.kind !== "refuse") throw new Error("unreachable");
    expect(said.detail).toContain("no longer in the deck");
    expect(said.detail).toContain("Nothing was changed");
  });

  it("refuses when the deck holds it twice, and says which slides", () => {
    // A duplicated slide. Probe question 8 measured a creation id being unique
    // in the listing on both hosts, so this is a premise failing rather than a
    // case the hosts produce — and a premise that fails must not delete.
    const said = undoAim(["256#222", "257#222", "258#333"], 222, 0);
    expect(said.kind).toBe("refuse");
    if (said.kind !== "refuse") throw new Error("unreachable");
    expect(said.detail).toContain("more than once");
    expect(said.detail).toContain("slide 1 and slide 2");
    expect(said.detail).toContain("Nothing was changed");
  });

  it("bounds the sentence when the premise fails badly, rather than listing every slide", () => {
    // A footer full of slide numbers is not a sentence a user can act on, and
    // this is the branch least entitled to assume the number is small: it only
    // runs when a host has broken the uniqueness both hosts were measured
    // keeping.
    const many = Array.from({ length: 30 }, (_, i) => `${256 + i}#222`);
    const said = undoAim(many, 222, 0);
    expect(said.kind).toBe("refuse");
    if (said.kind !== "refuse") throw new Error("unreachable");
    expect(said.detail).toContain("30 times, the first as slide 1");
    expect(said.detail.length, "the footer sentence grew with the deck").toBeLessThan(220);
  });

  it("compares the suffix ONLY, so a reused prefix cannot stand in for it", () => {
    // `sameSlideId` would answer true for `256` against `256#222`, because a
    // selection id may carry no suffix (office-js#2474). Here that tolerance
    // would license a delete: the prefix is the deck's and is reused, so the
    // slide at index 1 carrying prefix 256 is not evidence of anything.
    const said = undoAim(["999#111", "256#777", "258#333"], 222, 1);
    expect(said.kind).toBe("refuse");
    if (said.kind !== "refuse") throw new Error("unreachable");
    expect(said.detail).toContain("no longer in the deck");
  });

  it("does not take a matching suffix at the wrong index for the right one", () => {
    expect(undoAim(marked, 333, 1).kind).toBe("refuse");
  });
});

/**
 * The three ways there is nothing to check WITH.
 *
 * Each falls back to the undo exactly as it behaved before this check existed
 * — count-checked and positional — rather than refusing. Refusing would break
 * the Undo on every host that does not mark its ids and on every read that
 * timed out, which is a worse trade than the drag it would close, and Mac and
 * iPad are unmeasured (`docs/DESIGN.md` section 15).
 */
describe("when there is nothing for the undo to check with", () => {
  it("says so when the entry carries no creation id", () => {
    expect(undoAim(["256#111"], undefined, 0)).toEqual({ kind: "unmarked", why: "no-creation-id" });
  });

  it("says so when the listing did not answer", () => {
    // A read that failed is not a deck that changed.
    expect(undoAim(undefined, 222, 1)).toEqual({ kind: "unmarked", why: "no-listing" });
  });

  it("says so when the host marks no id with a suffix at all", () => {
    expect(undoAim(["256", "257", "258"], 222, 1)).toEqual({ kind: "unmarked", why: "host-marks-nothing" });
  });

  it("still decides when only SOME ids carry a suffix, because one of them may be ours", () => {
    // A host that marks some and not others is not one that marks nothing, and
    // treating it as such would skip the check on a deck that can answer it.
    expect(undoAim(["256", "257#222", "258"], 222, 1)).toEqual({ kind: "ok" });
    expect(undoAim(["256", "257#222", "258"], 222, 0).kind).toBe("refuse");
  });

  it("reads an empty listing as a host that marks nothing, not as a slide that is gone", () => {
    // An empty deck cannot be the deck an undo was armed against, so nothing
    // here may be read as evidence the slide was deleted.
    expect(undoAim([], 222, 0)).toEqual({ kind: "unmarked", why: "host-marks-nothing" });
  });
});

/**
 * Whether the copy a cycle inserted is in the deck.
 *
 * The question a COUNT cannot answer, and the reason this exists: a run
 * confirms its insert by counting, and when the count does not confirm, two
 * very different decks produce that — an insert that never landed, and one that
 * landed while something else moved the count back. The second leaves a copy
 * the user has to be told about.
 *
 * Measured on Windows on 2026-09-24: a slide deleted mid-run took the count
 * back to where it started, the cycle broke, and the run said "The rest are as
 * they were" over a deck holding SLIDE-010 twice with the stamp on the copy.
 */
describe("whether the copy a cycle inserted is in the deck", () => {
  it("says yes when a listed slide carries the copy's creation id", () => {
    expect(copyLanded(["256#111", "257#4242", "258#333"], 4242)).toBe("yes");
  });

  it("says no when the host marks its ids and none of them is the copy", () => {
    // The honest "nothing landed": the deck is untouched and the milder
    // sentence is TRUE of it. This is the case the bare break assumed always
    // held.
    expect(copyLanded(["256#111", "258#333"], 4242)).toBe("no");
  });

  it("says it cannot tell when the listing did not answer", () => {
    expect(copyLanded(undefined, 4242)).toBe("unknown");
  });

  it("says it cannot tell when no id carries a suffix at all", () => {
    // A host that marks nothing. Mac and iPad are unmeasured and assumed to be
    // this shape, and "no" would be a claim about a deck nobody looked at.
    expect(copyLanded(["256", "257", "258"], 4242)).toBe("unknown");
  });

  it("says it cannot tell when there is no creation id to look for", () => {
    expect(copyLanded(["256#111"], undefined)).toBe("unknown");
  });

  it("compares the suffix ONLY, so a reused prefix is not the copy", () => {
    // `sameSlideId` would take a bare `4242` for `4242#anything`; the prefix is
    // the deck's and is reused, and reading it as the copy would tell a user a
    // slide is there when it is not.
    expect(copyLanded(["4242#999", "257#111"], 4242)).toBe("no");
  });
});
