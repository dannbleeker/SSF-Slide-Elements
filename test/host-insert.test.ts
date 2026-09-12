import { describe, expect, it, vi } from "vitest";
import { INSERTING, announcement, landedOn, mayRemove, outcomeOf, undoPlan, type Attempt } from "../src/host/insert.js";
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

  it("treats a deck that SHRANK the same way, rather than reporting a negative", () => {
    const out = outcomeOf({ ...base, inserted: 11, removed: undefined });
    expect(out.ok).toBe(false);
    expect(out.detail).toContain("nothing was changed");
  });
});

describe("an insert that landed and could not be tidied", () => {
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
