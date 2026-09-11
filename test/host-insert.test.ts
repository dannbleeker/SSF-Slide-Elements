import { describe, expect, it, vi } from "vitest";
import { INSERTING, announcement, mayRemove, outcomeOf, type Attempt } from "../src/host/insert.js";
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
