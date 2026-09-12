import { describe, expect, it } from "vitest";
import { GLOBAL_KEY } from "../src/host/memory.js";
import { EMPTY, type PaneState } from "../src/pane/steps.js";
import { restored, shouldRestoreScroll, storedScroll, writes } from "../src/pane/storage.js";

/**
 * What the pane remembers and where (`docs/DESIGN.md` section 4).
 *
 * These rules were written inside `src/pane/main.ts`, which is excluded from
 * the coverage floor on the grounds that it is Office.js calls and DOM plumbing.
 * They are neither, so for as long as they lived there they were exempt from a
 * floor they should always have been under. Moving them into
 * `src/pane/storage.ts` is what this file exists to make good on: every branch
 * below was unmeasured until 2026-09-12.
 */
const state: PaneState = {
  ...EMPTY,
  favourites: ["one-box"],
  coached: true,
  recent: ["one-box", "marker"],
  open: ["boxes"],
  query: "flow",
  tags: ["stamp"],
  chosen: "one-box",
};

describe("what comes back out of the two buckets", () => {
  it("takes the favourites and the first-run flag from the MACHINE's bucket", () => {
    // A star is a statement about the library and the coach marks are
    // "dismissed once", so neither may come from the deck's half.
    const got = restored({ favourites: ["a"], coached: true }, { favourites: ["b"], coached: false });
    expect(got.favourites).toEqual(["a"]);
    expect(got.coached).toBe(true);
  });

  it("takes everything about how the library was being read from the DECK's", () => {
    const got = restored({}, { query: "flow", tags: ["stamp"], open: ["boxes"], recent: ["one-box"] });
    expect(got.query).toBe("flow");
    expect(got.tags).toEqual(["stamp"]);
    expect(got.open).toEqual(["boxes"]);
    expect(got.recent).toEqual(["one-box"]);
  });

  it("fills the settings in from the defaults, so an older bucket gains a new one", () => {
    // A bucket written before a setting existed must not leave that setting
    // undefined — the gear would draw a control with no value in it.
    const got = restored({}, { settings: { target: "new" } as PaneState["settings"] });
    expect(got.settings).toEqual({ ...EMPTY.settings, target: "new" });
  });

  it("answers an empty pane for two empty buckets rather than undefined fields", () => {
    const got = restored({}, {});
    expect(got.favourites).toEqual([]);
    expect(got.coached).toBe(false);
    expect(got.recent).toEqual([]);
    expect(got.open).toEqual([]);
    expect(got.query).toBe("");
    expect(got.tags).toEqual([]);
    expect(got.settings).toEqual(EMPTY.settings);
  });

  it("ignores a stored value that is not the type it should be", () => {
    // Hand-edited storage, or a bucket written by a build that spelled a field
    // differently. The pane opens either way: a wrong value costs that field,
    // never the open.
    const got = restored({}, { query: 7 as unknown as string, chosen: 7 as unknown as string });
    expect(got.query).toBe("");
    expect(got).not.toHaveProperty("chosen");
  });

  it("leaves `chosen` out entirely rather than carrying an undefined one", () => {
    expect(restored({}, {})).not.toHaveProperty("chosen");
    expect(restored({}, { chosen: "one-box" }).chosen).toBe("one-box");
  });
});

describe("how far down the list", () => {
  it("reads a positive number and nothing else", () => {
    expect(storedScroll({ scroll: 640 })).toBe(640);
    expect(storedScroll({})).toBe(0);
    expect(storedScroll({ scroll: 0 })).toBe(0);
    // Negative, not-a-number and infinite all mean "no usable value", because
    // the alternative is scrolling somebody to NaN on open.
    expect(storedScroll({ scroll: -20 })).toBe(0);
    expect(storedScroll({ scroll: Number.NaN })).toBe(0);
    expect(storedScroll({ scroll: Number.POSITIVE_INFINITY })).toBe(0);
    expect(storedScroll({ scroll: "640" })).toBe(0);
  });
});

describe("what gets written, and to how many keys", () => {
  it("splits the two halves when the host named a deck", () => {
    const out = writes("ssf-slide-elements:deadbeef", state, 640);
    expect(out.map(([key]) => key)).toEqual([GLOBAL_KEY, "ssf-slide-elements:deadbeef"]);
    expect(out[0]?.[1]).toEqual({ favourites: ["one-box"], coached: true });
    const deck = out[1]?.[1] as Record<string, unknown>;
    expect(Object.keys(deck).sort()).toEqual(["chosen", "open", "query", "recent", "scroll", "settings", "tags"]);
    // The deck's half must not carry the machine's, or a second deck would
    // inherit somebody's stars.
    expect(deck).not.toHaveProperty("favourites");
  });

  it("writes ONE merged key when the host would not name a deck", () => {
    // Two writes under the same key leave the second overwriting the first a
    // moment later, so the unsaved-deck fallback has to merge rather than
    // repeat. This is the rule that branch exists for.
    const out = writes(GLOBAL_KEY, state, 640);
    expect(out).toHaveLength(1);
    expect(out[0]?.[0]).toBe(GLOBAL_KEY);
    const both = out[0]?.[1] as Record<string, unknown>;
    expect(both["favourites"]).toEqual(["one-box"]);
    expect(both["query"]).toBe("flow");
  });

  it("leaves out a scroll of zero and a `chosen` nobody picked", () => {
    const deck = writes("ssf-slide-elements:deadbeef", { ...state, chosen: undefined }, 0)[1]?.[1] as Record<
      string,
      unknown
    >;
    expect(deck).not.toHaveProperty("scroll");
    expect(deck).not.toHaveProperty("chosen");
  });

  it("round-trips: what it writes is what `restored` reads back", () => {
    // The two halves are written by one function and read by another, and a
    // field renamed in one of them would go quiet rather than fail. This is the
    // case that would not.
    const out = writes("ssf-slide-elements:deadbeef", state, 640);
    const machine = out[0]?.[1] as Record<string, unknown>;
    const deck = out[1]?.[1] as Record<string, unknown>;
    const back = restored(machine, deck);
    expect(back.favourites).toEqual(state.favourites);
    expect(back.coached).toBe(true);
    expect(back.query).toBe(state.query);
    expect(back.tags).toEqual(state.tags);
    expect(back.open).toEqual(state.open);
    expect(back.recent).toEqual(state.recent);
    expect(back.chosen).toBe(state.chosen);
    expect(storedScroll(deck)).toBe(640);
  });
});

describe("whether to put the scroll back on this draw", () => {
  it("only once, only with tiles, and only if there is somewhere to go", () => {
    expect(shouldRestoreScroll({ restored: false, kept: 640, hasTiles: true })).toBe(true);
    // Already done: a later draw is the user's own doing.
    expect(shouldRestoreScroll({ restored: true, kept: 640, hasTiles: true })).toBe(false);
    // No tiles yet: the loading screen is one short paragraph, and scrolling
    // that to 640 leaves the user looking at nothing.
    expect(shouldRestoreScroll({ restored: false, kept: 640, hasTiles: false })).toBe(false);
    // Never scrolled.
    expect(shouldRestoreScroll({ restored: false, kept: 0, hasTiles: true })).toBe(false);
  });
});
