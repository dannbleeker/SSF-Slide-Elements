import { describe, expect, it } from "vitest";
import type { Element } from "../src/core/catalogue/types.js";
import {
  arrowTo,
  blockedReason,
  borrowedLine,
  COACH,
  coaching,
  DEFAULT_SETTINGS,
  EMPTY,
  escapeCloses,
  footerOf,
  moveableAfter,
  offersOtherTarget,
  otherTarget,
  otherTargetLabel,
  primary,
  RECENT_DEPTH,
  remember,
  removableFrom,
  removalOutcome,
  removeLabel,
  removeQuestion,
  settingsLine,
  STEP_TITLE,
  stepFor,
  STEPS,
  toggle,
  wrapsSelection,
  type PaneState,
} from "../src/pane/steps.js";

/**
 * The pane's decisions, checked without a browser.
 *
 * Everything the picker does that a person could argue with is here rather than
 * in `render.ts`: which elements a search leaves, which category is open, what
 * the footer says, why the button cannot be pressed. That split is the point —
 * a rule expressed as a function over plain values is one the suite can hold to
 * an answer, where the same rule inside a DOM builder can only be checked by
 * looking at a screenshot.
 */

import { LIBRARY, browsing, element } from "./fixtures/pane.js";

describe("the steps", () => {
  it("each have a title the manual can quote", () => {
    expect(STEPS.length).toBeGreaterThan(0);
    for (const id of STEPS) expect(STEP_TITLE[id].length, id).toBeGreaterThan(0);
  });

  it("are states rather than a wizard: a pane with no library is loading, one with a problem says so", () => {
    expect(stepFor(EMPTY)).toBe("loading");
    expect(stepFor(browsing)).toBe("browse");
    // The problem WINS over having a library, because a pane that failed to
    // reload after a retry must not silently show the old one as if nothing
    // had happened.
    expect(stepFor({ ...browsing, problem: "offline" })).toBe("problem");
  });
});

describe("the one control", () => {
  it("cannot be pressed until an element is chosen, and says why", () => {
    const action = primary(browsing, "browse");
    expect(action.label).toBe("Insert an element");
    expect(action.disabled).toBe(true);
    expect(blockedReason(browsing, "browse")).toMatch(/choose an element/i);
  });

  it("can be pressed once one is", () => {
    expect(primary({ ...browsing, chosen: "one-box" }, "browse").disabled).toBe(false);
    expect(blockedReason({ ...browsing, chosen: "one-box" }, "browse")).toBe("");
  });

  it("locks while an insert runs, and says that instead", () => {
    // One insert at a time: two 0.4 s apart killed a sibling's tab.
    const busy = { ...browsing, chosen: "one-box", busy: true };
    expect(primary(busy, "browse").disabled).toBe(true);
    expect(blockedReason(busy, "browse")).toMatch(/one insert at a time/i);
  });

  it("offers a retry when the library did not come", () => {
    const failed = { ...EMPTY, problem: "the network refused" };
    expect(primary(failed, "problem")).toEqual({ label: "Try again", disabled: false });
    expect(blockedReason(failed, "problem")).toBe("the network refused");
  });

  it("has a sentence of its own when the failure did not come with one", () => {
    // `state.problem` is whatever the loader could say, and it is optional: a
    // raise with no message, or a step reached some other way, leaves it unset.
    // The fallback is the difference between a screen that says what to do and
    // one that says "undefined" — and nothing was executing it.
    const mute = { ...EMPTY, problem: undefined };
    expect(blockedReason(mute, "problem")).toContain("did not load");
    expect(blockedReason(mute, "problem")).toContain("try again");
  });

  it("says what it is doing while the library loads", () => {
    expect(blockedReason(EMPTY, "loading")).toMatch(/library/i);
    expect(primary(EMPTY, "loading").disabled).toBe(true);
  });
});

describe("what the pane says about the deck", () => {
  it("says which library was borrowed, and only when one was", () => {
    expect(borrowedLine(browsing)).toBeUndefined();
    expect(borrowedLine({ ...browsing, library: { ...LIBRARY, borrowed: "4:3 library, scaled to A4 slides." } })).toBe(
      "4:3 library, scaled to A4 slides.",
    );
  });

  it("shows the settings without opening the gear", () => {
    expect(settingsLine(DEFAULT_SETTINGS)).toBe("Inserting onto this slide, as one group.");
    expect(settingsLine({ target: "new", group: false, colours: "deck" })).toBe("Inserting as a new slide, loose.");
  });

  it("says the colours out loud only when they are NOT the deck's own", () => {
    // The default is the quiet one, because it is what almost everybody is on
    // and what the pane would be doing anyway. The other is a surprise worth a
    // clause: an element that ignores the deck it is in.
    expect(settingsLine(DEFAULT_SETTINGS)).not.toContain("colours");
    expect(settingsLine({ ...DEFAULT_SETTINGS, colours: "library" })).toBe(
      "Inserting onto this slide, as one group, in the library's own colours.",
    );
  });
});

describe("the footer", () => {
  it("is empty before anything has happened", () => {
    expect(footerOf(browsing)).toEqual({ detail: "", byHand: false, undo: 0, again: false, move: false });
  });

  it("carries the outcome, and offers Again only once something has been inserted", () => {
    const after = {
      ...browsing,
      recent: ["one-box"],
      undo: 1,
      outcome: { ok: true, byHand: false, name: "One box", detail: "3 → 4 → 3 slides, slide 2 replaced." },
    };
    const report = footerOf(after);
    expect(report.detail).toBe("3 → 4 → 3 slides, slide 2 replaced.");
    expect(report.again).toBe(true);
    expect(report.undo).toBe(1);
  });

  it("does not offer Again while an insert is still going", () => {
    expect(footerOf({ ...browsing, recent: ["one-box"], busy: true }).again).toBe(false);
  });

  it("passes on that the deck needs a hand", () => {
    const stuck = {
      ...browsing,
      outcome: { ok: false, byHand: true, name: "One box", detail: "delete slide 2 by hand." },
    };
    expect(footerOf(stuck).byHand).toBe(true);
  });

  it("offers Move to a new slide when the last insert can be moved", () => {
    // `docs/DESIGN.md` section 6. `main.ts` decides WHETHER by asking the
    // splice what the destination held; this is the footer's half of it.
    expect(footerOf({ ...browsing, moveable: "one-box", undo: 1 }).move).toBe(true);
  });

  it("does not offer it when nothing was moved onto a busy slide", () => {
    expect(footerOf({ ...browsing, undo: 1 }).move).toBe(false);
  });

  it("does not offer it once the history it needs has gone", () => {
    // The move is an undo followed by a second insert, so an offer standing
    // after Undo has been spent would be a button that cannot do what it says.
    expect(footerOf({ ...browsing, moveable: "one-box", undo: 0 }).move).toBe(false);
  });

  it("keeps offering it while an insert is going, for the render to disable", () => {
    // Unlike Again beside it. Again is a new insert and has no business being
    // drawn during one; this is about the insert that just happened, which is
    // still the last one. `pane-render` holds the disabling half.
    expect(footerOf({ ...browsing, moveable: "one-box", undo: 1, busy: true }).move).toBe(true);
  });
});

describe("the lists the pane remembers", () => {
  it("puts the newest first and never repeats one", () => {
    expect(remember(["b", "c"], "a", 6)).toEqual(["a", "b", "c"]);
    expect(remember(["a", "b", "c"], "b", 6)).toEqual(["b", "a", "c"]);
  });

  it("keeps only as many as the design says", () => {
    expect(RECENT_DEPTH).toBe(6);
    const many = ["1", "2", "3", "4", "5", "6"];
    expect(remember(many, "7", RECENT_DEPTH)).toEqual(["7", "1", "2", "3", "4", "5"]);
  });

  it("toggles a star on and off", () => {
    expect(toggle([], "a")).toEqual(["a"]);
    expect(toggle(["a", "b"], "a")).toEqual(["b"]);
  });
});

describe("the first-run coach marks", () => {
  it("are shown on the first open, and not once dismissed", () => {
    expect(coaching({ ...EMPTY, library: LIBRARY })).toBe(true);
    expect(coaching({ ...EMPTY, library: LIBRARY, coached: true })).toBe(false);
  });

  it("are not shown before the library is there", () => {
    // Nothing to be coached about yet, and the loading screen has its own job.
    expect(coaching(EMPTY)).toBe(false);
  });

  it("says three things, which is what section 4 asks for", () => {
    expect(COACH).toHaveLength(3);
    for (const line of COACH) expect(line.length).toBeGreaterThan(0);
  });
});

describe("the other insert target, for one insert", () => {
  it("is whichever one the gear is not set to", () => {
    expect(otherTarget(DEFAULT_SETTINGS)).toBe("new");
    expect(otherTarget({ ...DEFAULT_SETTINGS, target: "new" })).toBe("onto");
  });

  it("says what it would do, as an action rather than a setting", () => {
    expect(otherTargetLabel(DEFAULT_SETTINGS)).toBe("Insert as a new slide");
    expect(otherTargetLabel({ ...DEFAULT_SETTINGS, target: "new" })).toBe("Insert onto this slide");
  });

  it("offers nothing on a part, because a part ignores the target", () => {
    // `docs/DESIGN.md` section 5: a stamp or a marker always lands on the slide
    // the user is on. A menu offering "as a new slide" over one would be the
    // pane promising something the engine does not do — the same reason
    // `landingLine` refuses to say it.
    expect(offersOtherTarget(element({ id: "box", kind: "slide" }))).toBe(true);
    expect(offersOtherTarget(element({ id: "stamp", kind: "part", landing: "top-right" }))).toBe(false);
  });
});

/**
 * Which elements go AROUND the selected shape instead of on top of it.
 *
 * `docs/DESIGN.md` section 5 says markers wrap and the other cursor-landing
 * parts do not, and the library says which are markers by the CATEGORY —
 * nothing else in an element separates them, since `landing` is `cursor` and
 * `kind` is `part` for markers, flowchart shapes and icons alike.
 *
 * Until 2026-09-12 the test was `key.toLowerCase().includes("mark")`, written
 * inline in `insert()` in `src/pane/main.ts` and therefore never run by
 * anything. The cases below are why moving it was worth doing rather than just
 * tidy: the substring is right on the library as it stands and wrong on names
 * the same library could plausibly grow.
 */
describe("whether an element wraps the selection", () => {
  const inCategory = (key: string, name = key): Element =>
    element({ id: "m", kind: "part", landing: "cursor", category: { key, name } });

  it("wraps for the library's own markers category, by its Danish key", () => {
    // The committed catalogue's key, in both slide sizes.
    expect(wrapsSelection(inCategory("Markeringer", "Markers"))).toBe(true);
  });

  it("does not wrap for the other parts that also land at the cursor", () => {
    // Same `kind` and same `landing` as a marker. Only the category differs,
    // which is exactly why the category is what is asked.
    expect(wrapsSelection(inCategory("Flowchart ikoner", "Flowchart shapes"))).toBe(false);
    expect(wrapsSelection(inCategory("Stempler og lignende", "Stamps and labels"))).toBe(false);
  });

  it("does not wrap for a category that merely CONTAINS the word", () => {
    // The reason this function exists. "Danmarkskort" is a map of Denmark and
    // an entirely plausible category in a Danish slide library; the old
    // substring test matched it, and every element in it would have started
    // resizing itself around whatever the user had selected. "Markedsandel" is
    // market share, and gets past a `startsWith("marked")` reading of the same
    // idea. Both are near-misses by construction: change `startsWith` back to
    // `includes` in `steps.ts` and this is the assertion that goes red.
    expect(wrapsSelection(inCategory("Danmarkskort", "Map of Denmark"))).toBe(false);
    expect(wrapsSelection(inCategory("Markedsandel", "Market share"))).toBe(false);
  });

  it("reads the English name too, so a renamed deck keeps working", () => {
    // The key is the deck's own title and the owner may write it in either
    // language. Either side answering is enough.
    expect(wrapsSelection(inCategory("Cirkler", "Markers"))).toBe(true);
    expect(wrapsSelection(inCategory("Markers", "Cirkler"))).toBe(true);
  });

  it("does not mind case or the space around it", () => {
    expect(wrapsSelection(inCategory("  MARKERINGER  ", "  MARKERINGER  "))).toBe(true);
  });
});

/**
 * Whether the insert that just happened can be offered "Move to a new slide".
 *
 * `docs/DESIGN.md` section 6. Four conditions, and the offer is worth nothing
 * unless every one of them is checked: the move is an undo followed by a second
 * insert, so it can promise nothing an undo cannot deliver.
 *
 * Inline in `insert()` until 2026-09-12, in the file the coverage floor
 * exempts. `test/pane-wiring.test.ts` has always covered the offer being made
 * and withdrawn through the DOM; what it could not do is name each condition
 * and show it alone deciding the answer.
 */
describe("whether the last insert can be moved to a new slide", () => {
  const busy = { ok: true, target: "onto" as const, held: 2 };
  const whole = element({ id: "one-box", kind: "slide" });

  it("offers the element's id when all four conditions hold", () => {
    // The id rather than a flag: the second insert must place the element the
    // user actually placed, not whatever tile the pointer has wandered onto.
    expect(moveableAfter(busy, whole)).toBe("one-box");
  });

  it("offers nothing when the insert did not work", () => {
    expect(moveableAfter({ ...busy, ok: false }, whole)).toBeUndefined();
  });

  it("offers nothing when it already went onto a slide of its own", () => {
    // "As a new slide" is where the offer would put it. It is already there.
    expect(moveableAfter({ ...busy, target: "new" }, whole)).toBeUndefined();
  });

  it("offers nothing on a part, which ignores the target switch", () => {
    expect(moveableAfter(busy, element({ id: "stamp", kind: "part", landing: "top-right" }))).toBeUndefined();
  });

  it("offers nothing when the slide was empty, because the element covers nothing", () => {
    // `held` is what the splice counted in the bytes it was handed — no second
    // read and no host call. Zero is the whole reason the offer is conditional:
    // on an empty slide it is noise.
    expect(moveableAfter({ ...busy, held: 0 }, whole)).toBeUndefined();
  });

  it("offers on a slide holding exactly one thing, because the rule is more-than-none", () => {
    // The boundary. A `> 1` here would silently drop the commonest case there
    // is: a title and nothing else.
    expect(moveableAfter({ ...busy, held: 1 }, whole)).toBe("one-box");
  });
});

describe("taking a part off the slides it is on", () => {
  const stamp = element({ id: "approved", kind: "part", name: "Approved stamp", landing: "top-right" });
  const box = element({ id: "one-box", kind: "slide", name: "One box" });
  const read = { ...browsing, used: [{ element: "approved", slides: [2, 5, 9] }] };

  it("offers nothing until the deck has been read", () => {
    // Before the read the pane does not know what is in the deck, and a button
    // offering to remove something from nowhere is worse than no button.
    expect(removableFrom(stamp, browsing)).toEqual([]);
    expect(removableFrom(stamp, read)).toEqual([2, 5, 9]);
  });

  it("offers nothing on a whole-slide element", () => {
    // Section 4 puts this on a PART. "Remove" for a slide's worth of content is
    // the slide's own delete key.
    expect(removableFrom(box, { ...read, used: [{ element: "one-box", slides: [1] }] })).toEqual([]);
  });

  it("offers nothing for a part the deck has been read and does NOT hold", () => {
    // The third state, and the one that had no case: the deck HAS been read,
    // so the pane knows what is in it, and this element is not among it. Not
    // the same as "never asked" above, which is what `used === undefined` means
    // — and the `?? []` here is what keeps a `find` that matched nothing from
    // reaching the caller as undefined.
    const other = element({ id: "rejected", kind: "part", name: "Rejected stamp", landing: "top-right" });
    expect(removableFrom(other, read)).toEqual([]);
  });

  it("counts slides rather than shapes, and says so in the singular too", () => {
    expect(removeLabel([2, 5, 9])).toBe("Remove from 3 slides");
    expect(removeLabel([2])).toBe("Remove from 1 slide");
  });

  it("asks a question that names the slides and admits it cannot be undone", () => {
    // "3 slides" is not something a user can check; "slides 2, 5 and 9" is. And
    // the pane genuinely cannot undo it — Undo is one INSERT deep.
    const asked = removeQuestion(stamp, [2, 5, 9]);
    expect(asked).toBe("Take Approved stamp off slides 2, 5 and 9? The pane cannot undo this.");
  });

  it("reports how far a run got, and marks a short one as the user's to finish", () => {
    expect(removalOutcome("Approved stamp", 3, 3)).toEqual({
      ok: true,
      byHand: false,
      name: "Approved stamp",
      detail: "Removed from 3 slides.",
    });
    const short = removalOutcome("Approved stamp", 2, 3);
    expect(short.ok).toBe(false);
    expect(short.byHand).toBe(true);
    expect(short.detail).toContain("Removed from 2 of 3 slides");
    expect(short.detail).toContain("as they were");
  });

  it("says nothing changed when there was nothing left to remove", () => {
    /**
     * Reachable, and reachable BECAUSE of a fix. The removal re-reads which
     * slides carry the element from the deck it is about to change rather than
     * trusting the list the question was asked about — so a user who takes the
     * shapes off by hand between the question and the answer leaves it with an
     * empty list.
     *
     * It used to report `Removed from 0 slides.` and call that a success: a
     * sentence that reads as a glitch, on a run where nothing was wrong.
     */
    const none = removalOutcome("Approved stamp", 0, 0);
    expect(none.detail).not.toContain("0 slides");
    expect(none.detail).toBe("It is not on any slide any more, so nothing changed.");
    // Not a failure, and nothing for the user to finish by hand: there is
    // nothing left to do.
    expect(none.ok).toBe(true);
    expect(none.byHand).toBe(false);
  });
});

describe("what Escape shuts", () => {
  /**
   * The ladder, which lived inside the pane's key handler where nothing could
   * reach it. The ORDER is the rule: getting it wrong loses somebody's search
   * while they were only trying to shut a preview card.
   */
  const open = {
    ...browsing,
    removing: { id: "markeringer-1", slides: [1, 2], done: 0 } as PaneState["removing"],
    menuFor: "boxes:one-box",
    previewing: "one-box",
    gear: true,
    query: "flow",
    tags: ["stamp"],
  };

  it("backs out of what was opened last, in order", () => {
    // Peeled one at a time, so each rung is asserted against everything below
    // it still being open rather than against a state built to suit it.
    expect(escapeCloses(open)).toBe("removing");
    const noAsk = { ...open, removing: undefined };
    expect(escapeCloses(noAsk)).toBe("menu");
    const noMenu = { ...noAsk, menuFor: undefined };
    expect(escapeCloses(noMenu)).toBe("preview");
    const noCard = { ...noMenu, previewing: undefined };
    expect(escapeCloses(noCard)).toBe("gear");
    const noGear = { ...noCard, gear: false };
    expect(escapeCloses(noGear)).toBe("search");
  });

  it("clears the search for a tag with no query, and answers nothing when all is shut", () => {
    expect(escapeCloses({ ...browsing, tags: ["stamp"] })).toBe("search");
    expect(escapeCloses({ ...browsing, query: "flow" })).toBe("search");
    expect(escapeCloses(browsing)).toBeUndefined();
  });

  it("never reaches the search while anything else is open", () => {
    // The whole point of the ladder: a user shutting a card keeps their search.
    for (const also of [
      { previewing: "one-box" },
      { gear: true },
      { menuFor: "boxes:one-box" },
    ] as Partial<PaneState>[]) {
      expect(escapeCloses({ ...browsing, query: "flow", tags: ["stamp"], ...also })).not.toBe("search");
    }
  });
});

describe("where an arrow key moves the focus", () => {
  it("steps forward and back, and treats both axes the same", () => {
    expect(arrowTo("ArrowRight", 2, 6)).toBe(3);
    expect(arrowTo("ArrowDown", 2, 6)).toBe(3);
    expect(arrowTo("ArrowLeft", 2, 6)).toBe(1);
    expect(arrowTo("ArrowUp", 2, 6)).toBe(1);
  });

  it("lands on the FIRST tile when the focus is on none of them", () => {
    // Whichever direction: a user pressing Down from the search box expects the
    // first tile, and one pressing Up expects the same rather than the last.
    // This falls out of the clamp rather than out of a case of its own —
    // `main.ts` carried `if (at < 0) return 0`, and taking it out left every
    // case here green, so it went. The behaviour is still pinned here.
    expect(arrowTo("ArrowDown", -1, 6)).toBe(0);
    expect(arrowTo("ArrowUp", -1, 6)).toBe(0);
  });

  it("clamps at both ends rather than wrapping", () => {
    // A grid that jumps from the last tile back to the first reads as a glitch,
    // and the two ends are exactly where an off-by-one hides.
    expect(arrowTo("ArrowRight", 5, 6)).toBe(5);
    expect(arrowTo("ArrowLeft", 0, 6)).toBe(0);
  });

  it("answers nothing for a key that is not an arrow, or a list with no tiles", () => {
    expect(arrowTo("Enter", 2, 6)).toBeUndefined();
    expect(arrowTo("a", 2, 6)).toBeUndefined();
    expect(arrowTo("ArrowDown", -1, 0)).toBeUndefined();
  });

  it("handles a list of one, where every arrow stays put", () => {
    expect(arrowTo("ArrowRight", 0, 1)).toBe(0);
    expect(arrowTo("ArrowLeft", 0, 1)).toBe(0);
    expect(arrowTo("ArrowDown", -1, 1)).toBe(0);
  });
});
