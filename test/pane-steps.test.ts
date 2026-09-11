import { describe, expect, it } from "vitest";
import type { Element } from "../src/core/catalogue/types.js";
import {
  COACH,
  DEFAULT_SETTINGS,
  EMPTY,
  RECENT_DEPTH,
  STEPS,
  STEP_TITLE,
  blockedReason,
  borrowedLine,
  categoryHits,
  coaching,
  didYouMean,
  elementOf,
  footerOf,
  groups,
  isOpen,
  landingLine,
  matches,
  primary,
  remember,
  runOf,
  settingsLine,
  slideLine,
  stepFor,
  stepMatches,
  tagsOf,
  tileCount,
  toggle,
  type Library,
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

function element(over: Partial<Element> & { id: string }): Element {
  return {
    key: over.key ?? over.id,
    name: over.name ?? over.id,
    category: over.category ?? { key: "boxes", name: "White boxes" },
    slide: over.slide ?? 1,
    kind: over.kind ?? "slide",
    box: over.box ?? { x: 0.1, y: 0.2, w: 0.5, h: 0.5 },
    landing: over.landing ?? "layout",
    shapes: over.shapes ?? 1,
    tags: over.tags ?? [],
    markup: over.markup ?? { xml: "", rels: [], parts: [] },
    ...over,
  };
}

const LIBRARY: Library = {
  size: "16:9",
  width: 12192000,
  height: 6858000,
  version: "v1",
  categories: [
    { key: "boxes", name: "White boxes" },
    { key: "stamps", name: "Stamps and labels" },
  ],
  elements: [
    element({ id: "one-box", name: "One box", tags: ["boxes", "white"] }),
    element({ id: "two-boxes", name: "Two boxes", tags: ["boxes"] }),
    element({
      id: "flow-1",
      name: "Process flow, 1 box",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 1 },
    }),
    element({
      id: "flow-2",
      name: "Process flow, 2 boxes",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 2 },
    }),
    element({
      id: "approved",
      name: "Approved stamp",
      key: "Godkendt",
      category: { key: "stamps", name: "Stamps and labels" },
      landing: "top-right",
      kind: "part",
      tags: ["stamp"],
    }),
  ],
};

const browsing: PaneState = { ...EMPTY, library: LIBRARY };

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

describe("search", () => {
  it("matches the English name, the Danish key, the category and the tags", () => {
    const stamp = LIBRARY.elements[4] as Element;
    expect(matches(stamp, "approved")).toBe(true);
    // The key is Danish and the pane never shows it, but somebody who knows the
    // library by its Danish names should still find things.
    expect(matches(stamp, "godkendt")).toBe(true);
    expect(matches(stamp, "stamps")).toBe(true);
    expect(matches(stamp, "stamp")).toBe(true);
    expect(matches(stamp, "flowchart")).toBe(false);
  });

  it("matches every word rather than the phrase, so word order does not matter", () => {
    const boxes = LIBRARY.elements[1] as Element;
    // "Two boxes" is the name; a substring match on the whole phrase would find
    // nothing here, which is the defect this rule exists for.
    expect(matches(boxes, "boxes two")).toBe(true);
    expect(matches(boxes, "two boxes")).toBe(true);
    expect(matches(boxes, "two three")).toBe(false);
  });

  it("matches everything when the box is empty or only spaces", () => {
    const boxes = LIBRARY.elements[1] as Element;
    expect(matches(boxes, "")).toBe(true);
    expect(matches(boxes, "   ")).toBe(true);
  });
});

describe("what the list shows", () => {
  it("collapses a sized run to one tile, so 117 elements show as fewer", () => {
    const found = groups(LIBRARY, browsing);
    const names = found.flatMap((g) => g.elements.map((e) => e.id));
    expect(names).toContain("flow-1");
    // The second member of the run is reachable through the stepper, not as a
    // tile of its own.
    expect(names).not.toContain("flow-2");
    expect(tileCount(found)).toBe(4);
  });

  it("hands the whole run back for the stepper, in count order", () => {
    const flow = LIBRARY.elements[2] as Element;
    expect(runOf(LIBRARY, flow).map((e) => e.run?.count)).toEqual([1, 2]);
    // An element that is not part of a run is its own run of one, so the caller
    // never has to ask which kind it has.
    expect(runOf(LIBRARY, LIBRARY.elements[0] as Element).map((e) => e.id)).toEqual(["one-box"]);
  });

  it("drops a category a search emptied rather than showing an empty header", () => {
    const found = groups(LIBRARY, { ...browsing, query: "stamp" });
    expect(found.map((g) => g.key)).toEqual(["stamps"]);
  });

  it("narrows by every picked tag at once", () => {
    expect(tileCount(groups(LIBRARY, { ...browsing, tags: ["boxes"] }))).toBe(2);
    expect(tileCount(groups(LIBRARY, { ...browsing, tags: ["boxes", "white"] }))).toBe(1);
    expect(tileCount(groups(LIBRARY, { ...browsing, tags: ["boxes", "stamp"] }))).toBe(0);
  });

  it("orders tags by how many elements carry them", () => {
    // Most used first, so the line that is visible when it is closed is the one
    // worth having.
    expect(tagsOf(LIBRARY)[0]).toBe("boxes");
  });
});

describe("which categories are open", () => {
  it("starts collapsed and opens what the user opened", () => {
    expect(isOpen(browsing, "boxes")).toBe(false);
    expect(isOpen({ ...browsing, open: ["boxes"] }, "boxes")).toBe(true);
  });

  it("opens everything while a search or a tag is on", () => {
    // A query that matched three things in two categories would otherwise show
    // two closed headers and look like nothing was found.
    expect(isOpen({ ...browsing, query: "box" }, "stamps")).toBe(true);
    expect(isOpen({ ...browsing, tags: ["stamp"] }, "boxes")).toBe(true);
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

  it("says what it is doing while the library loads", () => {
    expect(blockedReason(EMPTY, "loading")).toMatch(/library/i);
    expect(primary(EMPTY, "loading").disabled).toBe(true);
  });
});

describe("what the pane says about the deck", () => {
  it("names the slide an element would land on", () => {
    expect(slideLine({ ...browsing, slide: 4 })).toBe("Slide 4.");
  });

  it("admits it when PowerPoint would not say", () => {
    // Honest rather than silently defaulting: an element landing on a slide the
    // user was not looking at is the complaint this avoids.
    expect(slideLine(browsing)).toMatch(/did not say/i);
  });

  it("says which library was borrowed, and only when one was", () => {
    expect(borrowedLine(browsing)).toBeUndefined();
    expect(borrowedLine({ ...browsing, library: { ...LIBRARY, borrowed: "4:3 library, scaled to A4 slides." } })).toBe(
      "4:3 library, scaled to A4 slides.",
    );
  });

  it("shows the settings without opening the gear", () => {
    expect(settingsLine(DEFAULT_SETTINGS)).toBe("Inserting onto this slide, as one group.");
    expect(settingsLine({ target: "new", group: false })).toBe("Inserting as a new slide, loose.");
  });
});

describe("where the preview card says an element will land", () => {
  it("follows the gear for a whole-slide element", () => {
    const whole = element({ id: "box", kind: "slide", landing: "layout" });
    expect(landingLine(whole, DEFAULT_SETTINGS)).toBe(
      "Lands below your slide's own title, scaled to fit the room under it.",
    );
    expect(landingLine(whole, { target: "new", group: true })).toBe("Lands as a new slide after this one.");
  });

  it("IGNORES the gear for a part, because the engine does", () => {
    // docs/DESIGN.md section 5: a part always lands on the slide the user is
    // on. Saying "as a new slide" over a stamp would be the pane promising
    // something the insert does not do.
    const stamp = element({ id: "stamp", kind: "part", landing: "top-right" });
    const asNew = landingLine(stamp, { target: "new", group: true });
    expect(asNew).toBe(landingLine(stamp, DEFAULT_SETTINGS));
    expect(asNew).not.toContain("new slide");
  });

  it("names the three ways a part lands", () => {
    expect(landingLine(element({ id: "a", kind: "part", landing: "top-right" }), DEFAULT_SETTINGS)).toContain(
      "top-right",
    );
    expect(landingLine(element({ id: "b", kind: "part", landing: "cursor" }), DEFAULT_SETTINGS)).toContain(
      "shape you have selected",
    );
    expect(landingLine(element({ id: "c", kind: "part", landing: "as-authored" }), DEFAULT_SETTINGS)).toContain(
      "where it sits in the library",
    );
  });
});

describe("the footer", () => {
  it("is empty before anything has happened", () => {
    expect(footerOf(browsing)).toEqual({ detail: "", byHand: false, undo: 0, again: false });
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

describe("finding an element by id", () => {
  it("answers the element, and nothing when the library has moved on", () => {
    expect(elementOf(LIBRARY, "one-box")?.name).toBe("One box");
    // A favourite kept from an older catalogue names an element that may no
    // longer exist, and the pane must not draw a tile for it.
    expect(elementOf(LIBRARY, "gone")).toBeUndefined();
    expect(elementOf(undefined, "one-box")).toBeUndefined();
    expect(elementOf(LIBRARY, undefined)).toBeUndefined();
  });
});

describe("what the user might have meant", () => {
  const library: Library = {
    ...LIBRARY,
    elements: [
      element({ id: "a", name: "Triangle, simple, with text at the corners" }),
      element({ id: "b", name: "White box, 1 large" }),
      element({ id: "c", name: "Confidential stamp" }),
      element({ id: "d", name: "Process flow, 3 boxes" }),
    ],
  };

  it("reaches a name through one of its WORDS, not only the whole string", () => {
    // "triangel" is thirty edits from the whole name and one from its first
    // word, and a query is usually one word.
    expect(didYouMean(library, "triangel")).toContain("Triangle, simple, with text at the corners");
  });

  it("only ever offers names the library really has, so picking one cannot fail", () => {
    const names = new Set(library.elements.map((e) => e.name));
    for (const suggestion of didYouMean(library, "stemp")) expect(names.has(suggestion)).toBe(true);
  });

  it("says nothing for a query too short to be a typo of anything", () => {
    expect(didYouMean(library, "ab")).toEqual([]);
    expect(didYouMean(library, "  ")).toEqual([]);
  });

  it("says nothing when the query is near nothing", () => {
    expect(didYouMean(library, "xylophone")).toEqual([]);
  });

  it("holds a short query to a tighter threshold than a long one", () => {
    // Without that, a three-letter typo reaches half the library.
    expect(didYouMean(library, "bax")).toContain("White box, 1 large");
    expect(didYouMean(library, "zzz")).toEqual([]);
  });

  it("gives at most three", () => {
    const many: Library = {
      ...LIBRARY,
      elements: Array.from({ length: 10 }, (_, i) => element({ id: String(i), name: "Box " + i })),
    };
    expect(didYouMean(many, "box").length).toBeLessThanOrEqual(3);
  });
});

describe("the category chips a search shows", () => {
  it("counts tiles, so they agree with the number beside the search", () => {
    const hits = categoryHits(LIBRARY, { ...EMPTY, query: "" });
    const total = hits.reduce((n, h) => n + h.count, 0);
    expect(total).toBe(tileCount(groups(LIBRARY, { ...EMPTY, query: "" })));
  });

  it("keeps showing the other categories once one is picked", () => {
    // The counts are taken WITHOUT the picked category applied. Otherwise
    // picking one would leave a single chip and no way across or back.
    const all = categoryHits(LIBRARY, { ...EMPTY, query: "" });
    const narrowed = categoryHits(LIBRARY, { ...EMPTY, query: "", category: all[0]?.key });
    expect(narrowed).toEqual(all);
    expect(narrowed.length).toBeGreaterThan(1);
  });

  it("narrows the list to the category that was picked", () => {
    const all = groups(LIBRARY, { ...EMPTY, query: "" });
    expect(all.length).toBeGreaterThan(1);
    const one = groups(LIBRARY, { ...EMPTY, query: "", category: all[0]!.key });
    expect(one.map((g) => g.key)).toEqual([all[0]!.key]);
  });

  it("still narrows by the query and the tags", () => {
    const hits = categoryHits(LIBRARY, { ...EMPTY, query: "box" });
    for (const hit of hits) expect(hit.count).toBeGreaterThan(0);
    expect(hits.length).toBeLessThanOrEqual(categoryHits(LIBRARY, { ...EMPTY, query: "" }).length);
  });
});

describe("greying the sizes a search did not ask for", () => {
  it("says yes for a member the query matches and no for one it does not", () => {
    const one = element({ id: "flow-1", name: "Process flow, 1 box" });
    const two = element({ id: "flow-2", name: "Process flow, 2 boxes" });
    expect(stepMatches(one, { ...EMPTY, query: "1 box" })).toBe(true);
    expect(stepMatches(two, { ...EMPTY, query: "1 box" })).toBe(false);
  });

  it("says yes for everything when nothing is being searched for", () => {
    const one = element({ id: "flow-1", name: "Process flow, 1 box" });
    expect(stepMatches(one, EMPTY)).toBe(true);
  });

  it("respects a picked tag as well as the query", () => {
    const tagged = element({ id: "a", name: "A", tags: ["boxes"] });
    expect(stepMatches(tagged, { ...EMPTY, tags: ["boxes"] })).toBe(true);
    expect(stepMatches(tagged, { ...EMPTY, tags: ["stamps"] })).toBe(false);
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
