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
  fractionOf,
  groups,
  isOpen,
  landingLine,
  matches,
  occupiedFor,
  offersOtherTarget,
  otherTarget,
  otherTargetLabel,
  primary,
  remember,
  removableFrom,
  removalOutcome,
  removeLabel,
  removeQuestion,
  runOf,
  settingsLine,
  slideLine,
  offersOpenAll,
  slideList,
  slideParts,
  stepFor,
  stepMatches,
  tagsOf,
  tileCount,
  toggle,
  usedHeading,
  usedRows,
  withInsert,
  withLanded,
  withoutInsert,
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

describe("where the preview card says an element will land", () => {
  it("follows the gear for a whole-slide element", () => {
    const whole = element({ id: "box", kind: "slide", landing: "layout" });
    expect(landingLine(whole, DEFAULT_SETTINGS)).toBe(
      "Lands below your slide's own title, scaled to fit the room under it.",
    );
    expect(landingLine(whole, { target: "new", group: true, colours: "deck" })).toBe(
      "Lands as a new slide after this one.",
    );
  });

  it("IGNORES the gear for a part, because the engine does", () => {
    // docs/DESIGN.md section 5: a part always lands on the slide the user is
    // on. Saying "as a new slide" over a stamp would be the pane promising
    // something the insert does not do.
    const stamp = element({ id: "stamp", kind: "part", landing: "top-right" });
    const asNew = landingLine(stamp, { target: "new", group: true, colours: "deck" });
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

describe("what this deck already uses", () => {
  const used = [
    { element: "one-box", slides: [2] },
    { element: "two-boxes", slides: [3, 5, 11] },
  ];

  it("reads the slide numbers out as a person would say them", () => {
    expect(slideList([2])).toBe("slide 2");
    expect(slideList([2, 5])).toBe("slides 2 and 5");
    expect(slideList([11, 3, 5])).toBe("slides 3, 5 and 11");
    // Sorted and deduplicated, because the same element twice on one slide is
    // one place to look.
    expect(slideList([4, 4])).toBe("slide 4");
    expect(slideList([])).toBe("");
  });

  it("hands the renderer the same list in pieces, so each number can be a control", () => {
    expect(slideParts([2])).toEqual([{ text: "slide " }, { slide: 2 }]);
    expect(slideParts([11, 3, 5])).toEqual([
      { text: "slides " },
      { slide: 3 },
      { text: ", " },
      { slide: 5 },
      { text: " and " },
      { slide: 11 },
    ]);
    expect(slideParts([4, 4])).toEqual([{ text: "slide " }, { slide: 4 }]);
    expect(slideParts([])).toEqual([]);
    // One source for the wording: the string IS the pieces joined.
    for (const slides of [[2], [2, 5], [11, 3, 5], [4, 4], []]) {
      expect(
        slideParts(slides)
          .map((p) => ("text" in p ? p.text : String(p.slide)))
          .join(""),
      ).toBe(slideList(slides));
    }
  });

  it("names each element from the library", () => {
    const rows = usedRows(LIBRARY, used);
    expect(rows.map((r) => r.name)).toEqual(["One box", "Two boxes"]);
    expect(rows.map((r) => r.where)).toEqual(["slide 2", "slides 3, 5 and 11"]);
    expect(rows.every((r) => r.known)).toBe(true);
  });

  it("keeps an element this library cannot name, and says that is what it is", () => {
    // Eleven ids changed when the part keys were translated, so a deck stamped
    // before that names elements this catalogue does not have. Dropping the row
    // would make the deck look emptier than it is.
    const rows = usedRows(LIBRARY, [{ element: "fortroligt", slides: [1] }]);
    expect(rows[0]?.known).toBe(false);
    expect(rows[0]?.name).toBe("An element from an older version of the library");
    expect(rows[0]?.where).toBe("slide 1");
  });

  it("says which of the three states the section is in", () => {
    // Never asked, asked and empty, asked and answered. The first two look the
    // same on a screen unless the heading says otherwise, and they are not the
    // same fact at all.
    expect(usedHeading(browsing)).toBe("Used in this deck");
    expect(usedHeading({ ...browsing, reading: true })).toBe("Reading this deck…");
    expect(usedHeading({ ...browsing, used: [] })).toBe("Nothing from the library is in this deck yet");
    expect(usedHeading({ ...browsing, used })).toBe("Used in this deck (2)");
  });

  it("adds an insert to the list rather than re-reading the deck", () => {
    expect(withInsert(used, "one-box", 7)).toEqual([
      { element: "one-box", slides: [2, 7] },
      { element: "two-boxes", slides: [3, 5, 11] },
    ]);
    expect(withInsert(used, "new-thing", 1)).toContainEqual({ element: "new-thing", slides: [1] });
    // The same element onto a slide it is already on is one slide, not two.
    expect(withInsert(used, "one-box", 2)).toContainEqual({ element: "one-box", slides: [2] });
  });

  it("leaves the list alone when the deck has never been read", () => {
    // An insert is not a reason to start claiming the deck has been looked at:
    // undefined means "not asked", and it has to survive one.
    expect(withInsert(undefined, "one-box", 2)).toBeUndefined();
    expect(withoutInsert(undefined, "one-box", 2)).toBeUndefined();
  });

  it("takes an undone insert back out, and drops a row that is now empty", () => {
    expect(withoutInsert(used, "two-boxes", 5)).toEqual([
      { element: "one-box", slides: [2] },
      { element: "two-boxes", slides: [3, 11] },
    ]);
    expect(withoutInsert(used, "one-box", 2)).toEqual([{ element: "two-boxes", slides: [3, 5, 11] }]);
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

describe("what the card draws in grey", () => {
  const boxes = [{ x: 0.1, y: 0.1, w: 0.3, h: 0.2 }];
  const onSlide = { slide: 2, boxes };

  it("draws the snapshot when it is of the slide the user is on", () => {
    expect(occupiedFor({ ...browsing, slide: 2, onSlide })).toEqual(boxes);
  });

  it("draws nothing when the user has moved to another slide", () => {
    // A card showing slide two's furniture while the user is on slide five is
    // worse than a card showing none: the boxes exist to answer "will this land
    // on top of something", and an answer about another slide is a WRONG one.
    expect(occupiedFor({ ...browsing, slide: 5, onSlide })).toEqual([]);
  });

  it("draws nothing when nothing has been read, or the host will not say which slide", () => {
    expect(occupiedFor({ ...browsing, slide: 2 })).toEqual([]);
    expect(occupiedFor({ ...browsing, onSlide })).toEqual([]);
  });

  it("keeps up with an insert onto the slide it already knows", () => {
    const landed = { x: 0.2, y: 0.5, w: 0.6, h: 0.3 };
    expect(withLanded(onSlide, 2, landed, false)).toEqual({ slide: 2, boxes: [...boxes, landed] });
  });

  it("forgets rather than invents when the insert landed on a slide it had not read", () => {
    const landed = { x: 0.2, y: 0.5, w: 0.6, h: 0.3 };
    expect(withLanded(onSlide, 7, landed, false)).toBeUndefined();
    expect(withLanded(undefined, 2, landed, false)).toBeUndefined();
  });

  it("knows a NEW slide holds exactly what was put on it", () => {
    // A new slide is a clone with its placeholders emptied, and an empty
    // placeholder is not one of these boxes anyway — so the element is all
    // there is.
    const landed = { x: 0.2, y: 0.5, w: 0.6, h: 0.3 };
    expect(withLanded(onSlide, 3, landed, true)).toEqual({ slide: 3, boxes: [landed] });
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

describe("a landed rectangle, in fractions of the right slide", () => {
  // The splice reports where an element landed in the DESTINATION deck's EMU.
  const landed = { x: 3048000, y: 1714500, cx: 6096000, cy: 3429000 };
  const library16x9 = { width: 12192000, height: 6858000 };
  // A4 landscape, which is one of the sizes `libraryFor` hands the NEAREST
  // library to rather than an exact one.
  const a4 = { width: 10692000, height: 7560000 };

  it("measures against the deck the element landed in", () => {
    expect(fractionOf(landed, library16x9)).toEqual({ x: 0.25, y: 0.25, w: 0.5, h: 0.5 });
  });

  it("gives a DIFFERENT answer on a deck that borrowed the library, which is the whole point", () => {
    // This is the bug the function exists to prevent: the pane used to divide
    // by `library.width`/`library.height`, which are the LIBRARY deck's size.
    // On an exact 16:9 deck the two are the same number and nothing shows; on a
    // borrowed one the rectangle lands a sixth of a slide out.
    const borrowed = fractionOf(landed, a4) as { x: number; w: number };
    const wrong = fractionOf(landed, library16x9) as { x: number; w: number };
    expect(borrowed.x).toBeCloseTo(0.285, 3);
    expect(borrowed.w).toBeCloseTo(0.57, 3);
    expect(borrowed.x).not.toBeCloseTo(wrong.x, 3);
  });

  it("answers nothing rather than a fraction of a size nobody read", () => {
    expect(fractionOf(landed, undefined)).toBeUndefined();
    expect(fractionOf(landed, { width: 0, height: 0 })).toBeUndefined();
  });

  it("drops the snapshot when the size is unknown, rather than keeping a stale one", () => {
    const onSlide = { slide: 2, boxes: [{ x: 0.1, y: 0.1, w: 0.2, h: 0.2 }] };
    expect(withLanded(onSlide, 2, undefined, false)).toBeUndefined();
    expect(withLanded(onSlide, 2, undefined, true)).toBeUndefined();
  });
});

describe("offering to open every category", () => {
  const two: Library = {
    ...LIBRARY,
    categories: [
      { key: "boxes", name: "White boxes" },
      { key: "stamps", name: "Stamps and labels" },
    ],
  };

  it("offers while any category is still closed", () => {
    expect(offersOpenAll({ ...EMPTY, open: [] }, two)).toBe(true);
    expect(offersOpenAll({ ...EMPTY, open: ["boxes"] }, two)).toBe(true);
  });

  it("stops offering once they are all open, rather than staying as a button that does nothing", () => {
    expect(offersOpenAll({ ...EMPTY, open: ["boxes", "stamps"] }, two)).toBe(false);
  });

  it("stays out of the way of a search, which opens what it found by itself", () => {
    expect(offersOpenAll({ ...EMPTY, open: [], query: "box" }, two)).toBe(false);
    expect(offersOpenAll({ ...EMPTY, open: [], tags: ["white"] }, two)).toBe(false);
    // A query of only spaces is not a search.
    expect(offersOpenAll({ ...EMPTY, open: [], query: "   " }, two)).toBe(true);
  });
});
