import { describe, expect, it } from "vitest";
import type { Element } from "../src/core/catalogue/types.js";
import { LIBRARY, browsing, element } from "./fixtures/pane.js";
import {
  categoryHits,
  didYouMean,
  elementOf,
  groups,
  isOpen,
  matches,
  offersOpenAll,
  runOf,
  stepMatches,
  tagsOf,
  tileCount,
} from "../src/pane/search.js";
import { EMPTY, type Library } from "../src/pane/steps.js";

/**
 * Which elements the picker shows, checked without a browser.
 *
 * `src/pane/search.ts`: the search itself, the categories it leaves, which of
 * them are open, the chips that narrow it, the sizes it greys out, and what it
 * offers when it finds nothing. Split out of `pane-steps.test.ts` on 2026-09-12
 * with the module it covers — the two halves of one feature had ended up 640
 * lines apart in a single file.
 */

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

  it("orders the whole line by count, and a tie by name", () => {
    // The pane slices this line to the first few tags (`TAGS_SHOWN` in
    // `render.ts`), so the order decides which tags a reader ever sees. Count
    // first — `zebra` is on three elements and last in the alphabet — and the
    // name only where the counts are equal.
    const library: Library = {
      ...LIBRARY,
      elements: [
        element({ id: "a", tags: ["zebra", "beta"] }),
        element({ id: "b", tags: ["zebra", "alpha"] }),
        element({ id: "c", tags: ["zebra"] }),
      ],
    };
    expect(tagsOf(library)).toEqual(["zebra", "alpha", "beta"]);
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

  it("offers a name once even when two elements carry it", () => {
    // Two elements can share a name — the same stamp in two categories, or two
    // sizes the run collapser did not group. A suggestion list that said
    // "White box" twice would spend two of its three slots on one answer.
    const twins = {
      ...LIBRARY,
      elements: [
        element({ id: "a", name: "White box, 1 large" }),
        element({ id: "b", name: "White box, 1 large" }),
        element({ id: "c", name: "Triangle, simple" }),
      ],
    };
    // "bax" rather than a longer query: the threshold is a third of the query's
    // length, so a nine-character one is not within reach of any single word.
    expect(didYouMean(twins, "bax")).toEqual(["White box, 1 large"]);
  });

  it("copes with an element that has no name at all", () => {
    // The distance function used to shortcut an empty string on either side.
    // Both shortcuts went on 2026-09-12 — one was unreachable, the other was
    // doing nothing the loops do not already do — and this is what holds the
    // general path to the same answer. A nameless element is simply never near
    // enough to suggest.
    const nameless = {
      ...LIBRARY,
      elements: [element({ id: "blank", name: "" }), element({ id: "box", name: "Box" })],
    };
    expect(didYouMean(nameless, "bax")).toEqual(["Box"]);
  });

  it("counts one edit per character, whichever direction the typo went", () => {
    const typos: Library = {
      ...LIBRARY,
      elements: [element({ id: "a", name: "Stamp" }), element({ id: "b", name: "Box" })],
    };
    // A letter missing: "stam" is one INSERTION short of "Stamp". A letter too
    // many: "boxs" is one DELETION from "Box". Both are one edit, and a
    // four-character query is allowed exactly one — so an edit charged at two
    // would put each of these out of reach in turn.
    expect(didYouMean(typos, "stam")).toEqual(["Stamp"]);
    expect(didYouMean(typos, "boxs")).toEqual(["Box"]);
  });

  it("puts the nearest name first, ahead of the alphabet", () => {
    // "Arrow" sorts before "Arrows" and is one edit further away, so a list
    // that came back in name order would look right in every other respect.
    // The exact match costs nothing and has to lead.
    const arrows: Library = {
      ...LIBRARY,
      elements: [element({ id: "a", name: "Arrow" }), element({ id: "b", name: "Arrows" })],
    };
    expect(didYouMean(arrows, "arrows")).toEqual(["Arrows", "Arrow"]);
  });

  it("allows one edit per three characters of the query, and no more", () => {
    // `Math.max(1, Math.floor(q.length / 3))`. The floor of one is what keeps a
    // three-letter typo from reaching half the library; the third of the length
    // is what lets a longer one reach further.
    const arrows: Library = {
      ...LIBRARY,
      elements: [element({ id: "a", name: "Arrow" }), element({ id: "b", name: "Arrows" })],
    };
    // Four characters, one edit: "Arrow" is one away and "Arrows" two.
    expect(didYouMean(arrows, "arow")).toEqual(["Arrow"]);
    const triangles: Library = {
      ...LIBRARY,
      elements: [element({ id: "a", name: "Triangle" }), element({ id: "b", name: "Triangles" })],
    };
    // Six characters, two edits: "Triangle" is two away and "Triangles" three.
    expect(didYouMean(triangles, "triang")).toEqual(["Triangle"]);
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
