import { describe, expect, it } from "vitest";
import { browsing, LIBRARY } from "./fixtures/pane.js";
import { slideList, slideParts, usedHeading, usedRows, withInsert, withoutInsert } from "../src/pane/used.js";

/**
 * "Used in this deck", checked without a browser.
 *
 * `src/pane/used.ts`: the rows, the slide numbers as a sentence and as pieces a
 * renderer can turn into links, the heading's three states, and the two updates
 * that keep the list in step with an insert and its undo. Split out of
 * `pane-steps.test.ts` on 2026-09-12 with the module it covers.
 */

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
