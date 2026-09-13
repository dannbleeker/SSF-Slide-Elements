import { describe, expect, it } from "vitest";
import { browsing, element } from "./fixtures/pane.js";
import { fractionOf, landingLine, occupiedFor, slideLine, withLanded } from "../src/pane/card.js";
import { DEFAULT_SETTINGS } from "../src/pane/steps.js";

/**
 * The preview card, checked without a browser.
 *
 * `src/pane/card.ts`: which slide the card is describing, the sentence saying
 * where an element will land, the grey boxes behind it, and the rectangle drawn
 * once it has landed. Split out of `pane-steps.test.ts` on 2026-09-12 with the
 * module it covers.
 */

describe("which slide the card is describing", () => {
  it("names the slide an element would land on", () => {
    expect(slideLine({ ...browsing, slide: 4 })).toBe("Slide 4.");
  });

  it("admits it when PowerPoint would not say", () => {
    // Honest rather than silently defaulting: an element landing on a slide the
    // user was not looking at is the complaint this avoids.
    expect(slideLine(browsing)).toMatch(/did not say/i);
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

  it("refuses a size with EITHER side missing, not only both", () => {
    // Each half of the guard stands on its own. A read that came back with one
    // dimension and not the other is still a size nobody read, and dividing by
    // the zero half draws the rectangle at Infinity rather than not at all.
    expect(fractionOf(landed, { width: 0, height: 6858000 })).toBeUndefined();
    expect(fractionOf(landed, { width: 12192000, height: 0 })).toBeUndefined();
  });

  it("draws against any size above zero, because zero is the only thing wrong with it", () => {
    // The guard is `> 0` on each side, not a plausibility check on the size:
    // one EMU is a size the pane was told, so it is divided by. Whether the
    // rectangle that falls out is sensible is the deck's business.
    expect(fractionOf(landed, { width: 1, height: 1 })).toEqual({
      x: landed.x,
      y: landed.y,
      w: landed.cx,
      h: landed.cy,
    });
  });

  it("drops the snapshot when the size is unknown, rather than keeping a stale one", () => {
    const onSlide = { slide: 2, boxes: [{ x: 0.1, y: 0.1, w: 0.2, h: 0.2 }] };
    expect(withLanded(onSlide, 2, undefined, false)).toBeUndefined();
    expect(withLanded(onSlide, 2, undefined, true)).toBeUndefined();
  });
});
