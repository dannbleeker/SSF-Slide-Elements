import { describe, expect, it } from "vitest";
import {
  atCursor,
  authored,
  centreOf,
  centredOn,
  fitInside,
  isIdentity,
  moveFrom,
  ontoSlide,
  place,
  topRight,
  underTitle,
  wrapping,
} from "../src/core/splice/landing.js";
import type { Frames, Rect, SlideSize } from "../src/core/splice/landing.js";

/**
 * Where an element lands, held to the numbers `docs/DESIGN.md` section 5 asks for.
 *
 * The module is pure arithmetic over rectangles, which is the whole reason it
 * is a module: the landing is the part of the splice a user argues with — "it
 * landed on top of my title", "half of it is off the slide" — and every one of
 * those complaints is a rectangle this file can state exactly. A screenshot
 * cannot be diffed and a real deck cannot say WHICH rule moved a shape, so
 * nothing below opens a package.
 *
 * Two kinds of assertion, deliberately. Where the number IS the rule — a
 * stamp's margin is a tenth of its own size, a marker's air is eight per cent
 * of the shape it wraps, an element that fits is left byte-for-byte alone — the
 * rectangle is written out in EMU. Where the rule is a RELATION — scaled into
 * the body area, kept on the slide, proportions preserved — the relation is
 * what is asserted, because pinning a rounded coordinate there would fail on
 * the day somebody legitimately changed how the centring rounds and would say
 * nothing about whether the element still fits.
 *
 * Every case names the sentence of section 5 it holds the code to. A landing
 * rule that drifts from that prose is a bug in one of the two, and this file is
 * where the disagreement shows up.
 */

/** The 16:9 library slide, in EMU. 914400 to the inch, so this is 13.33in by 7.5in. */
const SLIDE: SlideSize = { width: 12192000, height: 6858000 };

/**
 * A square slide, used only where a case is about AREA.
 *
 * A rule compared by area must give a shape and the same shape stood on end the
 * same answer, and on a 16:9 slide a literal transposition cannot even be drawn
 * — anything as tall as the slide is wide runs off it, so the clamp confuses
 * the reading. On a square slide the transposition is exact and the arithmetic
 * has nowhere to hide.
 */
const SQUARE: SlideSize = { width: 6858000, height: 6858000 };

/** PowerPoint's own 16:9 title placeholder, which is what a customer deck usually has. */
const TITLE: Rect = { x: 838200, y: 365125, cx: 10515600, cy: 1325563 };

/** The bottom of that title: the line `underTitle` measures everything from. */
const TITLE_BOTTOM = TITLE.y + TITLE.cy;

/** A rectangle stood on its end: x for y, width for height. */
function transpose(rect: Rect): Rect {
  return { x: rect.y, y: rect.x, cx: rect.cy, cy: rect.cx };
}

/** Whether a rectangle is wholly on the slide, which is what every rule promises. */
function onSlide(rect: Rect, slide: SlideSize): boolean {
  return rect.x >= 0 && rect.y >= 0 && rect.x + rect.cx <= slide.width && rect.y + rect.cy <= slide.height;
}

describe("authored", () => {
  it("turns the catalogue's fractions into EMU on the destination slide", () => {
    // The one place the two spellings meet: the catalogue stores a box as
    // fractions of the LIBRARY slide, and everything downstream is EMU.
    expect(authored({ x: 0.25, y: 0.5, w: 0.5, h: 0.25 }, SLIDE)).toEqual({
      x: 3048000,
      y: 3429000,
      cx: 6096000,
      cy: 1714500,
    });
  });

  it("rounds to a whole EMU, because a fractional coordinate is not a coordinate", () => {
    // `<a:off x="1354666.6666666667"/>` is not a number the schema allows, and
    // PowerPoint answers a package containing one by offering to repair it. A
    // ninth of the slide is the sharp case: rounding and truncating differ.
    const rect = authored({ x: 1 / 9, y: 1 / 9, w: 1 / 9, h: 1 / 9 }, SLIDE);
    expect(rect.x, "rounded, not truncated").toBe(1354667);
    expect(Math.trunc(SLIDE.width / 9), "which the other answer would have been").toBe(1354666);
    for (const v of Object.values(rect)) expect(Number.isInteger(v)).toBe(true);
  });

  it("measures against the destination slide, so a 4:3 deck gets a narrower element", () => {
    // The element is authored once and lands on whatever the user has open. A
    // box read against the LIBRARY's size instead would put a 16:9 element off
    // the right-hand edge of every 4:3 deck.
    const box = { x: 0.1, y: 0.1, w: 0.5, h: 0.5 };
    expect(authored(box, SLIDE).cx).toBe(6096000);
    expect(authored(box, { width: 9144000, height: 6858000 }).cx).toBe(4572000);
  });
});

describe("centredOn and centreOf", () => {
  it("moves a rectangle onto a centre without resizing it", () => {
    expect(centredOn({ x: 0, y: 0, cx: 1000, cy: 400 }, { x: 5000, y: 5000 })).toEqual({
      x: 4500,
      y: 4800,
      cx: 1000,
      cy: 400,
    });
  });

  it("answers the middle of a rectangle", () => {
    expect(centreOf({ x: 100, y: 200, cx: 400, cy: 600 })).toEqual({ x: 300, y: 500 });
  });

  it("keeps a half EMU in the centre and rounds once, when the rectangle moves", () => {
    // `centreOf` is not allowed to round: the cursor and wrap rules feed its
    // answer straight into `centredOn`, and rounding at both ends drifts a
    // shape by an EMU per insert in a direction nobody chose.
    expect(centreOf({ x: 0, y: 0, cx: 3, cy: 5 })).toEqual({ x: 1.5, y: 2.5 });
    expect(centredOn({ x: 0, y: 0, cx: 2, cy: 2 }, { x: 1.5, y: 2.5 })).toEqual({ x: 1, y: 2, cx: 2, cy: 2 });
  });

  it("leaves a rectangle centred on its own centre exactly where it is", () => {
    // The composition the cursor rule performs on every insert with a shape
    // selected. If it were not the identity, an element would creep.
    const rect: Rect = { x: 1000000, y: 2000000, cx: 400000, cy: 600000 };
    expect(centredOn(rect, centreOf(rect))).toEqual(rect);
  });
});

describe("ontoSlide", () => {
  it("moves an element that is merely in the wrong place, and does not shrink it", () => {
    // The order the module's own comment calls "the whole of the behaviour":
    // moved before it is shrunk. An element that would fit if it were somewhere
    // else keeps the proportions the owner drew.
    const rect: Rect = { x: SLIDE.width - 1000, y: 0, cx: 2000000, cy: 1000000 };
    const landed = ontoSlide(rect, SLIDE);
    expect(landed).toEqual({ x: 10192000, y: 0, cx: 2000000, cy: 1000000 });
    expect(landed.x + landed.cx, "flush with the right-hand edge").toBe(SLIDE.width);
  });

  it("shrinks nothing even when the element is exactly the size of the slide", () => {
    // The sharp end of the same rule. A slide-sized element dumped past the
    // bottom-right corner comes back whole, not scaled to a stamp.
    const rect: Rect = { x: SLIDE.width, y: SLIDE.height, cx: SLIDE.width, cy: SLIDE.height };
    expect(ontoSlide(rect, SLIDE)).toEqual({ x: 0, y: 0, cx: SLIDE.width, cy: SLIDE.height });
  });

  it("pulls an element back from a negative corner", () => {
    expect(ontoSlide({ x: -500000, y: -500000, cx: 1000000, cy: 1000000 }, SLIDE)).toEqual({
      x: 0,
      y: 0,
      cx: 1000000,
      cy: 1000000,
    });
  });

  it("scales an element wider than the slide rather than pushing it off", () => {
    // "It disappeared" is how a user reports half an element off the canvas,
    // and this is the last place any rule can stop that happening.
    const rect: Rect = { x: 0, y: 0, cx: SLIDE.width * 2, cy: SLIDE.height / 2 };
    const landed = ontoSlide(rect, SLIDE);
    expect(landed).toEqual({ x: 0, y: 0, cx: 12192000, cy: 1714500 });
    expect(landed.cx / landed.cy, "and keeps its proportions").toBeCloseTo(rect.cx / rect.cy, 6);
  });

  it("scales one taller than the slide the same way", () => {
    const landed = ontoSlide({ x: 0, y: 0, cx: SLIDE.width / 2, cy: SLIDE.height * 2 }, SLIDE);
    expect(landed).toEqual({ x: 0, y: 0, cx: 3048000, cy: 6858000 });
  });

  it("scales on both axes without stretching, which is the case a 4:3 element on a 16:9 slide reaches", () => {
    // Both guards fire, one after the other. Taking each axis to the slide's
    // own would answer the slide itself and stretch the element to fill it.
    const rect: Rect = { x: 0, y: 0, cx: SLIDE.width * 2, cy: SLIDE.height * 3 };
    const landed = ontoSlide(rect, SLIDE);
    expect(landed).toEqual({ x: 0, y: 0, cx: 8128000, cy: 6858000 });
    expect(landed.cx / landed.cy).toBeCloseTo(rect.cx / rect.cy, 5);
  });

  it("leaves an element that already fits on the slide exactly alone", () => {
    const rect: Rect = { x: 1000000, y: 1000000, cx: 2000000, cy: 1000000 };
    expect(ontoSlide(rect, SLIDE)).toEqual(rect);
  });
});

describe("fitInside", () => {
  it("hands back the very same rectangle when it already fits, rather than centring it", () => {
    // `fitInside` centres whatever it scales, so a rectangle that fits and got
    // centred anyway would be silently repositioned — which is exactly what the
    // layout rule promises never to do to an element that fits.
    const rect: Rect = { x: 100, y: 100, cx: 500, cy: 500 };
    expect(fitInside(rect, { x: 0, y: 0, cx: 1000, cy: 1000 })).toBe(rect);
  });

  it("treats a rectangle exactly the size of the area as fitting", () => {
    // The boundary. `<` rather than `<=` would scale an exact fit by 1 and
    // recentre it, moving an element that had no reason to move.
    const rect: Rect = { x: 400, y: 400, cx: 1000, cy: 1000 };
    expect(fitInside(rect, { x: 0, y: 0, cx: 1000, cy: 1000 })).toBe(rect);
  });

  it("scales a wide element by the smaller ratio, so it does not come back stretched", () => {
    // Taking each axis separately gives 1000 by 1000 here: a 4:1 banner
    // returned as a square, which is how an element comes back distorted.
    const fitted = fitInside({ x: 0, y: 0, cx: 2000, cy: 500 }, { x: 100, y: 100, cx: 1000, cy: 1000 });
    expect(fitted).toEqual({ x: 100, y: 475, cx: 1000, cy: 250 });
    expect(fitted.cx / fitted.cy, "still 4:1").toBe(4);
  });

  it("scales by the other ratio when the element is the tall one", () => {
    const fitted = fitInside({ x: 0, y: 0, cx: 500, cy: 2000 }, { x: 100, y: 100, cx: 1000, cy: 1000 });
    expect(fitted).toEqual({ x: 475, y: 100, cx: 250, cy: 1000 });
    expect(fitted.cx / fitted.cy).toBe(0.25);
  });

  it("centres what it scaled inside the area, on both axes", () => {
    const into: Rect = { x: 1000, y: 2000, cx: 900, cy: 900 };
    const fitted = fitInside({ x: 0, y: 0, cx: 1200, cy: 600 }, into);
    expect(fitted).toEqual({ x: 1000, y: 2225, cx: 900, cy: 450 });
    expect(centreOf(fitted)).toEqual(centreOf(into));
  });
});

describe("topRight", () => {
  it("leaves a tenth of the stamp's own width and height between it and the corner", () => {
    // `docs/DESIGN.md` section 5: a stamp lands top-right "with a 10% margin
    // because a rotated stamp's visible ellipse pokes past its frame".
    const stamp: Rect = { x: 0, y: 0, cx: 1000000, cy: 500000 };
    const landed = topRight(stamp, SLIDE);
    expect(landed).toEqual({ x: 11092000, y: 50000, cx: 1000000, cy: 500000 });
    expect(SLIDE.width - (landed.x + landed.cx), "a tenth of its width from the right edge").toBe(100000);
    expect(landed.y, "a tenth of its height from the top").toBe(50000);
  });

  it("gives a stamp twice the size twice the margin, because the ink that pokes out scales with it", () => {
    // The margin is a fraction of the ELEMENT, not of the slide, and this is
    // what makes it do its job. A slide-proportional margin would be too small
    // on the owner's big stamps and wasteful on the small ones.
    const small = topRight({ x: 0, y: 0, cx: 1000000, cy: 500000 }, SLIDE);
    const large = topRight({ x: 0, y: 0, cx: 2000000, cy: 1000000 }, SLIDE);
    expect(SLIDE.width - (small.x + small.cx)).toBe(100000);
    expect(SLIDE.width - (large.x + large.cx)).toBe(200000);
    expect(large.y).toBe(2 * small.y);
  });

  it("keeps a stamp too wide for its own margin on the slide", () => {
    // A slide-wide stamp cannot have a tenth of itself to spare, so the rule
    // asks for a negative x. Every rule ends in `ontoSlide` precisely so that
    // one of them wanting the impossible does not put ink off the canvas.
    const landed = topRight({ x: 0, y: 0, cx: SLIDE.width, cy: 1000000 }, SLIDE);
    expect(landed.x).toBe(0);
    expect(onSlide(landed, SLIDE)).toBe(true);
  });
});

describe("atCursor", () => {
  it("puts an element in the middle of the slide when nothing is selected", () => {
    // An add-in cannot see the mouse, so with no selection "at the cursor" is
    // the slide's centre and the user drags from there.
    const landed = atCursor({ x: 0, y: 0, cx: 2000000, cy: 1000000 }, SLIDE);
    expect(landed).toEqual({ x: 5096000, y: 2929000, cx: 2000000, cy: 1000000 });
    expect(centreOf(landed)).toEqual({ x: SLIDE.width / 2, y: SLIDE.height / 2 });
  });

  it("centres it on the selected shape when the host named one", () => {
    const selection: Rect = { x: 1000000, y: 1000000, cx: 2000000, cy: 2000000 };
    const landed = atCursor({ x: 0, y: 0, cx: 1000000, cy: 600000 }, SLIDE, selection);
    expect(landed).toEqual({ x: 1500000, y: 1700000, cx: 1000000, cy: 600000 });
    expect(centreOf(landed)).toEqual(centreOf(selection));
  });

  it("keeps it on the slide when the selection is against the edge, even though that un-centres it", () => {
    // The selected shape can be in the top-right corner, and centring on it
    // would hang most of the element off two edges. Landing somewhere
    // reasonable beats landing where the arithmetic said.
    const selection: Rect = { x: SLIDE.width - 100000, y: 0, cx: 100000, cy: 100000 };
    const landed = atCursor({ x: 0, y: 0, cx: 2000000, cy: 1000000 }, SLIDE, selection);
    expect(landed).toEqual({ x: 10192000, y: 0, cx: 2000000, cy: 1000000 });
    expect(onSlide(landed, SLIDE)).toBe(true);
    expect(centreOf(landed), "the clamp wins over the centring").not.toEqual(centreOf(selection));
  });
});

describe("wrapping", () => {
  /** The marker, at whatever size the library drew it. Most cases here ignore it, which is the point. */
  const MARKER: Rect = { x: 0, y: 0, cx: 1000000, cy: 600000 };

  it("wraps a small selected shape with eight per cent of its size of air on every side", () => {
    // Section 5: "sized to the shape with a little air and centred on it". A
    // marker drawn tight to the box reads as a border rather than as ink.
    const selection: Rect = { x: 2000000, y: 2000000, cx: 1000000, cy: 800000 };
    const landed = wrapping(MARKER, SLIDE, selection);
    expect(landed).toEqual({ x: 1920000, y: 1936000, cx: 1160000, cy: 928000 });
    expect(landed.x, "air on the left").toBe(selection.x - 80000);
    expect(landed.x + landed.cx, "the same air on the right").toBe(selection.x + selection.cx + 80000);
    expect(landed.y).toBe(selection.y - 64000);
    expect(landed.y + landed.cy).toBe(selection.y + selection.cy + 64000);
  });

  it("ignores the marker's authored size when it wraps", () => {
    // A marker that kept its authored size would never wrap anything: it would
    // sit on the shape rather than around it, which is the whole difference
    // between the wrap rule and the cursor rule.
    const selection: Rect = { x: 2000000, y: 2000000, cx: 1000000, cy: 800000 };
    const tiny = wrapping({ x: 0, y: 0, cx: 10, cy: 10 }, SLIDE, selection);
    const huge = wrapping({ x: 0, y: 0, cx: 9000000, cy: 5000000 }, SLIDE, selection);
    expect(tiny).toEqual(huge);
  });

  it("keeps the marker's authored size on a shape bigger than a third of the slide", () => {
    // "unless the shape is bigger than about a third of the slide, because ink
    // scales its stroke and a wrapped large box turns into a blob; then the
    // marker lands centred on the shape at its authored size".
    const selection: Rect = { x: 0, y: 0, cx: SLIDE.width / 2, cy: SLIDE.height };
    const landed = wrapping(MARKER, SLIDE, selection);
    expect({ cx: landed.cx, cy: landed.cy }).toEqual({ cx: MARKER.cx, cy: MARKER.cy });
    expect(centreOf(landed)).toEqual(centreOf(selection));
  });

  it("still wraps a shape that is exactly a third of the slide, because the rule is bigger-than", () => {
    // The prose says "bigger than about a third". Written `>=` the shape that
    // is exactly a third flips to the other rule, and the boundary is reachable
    // exactly: half the width by two thirds the height of a square slide.
    const selection: Rect = { x: 1000000, y: 1000000, cx: 3429000, cy: 4572000 };
    expect((selection.cx * selection.cy) / (SQUARE.width * SQUARE.height), "exactly a third").toBe(1 / 3);
    expect(wrapping(MARKER, SQUARE, selection)).toEqual({
      x: 1000000 - 274320,
      y: 1000000 - 365760,
      cx: 3429000 + 2 * 274320,
      cy: 4572000 + 2 * 365760,
    });
  });

  it("centres on the same shape one EMU taller", () => {
    // The other side of that boundary, one EMU across it, so the test says
    // which comparison is being made rather than merely that one is.
    const landed = wrapping(MARKER, SQUARE, { x: 1000000, y: 1000000, cx: 3429000, cy: 4572001 });
    expect({ cx: landed.cx, cy: landed.cy }).toEqual({ cx: MARKER.cx, cy: MARKER.cy });
  });

  it("wraps a wide shape and the same shape stood on end alike, because the threshold is an area", () => {
    // Compared by area, "because a wide thin banner and a tall thin column are
    // both cases where scaling one axis alone would mislead". Both of these are
    // 4000000 by 1000000 EMU about the slide's centre, one transposed, and both
    // are well under a third of the slide by area — while each is longer on its
    // long axis than a third of that axis, so a rule that compared extents
    // would centre both instead of wrapping them.
    const wide: Rect = { x: 1429000, y: 2929000, cx: 4000000, cy: 1000000 };
    const tall = transpose(wide);
    expect(tall).toEqual({ x: 2929000, y: 1429000, cx: 1000000, cy: 4000000 });
    expect(wide.cx, "longer than a third of the slide's width").toBeGreaterThan(SQUARE.width / 3);

    const wrappedWide = wrapping(MARKER, SQUARE, wide);
    const wrappedTall = wrapping(MARKER, SQUARE, tall);
    expect(wrappedWide, "wrapped, not centred").not.toMatchObject({ cx: MARKER.cx, cy: MARKER.cy });
    expect(transpose(wrappedWide), "and the transposed shape gets the transposed answer").toEqual(wrappedTall);
  });

  it("centres on a wide shape and on the same shape stood on end alike", () => {
    // The same pair above the threshold: 6000000 by 3000000 about the centre,
    // 38% of a square slide either way up. Same area, same verdict, and because
    // both sit on the slide's centre, the very same rectangle.
    const wide: Rect = { x: 429000, y: 1929000, cx: 6000000, cy: 3000000 };
    const tall = transpose(wide);
    const expected = { x: 2929000, y: 3129000, cx: MARKER.cx, cy: MARKER.cy };
    expect(wrapping(MARKER, SQUARE, wide)).toEqual(expected);
    expect(wrapping(MARKER, SQUARE, tall)).toEqual(expected);
  });

  it("is the cursor rule when nothing is selected, because there is nothing to wrap", () => {
    expect(wrapping(MARKER, SLIDE)).toEqual(atCursor(MARKER, SLIDE));
    expect(centreOf(wrapping(MARKER, SLIDE))).toEqual({ x: SLIDE.width / 2, y: SLIDE.height / 2 });
  });
});

describe("underTitle", () => {
  it("leaves an element that already fits under the title exactly where the library put it", () => {
    // "left where it sits in the library when it already fits", because an
    // element that fits is one the owner positioned deliberately. The same
    // OBJECT comes back, which is what lets `shapes.ts` skip the rewrite and
    // splice the markup in byte-for-byte as it was drawn.
    const rect: Rect = { x: 838200, y: 2000000, cx: 10515600, cy: 4000000 };
    const landed = underTitle(rect, SLIDE, { title: TITLE });
    expect(landed).toBe(rect);
    expect(isIdentity(moveFrom(rect, landed)), "so nothing is moved and nothing is scaled").toBe(true);
  });

  it("scales one that would overlap a taller title into the body area", () => {
    // The case the whole rule exists for: "a library authored for the SSF
    // layout would otherwise overlap a customer deck's taller title". Fits is
    // measured against the space below the title, not against the slide — this
    // element fits the slide perfectly well and still has to move.
    const frames: Frames = {
      title: { x: 838200, y: 365125, cx: 10515600, cy: 2500000 },
      body: { x: 838200, y: 3000000, cx: 10515600, cy: 3500000 },
    };
    const rect: Rect = { x: 838200, y: 1200000, cx: 10515600, cy: 5000000 };
    expect(onSlide(rect, SLIDE), "it fits the slide; it is the title it does not clear").toBe(true);

    const landed = underTitle(rect, SLIDE, frames);
    expect(landed).toEqual({ x: 2415540, y: 3000000, cx: 7360920, cy: 3500000 });
    expect(landed.y, "clear of the title").toBeGreaterThanOrEqual(2865125);
    expect(landed.cx / landed.cy, "scaled, not squashed").toBeCloseTo(rect.cx / rect.cy, 5);
  });

  it("measures the room from the title's bottom when the body placeholder starts above it", () => {
    // A deck whose body overlaps its own title is ordinary, and scaling into
    // the whole body would leave the element under the title anyway — the one
    // outcome the rule is named after.
    const frames: Frames = {
      title: { x: 838200, y: 365125, cx: 10515600, cy: 2500000 },
      body: { x: 838200, y: 1000000, cx: 10515600, cy: 5000000 },
    };
    const landed = underTitle({ x: 0, y: 0, cx: SLIDE.width, cy: SLIDE.height }, SLIDE, frames);
    expect(landed.y, "starts at the title's bottom, not at the body's top").toBe(2865125);
    expect(landed.cy, "and gets only what is left of the body below it").toBe(3134875);
    expect(landed.cx / landed.cy).toBeCloseTo(SLIDE.width / SLIDE.height, 5);
  });

  it("uses everything below the title when the slide has no body placeholder", () => {
    // The honest fallback rather than an invented margin: a blank layout, or
    // one the user has stripped, still gets an element that clears the title.
    const landed = underTitle({ x: 0, y: 0, cx: SLIDE.width, cy: SLIDE.height }, SLIDE, { title: TITLE });
    expect(landed.y).toBe(TITLE_BOTTOM);
    expect(landed.cy).toBe(SLIDE.height - TITLE_BOTTOM);
    expect(landed.x, "centred across the slide").toBe((SLIDE.width - landed.cx) / 2);
  });

  it("uses the body area when the slide has a body and no title", () => {
    // A title-less layout leaves the top at zero, and the body is still where
    // the deck itself says content goes.
    const body: Rect = { x: 1000000, y: 1000000, cx: 10000000, cy: 4000000 };
    const rect: Rect = { x: -1000000, y: 500000, cx: 14000000, cy: 5000000 };
    const landed = underTitle(rect, SLIDE, { body });
    expect(landed.x).toBeGreaterThanOrEqual(body.x);
    expect(landed.y).toBeGreaterThanOrEqual(body.y);
    expect(landed.x + landed.cx).toBeLessThanOrEqual(body.x + body.cx);
    expect(landed.y + landed.cy).toBeLessThanOrEqual(body.y + body.cy);
    expect(landed.cx / landed.cy).toBeCloseTo(rect.cx / rect.cy, 4);
  });

  it("slides an element that clears the title but hangs off the right edge back on, without shrinking it", () => {
    // It fits the room, so `fitInside` returns it untouched and the clamp does
    // the work. An element scaled here would lose size it never needed to lose.
    const rect: Rect = { x: 10000000, y: 2000000, cx: 4000000, cy: 1000000 };
    const landed = underTitle(rect, SLIDE, { title: TITLE });
    expect(landed).toEqual({ x: 8192000, y: 2000000, cx: 4000000, cy: 1000000 });
    expect(landed.x + landed.cx).toBe(SLIDE.width);
  });

  it("pulls one whose left edge is off the slide back to zero", () => {
    // The other half of the width test. Only `x + cx` being checked would let a
    // negative x through, and the library does hold shapes whose rotated extent
    // starts left of the slide.
    const landed = underTitle({ x: -1000000, y: 2000000, cx: 4000000, cy: 1000000 }, SLIDE, { title: TITLE });
    expect(landed).toEqual({ x: 0, y: 2000000, cx: 4000000, cy: 1000000 });
  });

  it("brings a bottom that hangs off the slide back up", () => {
    const landed = underTitle({ x: 0, y: 5000000, cx: 2000000, cy: 3000000 }, SLIDE, { title: TITLE });
    expect(landed).toEqual({ x: 0, y: 3858000, cx: 2000000, cy: 3000000 });
    expect(landed.y + landed.cy).toBe(SLIDE.height);
  });

  it("only keeps an element on the slide when a title leaves no room at all", () => {
    // A title placeholder covering the whole slide leaves a room of zero
    // height. Scaling into that gives a scale of zero and an element nobody can
    // see, so the rule gives up on the layout and merely keeps it on the slide.
    const landed = underTitle({ x: -100000, y: 100000, cx: 2000000, cy: 1000000 }, SLIDE, {
      title: { x: 0, y: 0, cx: SLIDE.width, cy: SLIDE.height },
    });
    expect(landed).toEqual({ x: 0, y: 100000, cx: 2000000, cy: 1000000 });
    expect(landed.cy, "still visible").toBeGreaterThan(0);
  });
});

describe("place", () => {
  const BOX = { x: 0.4, y: 0.4, w: 0.1, h: 0.1 };
  /** That box in EMU on the 16:9 slide, which is what every landing rule below starts from. */
  const RECT: Rect = { x: 4876800, y: 2743200, cx: 1219200, cy: 685800 };
  const SELECTION: Rect = { x: 2000000, y: 2000000, cx: 1000000, cy: 800000 };

  it("sends a stamp top-right", () => {
    expect(place({ box: BOX, landing: "top-right", slide: SLIDE })).toEqual(topRight(RECT, SLIDE));
    expect(place({ box: BOX, landing: "top-right", slide: SLIDE })).toEqual({
      x: 10850880,
      y: 68580,
      cx: 1219200,
      cy: 685800,
    });
  });

  it("sends a flowchart shape to the cursor", () => {
    expect(place({ box: BOX, landing: "cursor", slide: SLIDE, selection: SELECTION })).toEqual({
      x: 1890400,
      y: 2057100,
      cx: 1219200,
      cy: 685800,
    });
  });

  it("sends a marker around the selection instead, when the element wraps", () => {
    // The `wraps` flag is the only thing separating these two requests, and it
    // is what makes a marker a marker rather than another shape at the cursor.
    const marker = place({ box: BOX, landing: "cursor", slide: SLIDE, selection: SELECTION, wraps: true });
    expect(marker).toEqual({ x: 1920000, y: 1936000, cx: 1160000, cy: 928000 });
    expect(marker).not.toEqual(place({ box: BOX, landing: "cursor", slide: SLIDE, selection: SELECTION }));
  });

  it("sends a whole-slide element under the destination's title", () => {
    const frames: Frames = { title: TITLE };
    const whole = { x: 0, y: 0, w: 1, h: 1 };
    expect(place({ box: whole, landing: "layout", slide: SLIDE, frames })).toEqual(
      underTitle(authored(whole, SLIDE), SLIDE, frames),
    );
    expect(place({ box: whole, landing: "layout", slide: SLIDE, frames }).y).toBe(TITLE_BOTTOM);
  });

  it("does not need frames for a layout landing", () => {
    // A destination with no placeholders at all is a slide the user built by
    // hand, and it is common. Requiring frames would make it throw.
    const box = { x: 0.1, y: 0.1, w: 0.5, h: 0.5 };
    expect(place({ box, landing: "layout", slide: SLIDE })).toEqual(authored(box, SLIDE));
  });

  it("clamps an as-authored part onto the slide, which is not what as-authored means", () => {
    // "As authored" is about not RE-POSITIONING a part wider than half the
    // slide, not permission to hang off the edge. A box is a rotated shape's
    // rotated extent, so it can and does reach past 1.0 on the library slide.
    const landed = place({ box: { x: 0.9, y: 0.1, w: 0.3, h: 0.2 }, landing: "as-authored", slide: SLIDE });
    expect(landed).toEqual({ x: 8534400, y: 685800, cx: 3657600, cy: 1371600 });
    expect(landed.x + landed.cx, "brought flush rather than shrunk").toBe(SLIDE.width);
  });

  it("leaves an as-authored part that fits exactly where the library drew it", () => {
    const box = { x: 0.05, y: 0.8, w: 0.9, h: 0.1 };
    expect(place({ box, landing: "as-authored", slide: SLIDE })).toEqual(authored(box, SLIDE));
  });
});

describe("moveFrom and isIdentity", () => {
  it("describes the landing as a delta and a scale on each axis", () => {
    expect(moveFrom({ x: 1000, y: 2000, cx: 400, cy: 200 }, { x: 1500, y: 2600, cx: 800, cy: 200 })).toEqual({
      dx: 500,
      dy: 600,
      sx: 2,
      sy: 1,
    });
  });

  it("reports the scale per axis, so a caller that stretched is told it stretched", () => {
    // `fitInside` keeps proportions, so the two are equal whenever it was used.
    // Averaging them here would hide a deliberate stretch instead of describing
    // it, and the shapes would then be moved by a number nobody asked for.
    expect(moveFrom({ x: 0, y: 0, cx: 100, cy: 100 }, { x: 0, y: 0, cx: 300, cy: 150 })).toMatchObject({
      sx: 3,
      sy: 1.5,
    });
  });

  it("answers a scale of one for a zero-width authored box rather than infinity", () => {
    // A box is the union of real shapes so this should not come out of the
    // harvest, but `Infinity` reaching the file as a shape offset is a package
    // PowerPoint offers to repair, and the guard is one line.
    const move = moveFrom({ x: 0, y: 0, cx: 0, cy: 0 }, { x: 10, y: 20, cx: 500, cy: 400 });
    expect(move).toEqual({ dx: 10, dy: 20, sx: 1, sy: 1 });
    expect(Number.isFinite(move.sx) && Number.isFinite(move.sy)).toBe(true);
  });

  it("guards each axis separately", () => {
    // One guard covering both would answer 1 for the height of a box that has
    // a real height, silently refusing to scale it.
    expect(moveFrom({ x: 0, y: 0, cx: 0, cy: 100 }, { x: 0, y: 0, cx: 50, cy: 200 })).toEqual({
      dx: 0,
      dy: 0,
      sx: 1,
      sy: 2,
    });
    expect(moveFrom({ x: 0, y: 0, cx: 100, cy: 0 }, { x: 0, y: 0, cx: 200, cy: 50 })).toEqual({
      dx: 0,
      dy: 0,
      sx: 2,
      sy: 1,
    });
  });

  it("calls a move that changes nothing an identity", () => {
    expect(isIdentity({ dx: 0, dy: 0, sx: 1, sy: 1 })).toBe(true);
  });

  it("calls a move that changes any one thing not an identity", () => {
    // `shapes.ts` returns early on an identity and leaves the markup untouched,
    // so a dropped term in that `&&` chain is an element that silently refuses
    // to land — moved by the rule, unmoved in the file, and green throughout.
    for (const move of [
      { dx: 1, dy: 0, sx: 1, sy: 1 },
      { dx: 0, dy: 1, sx: 1, sy: 1 },
      { dx: 0, dy: 0, sx: 1.5, sy: 1 },
      { dx: 0, dy: 0, sx: 1, sy: 0.5 },
    ]) {
      expect(isIdentity(move), JSON.stringify(move)).toBe(false);
    }
  });

  it("reports an identity for the element the layout rule left alone, end to end", () => {
    // The promise the three functions make together: an element that fits under
    // the destination's title is spliced in exactly as the owner drew it.
    const box = { x: 0.1, y: 0.35, w: 0.8, h: 0.5 };
    const from = authored(box, SLIDE);
    const landed = place({ box, landing: "layout", slide: SLIDE, frames: { title: TITLE } });
    expect(from.y, "it does clear the title").toBeGreaterThanOrEqual(TITLE_BOTTOM);
    expect(isIdentity(moveFrom(from, landed))).toBe(true);
  });
});
