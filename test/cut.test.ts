/**
 * Where each element's preview is cut from its deck's print.
 *
 * The geometry only. Every rule `docs/DESIGN.md` section 3 states about the cut
 * is decided here rather than by looking at a rendered picture and judging it,
 * so the rasteriser that comes next has something to be wrong against.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AIR, cutFor, cutsFor, intersect, type Point, rotatedCorners, withAir } from "../src/core/catalogue/cut.js";
import type { Box, Element } from "../src/core/catalogue/types.js";

const el = (over: Partial<Element>): Element => ({
  id: "x",
  key: "x",
  name: "x",
  category: { key: "c", name: "C" },
  slide: 1,
  kind: "slide",
  box: { x: 0.2, y: 0.2, w: 0.4, h: 0.4 },
  landing: "layout",
  shapes: 1,
  tags: [],
  markup: { xml: "", rels: [], parts: [] },
  ...over,
});

const round = (b: Box, n = 4): Box => {
  const r = (v: number) => Math.round(v * 10 ** n) / 10 ** n;
  return { x: r(b.x), y: r(b.y), w: r(b.w), h: r(b.h) };
};

describe("air around the element", () => {
  it("grows the box by 3% of its own size on every side", () => {
    expect(round(withAir({ x: 0.2, y: 0.3, w: 0.4, h: 0.2 }))).toEqual(
      round({ x: 0.2 - 0.012, y: 0.3 - 0.006, w: 0.4 + 0.024, h: 0.2 + 0.012 }),
    );
    expect(AIR).toBe(0.03);
  });

  it("never reaches outside the page, however close to the edge the element sits", () => {
    const flush = withAir({ x: 0, y: 0, w: 1, h: 1 });
    expect(flush).toEqual({ x: 0, y: 0, w: 1, h: 1 });
    const corner = withAir({ x: 0.98, y: 0.98, w: 0.02, h: 0.02 });
    expect(corner.x + corner.w).toBeLessThanOrEqual(1);
    expect(corner.y + corner.h).toBeLessThanOrEqual(1);
    expect(corner.x).toBeGreaterThanOrEqual(0);
  });
});

describe("painting out the neighbours", () => {
  const a = el({ id: "a", kind: "part", slide: 9, box: { x: 0.1, y: 0.1, w: 0.2, h: 0.2 } });
  const touching = el({ id: "b", kind: "part", slide: 9, box: { x: 0.28, y: 0.1, w: 0.2, h: 0.2 } });
  const faraway = el({ id: "c", kind: "part", slide: 9, box: { x: 0.7, y: 0.7, w: 0.2, h: 0.2 } });
  const otherSlide = el({ id: "d", kind: "part", slide: 10, box: { x: 0.1, y: 0.1, w: 0.2, h: 0.2 } });

  it("paints out a part that intrudes on the crop, clipped to it", () => {
    const cut = cutFor(a, [a, touching, faraway, otherSlide]);
    expect(cut.whiteOut).toHaveLength(1);
    const [painted] = cut.whiteOut;
    // clipped: it may not reach past the crop it is painted onto
    expect(painted!.x).toBeGreaterThanOrEqual(cut.crop.x);
    expect(painted!.x + painted!.w).toBeLessThanOrEqual(cut.crop.x + cut.crop.w + 1e-9);
  });

  it("leaves a part on another slide alone, however much its box overlaps", () => {
    const cut = cutFor(a, [a, otherSlide]);
    expect(cut.whiteOut).toEqual([]);
  });

  it("paints nothing out for a whole-slide element, which IS the slide's content", () => {
    const whole = el({ id: "w", kind: "slide", slide: 9 });
    const cut = cutFor(whole, [whole, touching, faraway]);
    expect(cut.whiteOut).toEqual([]);
  });

  it("never paints out the element itself", () => {
    const cut = cutFor(a, [a, a]);
    expect(cut.whiteOut).toEqual([]);
  });

  it("never paints out a SEPARATE object carrying the element's own id", () => {
    // `cutFor` is called with an element and a list, and nothing makes the two
    // the same object: a caller that re-read the catalogue holds a copy. Identity
    // alone would let that copy paint the element out with its own box, which is
    // the whole tile. The id is what keeps the two apart.
    const copy = { ...a };
    expect(copy).not.toBe(a);
    const cut = cutFor(a, [copy]);
    expect(cut.whiteOut).toEqual([]);
  });
});

describe("masking a rotated part", () => {
  it("gives the frame's four corners turned about its centre", () => {
    // No page given, so this is the square-page case and a fraction-square is
    // a real square. The case below is the one that carries a real page.
    const square = { x: 0.25, y: 0.25, w: 0.5, h: 0.5 };
    const turned = rotatedCorners(square, 90);
    // a square turned a quarter turn is the same square, corners relabelled
    for (const p of turned) {
      expect(round(p as unknown as Box, 6).x).toBeGreaterThanOrEqual(0.25 - 1e-6);
      expect(p.x).toBeLessThanOrEqual(0.75 + 1e-6);
    }
    expect(turned).toHaveLength(4);
  });

  it("at rest gives the frame's own four corners, top-left first and clockwise", () => {
    // Every corner, not just the first: three of the four are otherwise free to
    // be built from the wrong half of the frame and no picture would say so.
    const frame = { x: 0.2, y: 0.3, w: 0.4, h: 0.2 };
    const at = (p: Point) => [Math.round(p.x * 1e6) / 1e6, Math.round(p.y * 1e6) / 1e6];
    expect(rotatedCorners(frame, 0).map(at)).toEqual([
      [0.2, 0.3],
      [0.6, 0.3],
      [0.6, 0.5],
      [0.2, 0.5],
    ]);
  });

  it("turns clockwise, the direction PowerPoint's rot counts in", () => {
    // A wide frame turned 90° clockwise puts its left-middle edge at the top.
    const wide = { x: 0, y: 0.4, w: 1, h: 0.2 };
    const [topLeft] = rotatedCorners(wide, 90);
    // the corner that was top-left (0, 0.4) swings to the top of the turned box
    expect(topLeft!.x).toBeCloseTo(0.6, 6);
    expect(topLeft!.y).toBeCloseTo(0, 6);
  });

  it("turns in PHYSICAL space, not in fractions of the page", () => {
    /**
     * The geometry here is fractions of the page, so an offset `(dx, dy)` is
     * `dx` of the WIDTH and `dy` of the HEIGHT — different distances on any
     * page that is not square, which is every page a print has. Turning that
     * pair with a plain rotation matrix is a shear.
     *
     * Measured 2026-09-22 on the 16:9 print's 960x540pt page, against the two
     * stamps the library actually carries: a 211x38pt frame at -29° had every
     * corner 23.5pt from where a rotated rectangle's corner belongs — more
     * than half that frame's height, against a `MASK_AIR` worth under 4pt. The
     * mask was cutting into the stamps it exists to frame.
     *
     * A fraction-SQUARE is the case that says it plainly: 0.2 by 0.2 is
     * physically wide, not square, so a quarter turn has to come back 0.2/a
     * wide and 0.2*a tall. Under the old spelling it came back 0.2 by 0.2, its
     * own shape, which is only true on a square page.
     */
    const a = 16 / 9;
    const frame = { x: 0.4, y: 0.4, w: 0.2, h: 0.2 };
    const turned = rotatedCorners(frame, 90, a);
    const xs = turned.map((p) => p.x);
    const ys = turned.map((p) => p.y);
    expect(Math.max(...xs) - Math.min(...xs)).toBeCloseTo(0.2 / a, 6);
    expect(Math.max(...ys) - Math.min(...ys)).toBeCloseTo(0.2 * a, 6);
    // Turned about its own centre, so the centre does not move.
    expect((Math.max(...xs) + Math.min(...xs)) / 2).toBeCloseTo(0.5, 6);
    expect((Math.max(...ys) + Math.min(...ys)) / 2).toBeCloseTo(0.5, 6);
  });

  it("is the plain rotation when the page IS square, so the default cannot drift", () => {
    // The pair: `aspect` defaults to 1, and at 1 the new spelling has to agree
    // with the old one exactly, or every caller that does not pass a page has
    // quietly changed.
    const frame = { x: 0.2, y: 0.3, w: 0.4, h: 0.2 };
    const at = (p: Point) => [Math.round(p.x * 1e6) / 1e6, Math.round(p.y * 1e6) / 1e6];
    expect(rotatedCorners(frame, 37).map(at)).toEqual(rotatedCorners(frame, 37, 1).map(at));
  });

  it("is absent when the element is not rotated, because the crop is the frame", () => {
    expect(cutFor(el({}), []).mask).toBeUndefined();
  });

  it("is present when it is", () => {
    const stamp = el({
      kind: "part",
      rotation: { deg: -29.056, frame: { x: 0.1923, y: 0.3017, w: 0.1655, h: 0.1155 } },
    });
    expect(cutFor(stamp, []).mask).toHaveLength(4);
  });

  it("leaves the element the same air the crop does, so a thick outline is not shaved", () => {
    // Masking to the BARE frame cut the ends off both stamps' ellipses in the
    // first prints taken with this: an outline is drawn centred on its path, so
    // it reaches past the frame. The mask removes what is in the CORNERS, not
    // part of the element.
    const frame = { x: 0.3, y: 0.4, w: 0.2, h: 0.1 };
    const stamp = el({ kind: "part", rotation: { deg: 0.0001, frame } });
    const mask = cutFor(stamp, []).mask!;
    const xs = mask.map((p) => p.x);
    const ys = mask.map((p) => p.y);
    expect(Math.min(...xs)).toBeLessThan(frame.x);
    expect(Math.max(...xs)).toBeGreaterThan(frame.x + frame.w);
    expect(Math.min(...ys)).toBeLessThan(frame.y);
    expect(Math.max(...ys)).toBeGreaterThan(frame.y + frame.h);
  });
});

describe("intersect", () => {
  it("is undefined for boxes that only touch at an edge", () => {
    expect(intersect({ x: 0, y: 0, w: 0.5, h: 1 }, { x: 0.5, y: 0, w: 0.5, h: 1 })).toBeUndefined();
  });

  it("is undefined for boxes that only touch along the horizontal edge too", () => {
    // The same rule on the other axis. Without it a part sitting exactly under
    // this one paints out a strip of no height, which is not a neighbour.
    expect(intersect({ x: 0, y: 0, w: 1, h: 0.5 }, { x: 0, y: 0.5, w: 1, h: 0.5 })).toBeUndefined();
  });
});

describe("the committed catalogue", () => {
  const catalogue = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
    sizes: Record<string, { elements: Element[]; width: number; height: number }>;
  };

  /** The page's own proportions, which is what `build-previews.mjs` passes. */
  const aspectOf = (size: "16:9" | "4:3"): number => catalogue.sizes[size]!.width / catalogue.sizes[size]!.height;

  for (const size of ["16:9", "4:3"] as const) {
    it(`${size}: every element gets a cut that stays on its page`, () => {
      const elements = catalogue.sizes[size]!.elements;
      const cuts = cutsFor(elements, undefined, aspectOf(size));
      expect(cuts).toHaveLength(elements.length);
      for (const cut of cuts) {
        expect(cut.page).toBeGreaterThanOrEqual(1);
        expect(cut.crop.w).toBeGreaterThan(0);
        expect(cut.crop.h).toBeGreaterThan(0);
        expect(cut.crop.x).toBeGreaterThanOrEqual(0);
        expect(cut.crop.y).toBeGreaterThanOrEqual(0);
        expect(cut.crop.x + cut.crop.w).toBeLessThanOrEqual(1 + 1e-9);
        expect(cut.crop.y + cut.crop.h).toBeLessThanOrEqual(1 + 1e-9);
      }
    });

    it(`${size}: only the two rotated stamps are masked`, () => {
      const masked = cutsFor(catalogue.sizes[size]!.elements, undefined, aspectOf(size)).filter((c) => c.mask);
      expect(masked.map((c) => c.id).sort()).toEqual(["confidential", "draft"]);
    });

    it(`${size}: a rotated stamp's mask is the size a turned rectangle actually is`, () => {
      /**
       * Derived independently of `rotatedCorners`, so this cannot agree with it
       * by sharing its arithmetic: a `w` by `h` rectangle turned by `d` has a
       * bounding box of `|w·cos d| + |h·sin d|` by `|w·sin d| + |h·cos d|`.
       * Everything is taken to EMU on the page first, because that is the space
       * the rotation is real in — which is the whole of the defect this pins.
       *
       * Under the fraction-space spelling this file used to hold, the 16:9
       * stamps came out with every corner 23.5pt from where a turned rectangle
       * puts one, on frames 38pt tall. The mask was cutting into the stamp.
       */
      const W = catalogue.sizes[size]!.width;
      const H = catalogue.sizes[size]!.height;
      const elements = catalogue.sizes[size]!.elements;
      const masked = cutsFor(elements, undefined, aspectOf(size)).filter((c) => c.mask);
      expect(masked.length, "the library's two rotated stamps").toBe(2);

      for (const cut of masked) {
        const source = elements.find((e) => e.id === cut.id)!;
        const turn = source.rotation!;
        const frame = withAir(turn.frame, 0.1);
        const t = (turn.deg * Math.PI) / 180;
        const wantW = Math.abs(frame.w * W * Math.cos(t)) + Math.abs(frame.h * H * Math.sin(t));
        const wantH = Math.abs(frame.w * W * Math.sin(t)) + Math.abs(frame.h * H * Math.cos(t));

        const xs = cut.mask!.map((p) => p.x * W);
        const ys = cut.mask!.map((p) => p.y * H);
        expect(Math.max(...xs) - Math.min(...xs), `${cut.id} width`).toBeCloseTo(wantW, 3);
        expect(Math.max(...ys) - Math.min(...ys), `${cut.id} height`).toBeCloseTo(wantH, 3);
      }
    });

    it(`${size}: the parts that share a collection slide paint each other out`, () => {
      const elements = catalogue.sizes[size]!.elements;
      const cuts = cutsFor(elements, undefined, aspectOf(size));
      const parts = elements.filter((e) => e.kind === "part");
      // Eleven since 2026-09-16: the ten Flowchart shapes were parts too, and
      // their slide went with the category.
      expect(parts.length).toBe(10);
      // at least some of them are close enough to intrude; if none were, the
      // white-out rule would be dead code and this test would be pinning nothing
      const withNeighbours = cuts.filter((c) => c.whiteOut.length > 0);
      expect(withNeighbours.length).toBeGreaterThan(0);
      for (const cut of cuts) {
        const source = elements.find((e) => e.id === cut.id)!;
        if (source.kind !== "part") expect(cut.whiteOut).toEqual([]);
      }
    });
  }
});
