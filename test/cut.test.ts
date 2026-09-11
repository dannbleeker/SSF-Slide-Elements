/**
 * Where each element's preview is cut from its deck's print.
 *
 * The geometry only. Every rule `docs/DESIGN.md` section 3 states about the cut
 * is decided here rather than by looking at a rendered picture and judging it,
 * so the rasteriser that comes next has something to be wrong against.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AIR, cutFor, cutsFor, intersect, rotatedCorners, withAir } from "../src/core/catalogue/cut.js";
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
});

describe("masking a rotated part", () => {
  it("gives the frame's four corners turned about its centre", () => {
    const square = { x: 0.25, y: 0.25, w: 0.5, h: 0.5 };
    const turned = rotatedCorners(square, 90);
    // a square turned a quarter turn is the same square, corners relabelled
    for (const p of turned) {
      expect(round(p as unknown as Box, 6).x).toBeGreaterThanOrEqual(0.25 - 1e-6);
      expect(p.x).toBeLessThanOrEqual(0.75 + 1e-6);
    }
    expect(turned).toHaveLength(4);
  });

  it("turns clockwise, the direction PowerPoint's rot counts in", () => {
    // A wide frame turned 90° clockwise puts its left-middle edge at the top.
    const wide = { x: 0, y: 0.4, w: 1, h: 0.2 };
    const [topLeft] = rotatedCorners(wide, 90);
    // the corner that was top-left (0, 0.4) swings to the top of the turned box
    expect(topLeft!.x).toBeCloseTo(0.6, 6);
    expect(topLeft!.y).toBeCloseTo(0, 6);
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
});

describe("the committed catalogue", () => {
  const catalogue = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
    sizes: Record<string, { elements: Element[] }>;
  };

  for (const size of ["16:9", "4:3"] as const) {
    it(`${size}: every element gets a cut that stays on its page`, () => {
      const elements = catalogue.sizes[size]!.elements;
      const cuts = cutsFor(elements);
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
      const masked = cutsFor(catalogue.sizes[size]!.elements).filter((c) => c.mask);
      expect(masked.map((c) => c.id).sort()).toEqual(["confidential", "draft"]);
    });

    it(`${size}: the parts that share a collection slide paint each other out`, () => {
      const elements = catalogue.sizes[size]!.elements;
      const cuts = cutsFor(elements);
      const parts = elements.filter((e) => e.kind === "part");
      expect(parts.length).toBe(21);
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
