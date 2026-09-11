/**
 * Where each element's preview is cut from its deck's print.
 *
 * `docs/DESIGN.md` section 3: "A part is cut to its box with 3% of air, the
 * boxes of neighbouring parts painted white, and a rotated part masked to its
 * rotated frame."
 *
 * This is the geometry only — rectangles and a polygon, in fractions of the
 * page. Nothing here opens a PDF or writes an image, so every rule below is
 * decided in a unit test rather than by looking at a picture and guessing.
 *
 * Fractions rather than points because a print's page is the slide: the same
 * cut describes the 16:9 page at 960x540pt and the 4:3 page at 720x540pt, and
 * the rasteriser multiplies by whatever the page actually measures.
 */
import type { Box, Element } from "./types.js";

/** How much room to leave around an element, as a fraction of its own size. */
export const AIR = 0.03;

export interface Point {
  x: number;
  y: number;
}

/** One element's preview, as instructions against a page of the print. */
export interface Cut {
  /** The element this is a preview of. */
  id: string;
  /** 1-based page of the print, which is the element's slide. */
  page: number;
  /** What to take, in fractions of the page. */
  crop: Box;
  /**
   * What to paint out first, in fractions of the page, already clipped to the
   * crop. A neighbouring part that shares a collection slide would otherwise
   * appear in the corner of this one's picture.
   */
  whiteOut: Box[];
  /**
   * For a rotated part, its rotated frame as a four-point polygon in fractions
   * of the page. Everything OUTSIDE it is not this element. Absent when the
   * element is not rotated, where the crop is the frame.
   */
  mask?: Point[];
}

const clamp01 = (n: number): number => (n < 0 ? 0 : n > 1 ? 1 : n);

/** A box grown by a fraction of its own size on all four sides, clipped to the page. */
export function withAir(box: Box, air = AIR): Box {
  const dx = box.w * air;
  const dy = box.h * air;
  const x = clamp01(box.x - dx);
  const y = clamp01(box.y - dy);
  return {
    x,
    y,
    w: clamp01(box.x + box.w + dx) - x,
    h: clamp01(box.y + box.h + dy) - y,
  };
}

/** The overlap of two boxes, or undefined when they do not touch. */
export function intersect(a: Box, b: Box): Box | undefined {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.w, b.x + b.w);
  const bottom = Math.min(a.y + a.h, b.y + b.h);
  if (right <= x || bottom <= y) return undefined;
  return { x, y, w: right - x, h: bottom - y };
}

/**
 * The four corners of a frame rotated about its own centre.
 *
 * `rot` is degrees clockwise, the direction PowerPoint's `<a:xfrm rot>` counts
 * in once it is divided by 60000. The y axis runs down the page, so a clockwise
 * rotation on screen is the ordinary positive direction here.
 */
export function rotatedCorners(frame: Box, rot: number): Point[] {
  const t = (rot * Math.PI) / 180;
  const cos = Math.cos(t);
  const sin = Math.sin(t);
  const cx = frame.x + frame.w / 2;
  const cy = frame.y + frame.h / 2;
  const corners: [number, number][] = [
    [-frame.w / 2, -frame.h / 2],
    [frame.w / 2, -frame.h / 2],
    [frame.w / 2, frame.h / 2],
    [-frame.w / 2, frame.h / 2],
  ];
  return corners.map(([dx, dy]) => ({ x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos }));
}

/**
 * The cut for one element, given every element harvested from the same deck.
 *
 * Only a PART gets neighbours painted out. A whole-slide element IS the slide's
 * content, so the things around it are the layout chrome the crop already
 * excludes, and painting over another element's box would white out its own.
 */
export function cutFor(element: Element, all: readonly Element[], air = AIR): Cut {
  const crop = withAir(element.box, air);

  const whiteOut: Box[] = [];
  if (element.kind === "part") {
    for (const other of all) {
      if (other === element || other.id === element.id) continue;
      if (other.kind !== "part" || other.slide !== element.slide) continue;
      const overlap = intersect(other.box, crop);
      if (overlap) whiteOut.push(overlap);
    }
  }

  const cut: Cut = { id: element.id, page: element.slide, crop, whiteOut };
  if (element.rotation) {
    cut.mask = rotatedCorners(element.rotation.frame, element.rotation.deg);
  }
  return cut;
}

/** The cut for every element in a catalogue, in catalogue order. */
export function cutsFor(elements: readonly Element[], air = AIR): Cut[] {
  return elements.map((e) => cutFor(e, elements, air));
}
