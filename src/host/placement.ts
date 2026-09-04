/**
 * Where an element goes on a slide that already has something on it.
 *
 * The naive answer — put it where it was authored — is right surprisingly
 * often, because the library's elements were drawn in the content area of a
 * slide with the same proportions. It is wrong in exactly one way that matters:
 * dropping a second element on top of the first, so the user sees nothing
 * happen and clicks again.
 *
 * So: authored position when it is free, and otherwise the nearest free place
 * below it. Nothing cleverer. An element that lands somewhere slightly wrong is
 * dragged in two seconds; an element that lands invisibly under another one is
 * a bug report.
 */
import type { Box } from "../core/catalogue/types.js";

/** A slide's usable area, in EMU. */
export interface Canvas {
  cx: number;
  cy: number;
}

/** How much of two boxes overlap, as a fraction of the smaller one's area. */
export function overlapFraction(a: Box, b: Box): number {
  const w = Math.min(a.x + a.cx, b.x + b.cx) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.cy, b.y + b.cy) - Math.max(a.y, b.y);
  if (w <= 0 || h <= 0) return 0;
  const smaller = Math.min(a.cx * a.cy, b.cx * b.cy);
  if (smaller <= 0) return 0;
  return (w * h) / smaller;
}

/**
 * The scale that keeps an element inside a canvas it was not authored for.
 *
 * 1 whenever it already fits, which is the common case: the library is 16:9 and
 * so is almost every deck. A 16:9 element in a 4:3 deck is the case this exists
 * for, and it scales UNIFORMLY — never to fit each axis separately. A stretched
 * element is worse than a small one, because it is wrong in a way that looks
 * deliberate.
 */
export function fitScale(bounds: Box, canvas: Canvas, margin = 0): number {
  const w = canvas.cx - margin * 2;
  const h = canvas.cy - margin * 2;
  if (bounds.cx <= w && bounds.cy <= h) return 1;
  if (bounds.cx <= 0 || bounds.cy <= 0) return 1;
  return Math.min(w / bounds.cx, h / bounds.cy);
}

export interface PlacementChoice {
  x: number;
  y: number;
  scale: number;
  /** What was decided and why, for the pane to show and the log to keep. */
  reason: "as-authored" | "nudged-clear" | "scaled-to-fit" | "centred";
}

/**
 * Choose where to put an element.
 *
 * `occupied` is what is already on the slide, in slide coordinates. The library
 * deck's own furniture is not in it — the caller passes what it read from the
 * TARGET slide, and a title placeholder there is content as far as this is
 * concerned: landing an element across somebody's title is exactly the collision
 * worth avoiding.
 */
export function choosePlacement(bounds: Box, canvas: Canvas, occupied: Box[], margin = 114300): PlacementChoice {
  const scale = fitScale(bounds, canvas, margin);
  const cx = bounds.cx * scale;
  const cy = bounds.cy * scale;

  const at = (x: number, y: number): Box => ({ x, y, cx, cy });
  // Anything under a fifth covered reads as "beside", not "on top of". Below
  // that threshold a nudge moves an element away from where its author put it
  // to solve a problem nobody would have noticed.
  const clear = (box: Box): boolean => occupied.every((o) => overlapFraction(box, o) < 0.2);

  const authored = at(bounds.x, scale === 1 ? bounds.y : bounds.y * scale);
  if (scale === 1 && clear(authored)) return { x: authored.x, y: authored.y, scale, reason: "as-authored" };

  // Down the slide in tenths, then back up: an element added to a slide whose
  // top half is taken belongs below it, and one added to a full slide belongs
  // wherever there is most room rather than wherever the loop happened to stop.
  const step = Math.max(1, Math.round(canvas.cy / 10));
  for (let y = margin; y + cy <= canvas.cy - margin; y += step) {
    const box = at(authored.x, y);
    if (clear(box)) {
      return { x: box.x, y: box.y, scale, reason: scale === 1 ? "nudged-clear" : "scaled-to-fit" };
    }
  }

  // Nowhere free. Centred is the honest answer: the user is going to move it,
  // and the middle is where they will look for it.
  return {
    x: Math.round((canvas.cx - cx) / 2),
    y: Math.round((canvas.cy - cy) / 2),
    scale,
    reason: scale === 1 ? "centred" : "scaled-to-fit",
  };
}
