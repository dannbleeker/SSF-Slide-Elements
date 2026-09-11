/**
 * Where a shape is on its slide, as a box in fractions of the slide.
 *
 * Two things the XML's frame does not say, both learned against the owner's
 * print (`docs/DESIGN.md` section 3): a rotated shape's frame is the UNROTATED
 * box, so the box here is the rotated extent; and a table's frame is narrower
 * than the table PowerPoint draws, so a table's box is the sum of its columns
 * and rows when that is larger than the frame.
 */
import { A_NS, P_NS, element, elements, children } from "../pptx/xml.js";
import { placeholderType, textOf } from "./text.js";
import type { Box } from "./types.js";

interface Frame {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Degrees, clockwise. */
  rot: number;
}

function numberAttr(el: Element | undefined, name: string): number | undefined {
  const v = el?.getAttribute(name);
  if (v === null || v === undefined || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** The `<a:xfrm>` (or `<p:xfrm>` on a graphic frame) that positions a top-level shape. */
function frameOf(shape: Element): Frame | undefined {
  const local = shape.localName;
  let xfrm: Element | undefined;
  if (local === "grpSp") {
    const grpSpPr = element(shape, P_NS, "grpSpPr");
    xfrm = grpSpPr ? element(grpSpPr, A_NS, "xfrm") : undefined;
  } else if (local === "graphicFrame") {
    xfrm = children(shape, P_NS, "xfrm")[0];
  } else {
    const spPr = element(shape, P_NS, "spPr");
    xfrm = spPr ? element(spPr, A_NS, "xfrm") : undefined;
  }
  if (!xfrm) return undefined;
  const off = element(xfrm, A_NS, "off");
  const ext = element(xfrm, A_NS, "ext");
  const x = numberAttr(off, "x");
  const y = numberAttr(off, "y");
  const w = numberAttr(ext, "cx");
  const h = numberAttr(ext, "cy");
  if (x === undefined || y === undefined || w === undefined || h === undefined) return undefined;
  return { x, y, w, h, rot: (numberAttr(xfrm, "rot") ?? 0) / 60000 };
}

/** A table's drawn size: its columns and rows added up. Undefined when the frame holds no table. */
function tableSize(shape: Element): { w: number; h: number } | undefined {
  const tbl = element(shape, A_NS, "tbl");
  if (!tbl) return undefined;
  const grid = element(tbl, A_NS, "tblGrid");
  const w = grid ? children(grid, A_NS, "gridCol").reduce((sum, col) => sum + (numberAttr(col, "w") ?? 0), 0) : 0;
  const h = children(tbl, A_NS, "tr").reduce((sum, tr) => sum + (numberAttr(tr, "h") ?? 0), 0);
  return { w, h };
}

/**
 * The box of a top-level shape in fractions of a `width` x `height` slide,
 * rotation and tables accounted for. Undefined for a shape without a frame.
 */
export function boxOf(shape: Element, width: number, height: number): Box | undefined {
  const f = frameOf(shape);
  if (!f) return undefined;
  let { w, h } = f;
  const table = tableSize(shape);
  if (table) {
    w = Math.max(w, table.w);
    h = Math.max(h, table.h);
  }
  let x = f.x;
  let y = f.y;
  if (f.rot % 180 !== 0) {
    const t = (f.rot * Math.PI) / 180;
    const cx = f.x + f.w / 2;
    const cy = f.y + f.h / 2;
    const rw = Math.abs(w * Math.cos(t)) + Math.abs(h * Math.sin(t));
    const rh = Math.abs(w * Math.sin(t)) + Math.abs(h * Math.cos(t));
    x = cx - rw / 2;
    y = cy - rh / 2;
    w = rw;
    h = rh;
  }
  return { x: x / width, y: y / height, w: w / width, h: h / height };
}

/** The smallest box around several. */
export function union(boxes: Box[]): Box {
  const first = boxes[0];
  if (!first) return { x: 0, y: 0, w: 1, h: 1 };
  let x0 = first.x;
  let y0 = first.y;
  let x1 = first.x + first.w;
  let y1 = first.y + first.h;
  for (const b of boxes) {
    x0 = Math.min(x0, b.x);
    y0 = Math.min(y0, b.y);
    x1 = Math.max(x1, b.x + b.w);
    y1 = Math.max(y1, b.y + b.h);
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/**
 * A rotated shape's UNROTATED frame and its angle, in fractions of the slide.
 * Undefined when the shape is not rotated, or has no frame at all.
 *
 * `boxOf` gives the rotated EXTENT, which is what the pane needs to know how
 * much room an element takes. The preview cut needs the other thing: the frame
 * the extent was computed from, so everything outside the rotated rectangle can
 * be masked away rather than showing whatever the slide has in the corners
 * (`docs/DESIGN.md` section 3). The owner's stamps are rotated 29° and 35°.
 */
export function rotationOf(shape: Element, width: number, height: number): { deg: number; frame: Box } | undefined {
  const f = frameOf(shape);
  if (!f || f.rot % 360 === 0) return undefined;
  let { w, h } = f;
  const table = tableSize(shape);
  if (table) {
    w = Math.max(w, table.w);
    h = Math.max(h, table.h);
  }
  return { deg: f.rot, frame: { x: f.x / width, y: f.y / height, w: w / width, h: h / height } };
}

/** A box rounded to four decimals, so the committed catalogue does not churn on floating-point noise. */
export function rounded(box: Box): Box {
  const r = (v: number) => Math.round(v * 10000) / 10000;
  return { x: r(box.x), y: r(box.y), w: r(box.w), h: r(box.h) };
}

/** True when the whole shape sits beyond the slide's right or bottom edge: the 4:3 deck's instruction boxes. */
export function offSlide(box: Box): boolean {
  return box.x >= 1 || box.y >= 1;
}

/** Every top-level shape of a slide's tree, in z-order, skipping the tree's own two property children. */
export function topLevelShapes(slide: Document): Element[] {
  const tree = elements(slide, P_NS, "spTree")[0];
  if (!tree) return [];
  const out: Element[] = [];
  for (let i = 0; i < tree.childNodes.length; i++) {
    const node = tree.childNodes.item(i);
    if (!node || node.nodeType !== 1) continue;
    const el = node as Element;
    if (["sp", "grpSp", "pic", "cxnSp", "graphicFrame", "AlternateContent"].includes(el.localName)) out.push(el);
  }
  return out;
}

/**
 * What a slide already holds, as boxes the preview card can draw in grey.
 *
 * `docs/DESIGN.md` sections 1 and 4: the card shows a small slide with where
 * the element will land, **next to what the slide already has**. This answers
 * the second half, read out of the FILE — the same route "Used in this deck"
 * takes, and for the same reason: it needs no host capability beyond the deck
 * read the pane already does, so nothing here rests on an unmeasured API.
 *
 * What is left out, and why each one:
 *
 * - **A shape with no frame of its own.** A placeholder inheriting its geometry
 *   from the layout says nothing about where it is, and `boxOf` answers
 *   undefined rather than guessing. Drawing it at the origin would be a lie in
 *   the exact place the user is looking for one.
 * - **A shape entirely off the slide.** The library decks carry hundreds of
 *   those as authoring notes; a destination deck can carry them too, and they
 *   are not on the slide the user can see.
 * - **An EMPTY placeholder.** The insert removes the "Click to add text" ghosts
 *   it lands over (section 6), so drawing them as occupied would show the user
 *   an obstacle the insert is about to take away. A placeholder with text or a
 *   picture in it is content and stays.
 *
 * In z-order, which is the order they are drawn in, so a caller painting them
 * in sequence gets the same stacking PowerPoint would.
 */
export function occupiedBoxes(slide: Document, width: number, height: number): Box[] {
  const out: Box[] = [];
  for (const shape of topLevelShapes(slide)) {
    const box = boxOf(shape, width, height);
    if (!box || offSlide(box)) continue;
    if (isEmptyPlaceholder(shape)) continue;
    out.push(rounded(box));
  }
  return out;
}

/** A placeholder with nothing in it: the ghost an insert removes rather than lands on. */
function isEmptyPlaceholder(shape: Element): boolean {
  if (placeholderType(shape) === undefined) return false;
  const hasText = textOf(shape).length > 0;
  const hasGraphic = elements(shape, A_NS, "graphic").length > 0 || elements(shape, A_NS, "blip").length > 0;
  return !(hasText || hasGraphic);
}
