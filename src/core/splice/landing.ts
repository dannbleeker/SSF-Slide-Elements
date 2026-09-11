/**
 * Where an element lands on the destination slide.
 *
 * `docs/DESIGN.md` section 5 is the prose this implements, and the whole of it
 * is arithmetic over rectangles: nothing here opens a package, reads a shape or
 * knows what an element is made of. That is deliberate. The landing is the part
 * of the splice a user sees and argues with — "it landed on top of my title" —
 * and a rule expressed as a pure function over boxes is one the suite can hold
 * to a number instead of a screenshot.
 *
 * Everything is EMU (English Metric Units, 914400 to the inch), which is what
 * the file uses. The catalogue's boxes are FRACTIONS of the library slide, so
 * `authored` is the one place the two meet.
 */

/** A rectangle in EMU, the spelling `<a:off>` and `<a:ext>` use between them. */
export interface Rect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/** A box in fractions of the slide: 0..1 on both axes, as the catalogue stores it. */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** The destination slide's size in EMU, from `<p:sldSz>`. */
export interface SlideSize {
  width: number;
  height: number;
}

/**
 * What the destination slide offers an element that has to fit around it.
 *
 * Both are optional because both are read from the slide's own placeholders,
 * and a slide can have neither — a blank layout, or one the user has stripped.
 * The rules below degrade to the slide itself rather than refusing, because an
 * element that lands somewhere reasonable is worth more than one that does not
 * land at all.
 */
export interface Frames {
  /** The title placeholder's rectangle, if the destination slide has a title. */
  title?: Rect;
  /** The body or content placeholder's rectangle, if it has one. */
  body?: Rect;
}

/**
 * How much of its own size a stamp keeps clear of the slide edge.
 *
 * `docs/DESIGN.md` section 5: a stamp lands top-right "with a 10% margin
 * because a rotated stamp's visible ellipse pokes past its frame". The margin
 * is a fraction of the ELEMENT, not of the slide, which is what makes it do
 * that job — the ink that pokes out is proportional to the stamp, so a margin
 * proportional to the slide would be too small on a big one and wasteful on a
 * small one. The owner's stamps are rotated 29° and 35°.
 */
const STAMP_MARGIN = 0.1;

/**
 * The air a marker leaves around the shape it wraps.
 *
 * A marker drawn tight to a box reads as a border rather than as ink around it.
 */
const WRAP_AIR = 0.08;

/**
 * How much of the slide a shape may cover and still be WRAPPED by a marker.
 *
 * `docs/DESIGN.md` section 5: "unless the shape is bigger than about a third of
 * the slide, because ink scales its stroke and a wrapped large box turns into a
 * blob; then the marker lands centred on the shape at its authored size."
 * Compared by area, because a wide thin banner and a tall thin column are both
 * cases where scaling one axis alone would mislead.
 */
const WRAP_LIMIT = 1 / 3;

/** The element's box as EMU on this slide: the one place fractions become units. */
export function authored(box: Box, slide: SlideSize): Rect {
  return {
    x: Math.round(box.x * slide.width),
    y: Math.round(box.y * slide.height),
    cx: Math.round(box.w * slide.width),
    cy: Math.round(box.h * slide.height),
  };
}

/** `rect` moved so its centre sits on `centre`, without resizing it. */
export function centredOn(rect: Rect, centre: { x: number; y: number }): Rect {
  return {
    x: Math.round(centre.x - rect.cx / 2),
    y: Math.round(centre.y - rect.cy / 2),
    cx: rect.cx,
    cy: rect.cy,
  };
}

/** The middle of a rectangle. */
export function centreOf(rect: Rect): { x: number; y: number } {
  return { x: rect.x + rect.cx / 2, y: rect.y + rect.cy / 2 };
}

/**
 * `rect` nudged until it is inside the slide, shrinking it only if it cannot fit.
 *
 * Every rule below ends here. A landing that puts half an element off the
 * canvas is the failure a user reports as "it disappeared", and each rule
 * separately guarding against it is how one of them ends up not doing so.
 *
 * Moved BEFORE it is shrunk, and that order is the whole of the behaviour: an
 * element that is merely in the wrong place keeps its authored proportions,
 * and only one that is genuinely too big for the slide is scaled — which is
 * the case a 4:3 element on a 16:9 slide reaches.
 */
export function ontoSlide(rect: Rect, slide: SlideSize): Rect {
  let { x, y, cx, cy } = rect;
  if (cx > slide.width) {
    cy = Math.round((cy * slide.width) / cx);
    cx = slide.width;
  }
  if (cy > slide.height) {
    cx = Math.round((cx * slide.height) / cy);
    cy = slide.height;
  }
  x = Math.min(Math.max(0, x), slide.width - cx);
  y = Math.min(Math.max(0, y), slide.height - cy);
  return { x, y, cx, cy };
}

/** `rect` scaled to fit inside `into`, keeping its proportions, never enlarged. */
export function fitInside(rect: Rect, into: Rect): Rect {
  if (rect.cx <= into.cx && rect.cy <= into.cy) return rect;
  // The SMALLER of the two ratios, so the result fits on both axes. Taking each
  // axis separately is how an element comes back stretched.
  const scale = Math.min(into.cx / rect.cx, into.cy / rect.cy);
  const cx = Math.round(rect.cx * scale);
  const cy = Math.round(rect.cy * scale);
  return {
    x: Math.round(into.x + (into.cx - cx) / 2),
    y: Math.round(into.y + (into.cy - cy) / 2),
    cx,
    cy,
  };
}

/** A stamp or a label, top-right, clear of the edge by a tenth of its own size. */
export function topRight(rect: Rect, slide: SlideSize): Rect {
  const inset = { x: Math.round(rect.cx * STAMP_MARGIN), y: Math.round(rect.cy * STAMP_MARGIN) };
  return ontoSlide({ ...rect, x: slide.width - rect.cx - inset.x, y: inset.y }, slide);
}

/**
 * A marker, a flowchart shape or an icon: at the cursor.
 *
 * An add-in cannot see the mouse on the canvas, so "at the cursor" is the
 * selected shape's position and, with nothing selected, the middle of the
 * slide. `docs/DESIGN.md` section 5 says so, and says why: `getSelectedShapes`
 * is read-only and the siblings measured it safe, while `setSelectedShapes`
 * wedges the web host's selection subsystem outright.
 */
export function atCursor(rect: Rect, slide: SlideSize, selection?: Rect): Rect {
  const centre = selection ? centreOf(selection) : { x: slide.width / 2, y: slide.height / 2 };
  return ontoSlide(centredOn(rect, centre), slide);
}

/**
 * A marker with a shape selected: around it.
 *
 * Sized to the shape with a little air and centred on it — unless the shape is
 * more than about a third of the slide, when the marker keeps its authored size
 * and is merely centred. The reason is in `docs/DESIGN.md` section 5: ink
 * scales its stroke, so a marker stretched around a large box stops looking
 * like a pen mark and turns into a blob.
 *
 * With nothing selected there is nothing to wrap, and this is the same rule as
 * `atCursor`.
 */
export function wrapping(rect: Rect, slide: SlideSize, selection?: Rect): Rect {
  if (!selection) return atCursor(rect, slide);
  const share = (selection.cx * selection.cy) / (slide.width * slide.height);
  if (share > WRAP_LIMIT) return ontoSlide(centredOn(rect, centreOf(selection)), slide);
  const air = { x: Math.round(selection.cx * WRAP_AIR), y: Math.round(selection.cy * WRAP_AIR) };
  return ontoSlide(
    {
      x: selection.x - air.x,
      y: selection.y - air.y,
      cx: selection.cx + air.x * 2,
      cy: selection.cy + air.y * 2,
    },
    slide,
  );
}

/**
 * A whole-slide element, placed against the destination slide's own layout.
 *
 * The library is authored for the SSF layout, and a customer deck's title is
 * routinely taller. Left where it sits, a whole-slide element would slide under
 * that title. So it lands below whatever title the destination has, scaled into
 * the body area when it would not otherwise fit, and is left exactly where the
 * library put it when it already fits — because an element that fits is one the
 * owner positioned deliberately, and moving it would be this function
 * second-guessing the deck.
 *
 * **"Fits" is measured against the ROOM, not against the slide**, and those are
 * different rectangles: an element that fits the slide and overlaps the title is
 * the whole case this exists for. An earlier version said this in its comment
 * and tested `rect.y + rect.cy <= slide.height` in its code, which let an
 * element overhang the room by millions of EMU and be returned untouched.
 *
 * The room is the space below the title, across the WHOLE slide — not the body
 * placeholder, even when the destination has one. A two-content layout's body is
 * one COLUMN, and squeezing a whole-slide element into the left column of
 * somebody's comparison slide is a worse answer than letting it use the width it
 * was drawn at. The body is used only when it is at least half the slide wide,
 * which is what tells a content area from a column.
 *
 * On the fits path this returns the CALLER'S OWN rectangle rather than a copy,
 * and that is deliberate: `moveFrom` then produces an exactly identity move and
 * `applyMove` touches nothing, so an element that lands where it was authored is
 * spliced in byte for byte. Nothing may write through the returned object.
 */
export function underTitle(rect: Rect, slide: SlideSize, frames: Frames): Rect {
  const top = frames.title ? frames.title.y + frames.title.cy : 0;
  const wide = frames.body && frames.body.cx >= slide.width / 2 ? frames.body : undefined;
  const area: Rect = wide ?? { x: 0, y: top, cx: slide.width, cy: slide.height - top };
  const y = Math.max(area.y, top);
  const room: Rect = { x: area.x, y, cx: area.cx, cy: area.y + area.cy - y };
  if (room.cy <= 0 || room.cx <= 0) return ontoSlide(rect, slide);
  const inside =
    rect.x >= room.x &&
    rect.y >= room.y &&
    rect.x + rect.cx <= room.x + room.cx &&
    rect.y + rect.cy <= room.y + room.cy;
  if (inside) return rect;
  return ontoSlide(fitInside(rect, room), slide);
}

/** Which rule an element's landing takes, before any of them is applied. */
export type Landing = "layout" | "top-right" | "cursor" | "as-authored";

export interface PlaceRequest {
  box: Box;
  landing: Landing;
  slide: SlideSize;
  frames?: Frames;
  /** The selected shape's rectangle, when the host could name one. */
  selection?: Rect;
  /** True for a marker, which wraps the selection rather than merely sitting on it. */
  wraps?: boolean;
}

/**
 * The one entry point: an element's box and its landing rule, in, a rectangle
 * on the destination slide, out.
 *
 * A switch rather than a table of functions, so a landing value with no rule is
 * a type error at the call site rather than an undefined at run time.
 */
export function place(request: PlaceRequest): Rect {
  const rect = authored(request.box, request.slide);
  switch (request.landing) {
    case "top-right":
      return topRight(rect, request.slide);
    case "cursor":
      return request.wraps
        ? wrapping(rect, request.slide, request.selection)
        : atCursor(rect, request.slide, request.selection);
    case "layout":
      return underTitle(rect, request.slide, request.frames ?? {});
    case "as-authored":
      // Deliberately still clamped. "As authored" is about not RE-POSITIONING a
      // part that is wider than half the slide, not about permission to hang
      // off the edge of a slide whose size differs from the library's.
      return ontoSlide(rect, request.slide);
  }
}

/**
 * The transform that moves an element's shapes from where they were authored to
 * where they are landing.
 *
 * Returned rather than applied, because the shapes are XML and this file does
 * not touch XML. Scale is per axis: `fitInside` keeps proportions, so the two
 * are equal whenever it was used, and a caller that stretches deliberately
 * still gets an honest description of what it asked for.
 */
export interface Move {
  dx: number;
  dy: number;
  sx: number;
  sy: number;
}

export function moveFrom(from: Rect, to: Rect): Move {
  return {
    dx: to.x - from.x,
    dy: to.y - from.y,
    // A zero-width authored box would divide by zero. It cannot come out of the
    // harvest — a box is the union of real shapes — but the guard is one line
    // and the alternative is `Infinity` reaching the file as a shape offset.
    sx: from.cx === 0 ? 1 : to.cx / from.cx,
    sy: from.cy === 0 ? 1 : to.cy / from.cy,
  };
}

/** Whether a move actually changes anything, so an untouched element stays byte-identical. */
export function isIdentity(move: Move): boolean {
  return move.dx === 0 && move.dy === 0 && move.sx === 1 && move.sy === 1;
}
