/**
 * The preview card: where an element will land, and what is already there.
 *
 * Split out of `steps.ts` on 2026-09-12. The card is one thing a reader looks
 * for by name and its parts were scattered: `occupiedFor` decides which grey
 * boxes to draw, `fractionOf` and `withLanded` turn what the splice reported
 * into a rectangle on the card's little slide, `slideLine` says which slide is
 * being described, and `landingLine` is the sentence under the drawing.
 *
 * The rectangle is in fractions of the USER'S slide, never the library's — a
 * deck that borrowed the nearest library has a different size, and dividing by
 * the wrong one draws the right rectangle in the wrong place. `fractionOf`
 * carries that rule and its history.
 */
import type { Box, Element } from "../core/catalogue/types.js";
import type { PaneState, Settings } from "./steps.js";

/**
 * The boxes the preview card should draw in grey behind the landing.
 *
 * Only when the snapshot is of the slide the user is ON. A card showing slide
 * two's furniture while the user is on slide five is worse than a card showing
 * none: the whole point of the boxes is to answer "will this land on top of
 * something", and an answer about another slide is a wrong answer rather than a
 * missing one.
 */
export function occupiedFor(state: PaneState): Box[] {
  if (!state.onSlide || state.slide === undefined) return [];
  return state.onSlide.slide === state.slide ? state.onSlide.boxes : [];
}

/**
 * A rectangle on the user's slide, in EMU, as a fraction of that slide.
 *
 * The size is the USER's, never the library's, and the two are different
 * numbers exactly when a deck borrowed the nearest library — A4, 16:10,
 * anything custom. The splice reports where an element landed in the
 * destination deck's EMU; dividing that by the library deck's size draws the
 * right rectangle in the wrong place, and only ever on the decks nobody tests
 * on. Answers undefined when the size is not known, because a fraction
 * measured against a size nobody read is a guess with a decimal point.
 */
export function fractionOf(
  landed: { x: number; y: number; cx: number; cy: number },
  deck: { width: number; height: number } | undefined,
): Box | undefined {
  if (!deck || !(deck.width > 0) || !(deck.height > 0)) return undefined;
  return {
    x: landed.x / deck.width,
    y: landed.y / deck.height,
    w: landed.cx / deck.width,
    h: landed.cy / deck.height,
  };
}

/**
 * The snapshot of what a slide holds, with an insert that just landed in it.
 *
 * The pane does not re-read the deck after an insert — it already knows what it
 * put where, and the read is the expensive thing (section 13's sixth open
 * question). So the box the splice reports is added to the snapshot instead.
 *
 * `blank` is the "as a new slide" case: that slide is a clone with its
 * placeholders emptied, and an empty placeholder is not in these boxes anyway,
 * so the element IS what the slide holds. Onto an existing slide the snapshot
 * only grows if it was already of THAT slide — extending a snapshot of some
 * other slide would invent an answer, and answering nothing is what the card is
 * built to survive.
 */
export function withLanded(
  onSlide: PaneState["onSlide"],
  slide: number,
  box: Box | undefined,
  blank: boolean,
): PaneState["onSlide"] {
  // No box means the pane never learned the user's slide size, so it cannot
  // turn EMU into a fraction. The snapshot goes rather than gaining a rectangle
  // measured against a size nobody read.
  if (!box) return undefined;
  if (blank) return { slide, boxes: [box] };
  if (onSlide?.slide !== slide) return undefined;
  return { slide, boxes: [...onSlide.boxes, box] };
}

/** Which slide an insert would land on, as the pane says it. */
export function slideLine(state: PaneState): string {
  return state.slide === undefined
    ? "PowerPoint did not say which slide you are on, so an element will land on the first."
    : `Slide ${state.slide}.`;
}

/**
 * Where this element would land, in one sentence, for the preview card.
 *
 * `docs/DESIGN.md` section 5 decides the landing per collection slide, so the
 * deck decides it and this only says what the catalogue already recorded.
 *
 * The insert target is the user's, from the gear — except that **a part ignores
 * it and always lands on the slide the user is on** (section 5, last bullet).
 * Saying "as a new slide" over a stamp would be the pane promising something
 * the engine does not do.
 */
export function landingLine(element: Element, settings: Settings): string {
  if (element.kind === "part") {
    switch (element.landing) {
      case "top-right":
        return "Lands top-right of this slide, clear of the edge.";
      case "cursor":
        return "Lands on the shape you have selected, or in the middle of this slide.";
      default:
        return "Lands on this slide, where it sits in the library.";
    }
  }
  return settings.target === "new"
    ? "Lands as a new slide after this one."
    : "Lands below your slide's own title, scaled to fit the room under it.";
}
