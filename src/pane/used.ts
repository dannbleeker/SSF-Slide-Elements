/**
 * "Used in this deck": which library elements the open deck already holds.
 *
 * Split out of `steps.ts` on 2026-09-12. `docs/DESIGN.md` section 4's feature,
 * read from the tags an insert writes into the package — measured surviving
 * `insertSlidesFromBase64` on 2026-09-11, which is what makes it buildable at
 * all. Everything about the list lives here: the rows, the slide numbers as
 * text or as links, the heading, and the two updates that keep the list in step
 * with an insert and its undo without re-reading the deck.
 *
 * `DeckUsage` is the shape `PaneState.used` holds, so `steps.ts` imports the
 * TYPE back from here. That direction is type-only on purpose — it is erased at
 * build time, so the two files have no runtime cycle. The one real edge is the
 * other way: `removeQuestion` in `steps.ts` calls `slideList`.
 */
import type { Library, PaneState } from "./steps.js";

/** One element the deck already carries: the engine's answer, as the pane holds it. */
export interface DeckUsage {
  /** The catalogue id out of the shape's tag. */
  element: string;
  /** The slides it is on, 1-based and in order. */
  slides: number[];
}

/** One row of "Used in this deck", ready to draw. */
export interface UsedRow {
  id: string;
  /** The element's name, or a sentence saying why there is none. */
  name: string;
  /** False when the catalogue has no element with this id. */
  known: boolean;
  slides: number[];
  /** "slide 2" or "slides 2, 5 and 9", which is what the row says after the name. */
  where: string;
}

/** One piece of "slides 2, 5 and 9": a word to print, or a slide number that can be a link. */
export type SlidePart = { text: string } | { slide: number };

/**
 * "slide 2", "slides 2 and 5", "slides 2, 5 and 9" — a list a person would read
 * aloud, in pieces, so the renderer can make each number a control and keep
 * the words as words. Sorted and deduplicated, because the same element twice
 * on one slide is one place to look.
 */
export function slideParts(slides: number[]): SlidePart[] {
  const sorted = [...new Set(slides)].sort((a, b) => a - b);
  if (sorted.length === 0) return [];
  const out: SlidePart[] = [{ text: sorted.length === 1 ? "slide " : "slides " }];
  sorted.forEach((slide, i) => {
    if (i > 0) out.push({ text: i === sorted.length - 1 ? " and " : ", " });
    out.push({ slide });
  });
  return out;
}

/** The same list as one string, for a sentence. */
export function slideList(slides: number[]): string {
  return slideParts(slides)
    .map((part) => ("text" in part ? part.text : String(part.slide)))
    .join("");
}

/**
 * What the deck carries, named against the library the pane has open.
 *
 * An id the catalogue no longer has is KEPT and says so. Eleven ids changed on
 * 2026-09-11 when the part keys were translated (`docs/DESIGN.md` section 2), so
 * a deck stamped before that names elements this library cannot — and a row
 * silently dropped would make the deck look emptier than it is, which is the
 * one thing this list exists not to do.
 */
export function usedRows(library: Library | undefined, used: DeckUsage[]): UsedRow[] {
  return used.map((use) => {
    const element = library?.elements.find((e) => e.id === use.element);
    return {
      id: use.element,
      name: element?.name ?? "An element from an older version of the library",
      known: element !== undefined,
      slides: use.slides,
      where: slideList(use.slides),
    };
  });
}

/** The heading over that list, which has to say which of three states the pane is in. */
export function usedHeading(state: PaneState): string {
  if (state.reading === true) return "Reading this deck…";
  if (state.used === undefined) return "Used in this deck";
  if (state.used.length === 0) return "Nothing from the library is in this deck yet";
  return `Used in this deck (${state.used.length})`;
}

/**
 * The deck's usage with one more insert in it.
 *
 * The pane does not re-read the deck after every insert: the read is the
 * expensive thing this feature costs, and the pane already knows exactly what
 * it just put where. So the list is updated rather than refetched, and the
 * next explicit read is what reconciles it with the file.
 *
 * Answers undefined when nothing has been read yet, because an insert is not a
 * reason to start claiming the deck has been looked at.
 */
export function withInsert(used: DeckUsage[] | undefined, element: string, slide: number): DeckUsage[] | undefined {
  if (used === undefined) return undefined;
  const found = used.find((u) => u.element === element);
  if (!found) return [...used, { element, slides: [slide] }];
  return used.map((u) =>
    u.element === element ? { element, slides: [...new Set([...u.slides, slide])].sort((a, b) => a - b) } : u,
  );
}

/**
 * The deck's usage with an insert taken back out.
 *
 * Undo puts the user's own slide back, so whatever the insert added to THAT
 * slide is gone with it. An element still on other slides keeps those.
 */
export function withoutInsert(used: DeckUsage[] | undefined, element: string, slide: number): DeckUsage[] | undefined {
  if (used === undefined) return undefined;
  return used
    .map((u) => (u.element === element ? { element, slides: u.slides.filter((n) => n !== slide) } : u))
    .filter((u) => u.slides.length > 0);
}
