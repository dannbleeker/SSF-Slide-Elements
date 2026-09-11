/**
 * What the pane decides, kept away from what it draws.
 *
 * `render.ts` builds DOM from the answers here and `main.ts` wires the host in;
 * neither decides anything of its own, and `test/architecture.test.ts` holds
 * that seam. Everything in this file is an ordinary function over plain values,
 * which is what lets the suite check the pane's behaviour with no PowerPoint
 * and no browser — and what lets `scripts/pane-shots.mjs` audit a state by
 * handing `render` a value rather than by driving a host.
 *
 * The picker is three steps, and they are STATES rather than a wizard: the
 * library is loading, the library is here, or it did not come. `docs/DESIGN.md`
 * section 10 is the rule they implement — the pane never shows an empty or
 * broken screen, and every message says what happened and what to do.
 */
import type { Element, SlideSize } from "../core/catalogue/types.js";

export type StepId = "loading" | "browse" | "problem";

export const STEPS: readonly StepId[] = ["loading", "browse", "problem"];

/** The heading each step draws. The manual must quote every one of these. */
export const STEP_TITLE: Record<StepId, string> = {
  loading: "Loading the library",
  browse: "Slide elements",
  problem: "The library did not load",
};

/** What the gear holds. `docs/DESIGN.md` section 7. */
export interface Settings {
  /** Onto the slide the user is on, or as a new slide after it. */
  target: "onto" | "new";
  /** Multi-shape elements land as one group, or loose. */
  group: boolean;
}

export const DEFAULT_SETTINGS: Settings = { target: "onto", group: true };

/** The catalogue as the pane holds it: one size's worth. */
export interface Library {
  size: SlideSize;
  width: number;
  height: number;
  version: string;
  categories: { key: string; name: string }[];
  elements: Element[];
  /**
   * Set when the deck is neither 16:9 nor 4:3 and the nearest library was
   * borrowed. `docs/DESIGN.md` section 3 puts the sentence under the header.
   */
  borrowed?: string;
}

/** The last insert, as the footer reports it. */
export interface Outcome {
  ok: boolean;
  detail: string;
  byHand: boolean;
  /** The element's name, for the live region. */
  name: string;
}

export interface PaneState {
  /** What the host or the pane has to say, announced as well as shown. */
  notice?: string;
  library?: Library;
  /** Why the library did not load, when it did not. */
  problem?: string;
  /** What the user typed in the search box. */
  query: string;
  /** Tags the user has picked. */
  tags: string[];
  /** Category keys the user has opened. */
  open: string[];
  /** The tile the keyboard is on, if any. */
  chosen?: string;
  settings: Settings;
  /** True while an insert runs: the pane locks and the chosen tile says so. */
  busy?: boolean;
  outcome?: Outcome;
  /** The slide the user is on, counting from one, when the host would say. */
  slide?: number;
  /** How many inserts can still be taken back. */
  undo: number;
  /** Element ids the user has starred, most recent first. */
  favourites: string[];
  /** Element ids inserted recently, most recent first. */
  recent: string[];
  /** True when the gear is open. */
  gear?: boolean;
  /**
   * The element whose preview card is open, if any.
   *
   * In the state rather than managed beside it, so the card cannot survive a
   * re-render that removed the tile it belongs to. `docs/DESIGN.md` section 4
   * opens it after a third of a second of hover or focus; the delay lives in
   * the wiring, and by the time it reaches here the card is simply open.
   */
  previewing?: string;
}

export const EMPTY: PaneState = {
  query: "",
  tags: [],
  open: [],
  settings: DEFAULT_SETTINGS,
  undo: 0,
  favourites: [],
  recent: [],
};

/** Which step a state is in. The pane has no other way to choose one. */
export function stepFor(state: PaneState): StepId {
  if (state.problem !== undefined) return "problem";
  return state.library ? "browse" : "loading";
}

/**
 * Whether a query matches an element.
 *
 * `docs/DESIGN.md` section 8: the English name, the Danish key, the category
 * and the tags. The key is searched even though the pane never shows it,
 * because the owner's decks are Danish and somebody who knows the library by
 * its Danish names should be able to find things.
 *
 * Every WORD has to match, not the whole string as a substring: "white box"
 * should find "White boxes, 2x1 vertical", where a substring match on the
 * phrase finds nothing at all.
 */
export function matches(element: Element, query: string): boolean {
  const q = query.trim().toLowerCase();
  if (q === "") return true;
  const hay = [element.name, element.key, element.category.name, element.category.key, ...element.tags]
    .join(" ")
    .toLowerCase();
  return q.split(/\s+/).every((word) => hay.includes(word));
}

/** Every tag in the library, most used first, for the tag line. */
export function tagsOf(library: Library): string[] {
  const counts = new Map<string, number>();
  for (const el of library.elements) for (const tag of el.tags) counts.set(tag, (counts.get(tag) ?? 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).map(([tag]) => tag);
}

/** One category with the elements a search left in it. */
export interface Group {
  key: string;
  name: string;
  elements: Element[];
}

/**
 * The elements to show, grouped by category, in the deck's own order.
 *
 * A sized element is ONE tile with a stepper (`docs/DESIGN.md` section 2), so a
 * run of "process flow with 1 box" through "with 6 boxes" collapses to its
 * first surviving member and the rest are reachable through the stepper. The
 * 16:9 deck has twelve such runs over 45 elements, so 117 elements show as 84
 * tiles: 117 − 45 + 12. Counted by running this function over the committed
 * catalogue, and confirmed against the count the pane itself draws.
 */
export function groups(library: Library, state: PaneState): Group[] {
  const wanted = library.elements.filter(
    (el) => matches(el, state.query) && state.tags.every((tag) => el.tags.includes(tag)),
  );
  const seen = new Set<string>();
  const tiles = wanted.filter((el) => {
    if (!el.run) return true;
    if (seen.has(el.run.key)) return false;
    seen.add(el.run.key);
    return true;
  });
  const out: Group[] = [];
  for (const category of library.categories) {
    const elements = tiles.filter((el) => el.category.key === category.key);
    if (elements.length === 0) continue;
    out.push({ key: category.key, name: category.name, elements });
  }
  return out;
}

/** Every member of a sized element's run, in count order, for the stepper. */
export function runOf(library: Library, element: Element): Element[] {
  if (!element.run) return [element];
  const key = element.run.key;
  return library.elements.filter((el) => el.run?.key === key).sort((a, b) => (a.run?.count ?? 0) - (b.run?.count ?? 0));
}

/** How many tiles a search left, for the count beside the categories. */
export function tileCount(found: Group[]): number {
  return found.reduce((n, g) => n + g.elements.length, 0);
}

/**
 * Whether a category is open.
 *
 * Categories start collapsed (`docs/DESIGN.md` section 4), and a search or a
 * tag opens what it finds — otherwise a query that matched three things in two
 * categories would show two closed headers and look like nothing was found.
 */
export function isOpen(state: PaneState, key: string): boolean {
  if (state.query.trim() !== "" || state.tags.length > 0) return true;
  return state.open.includes(key);
}

/** The element a tile id names, if the library still has it. */
export function elementOf(library: Library | undefined, id: string | undefined): Element | undefined {
  if (!library || id === undefined) return undefined;
  return library.elements.find((el) => el.id === id);
}

/**
 * Why the one primary control cannot be pressed.
 *
 * Stated rather than implied by a greyed-out control: a disabled button with no
 * sentence beside it is a pane that looks broken.
 */
export function blockedReason(state: PaneState, step: StepId): string {
  if (step === "loading") return "Fetching the element library from the add-in's own site.";
  if (step === "problem") {
    return (
      state.problem ??
      "The library did not load. That is usually the network rather than a fault in the add-in, so try again."
    );
  }
  if (state.busy === true) return "One insert at a time. This one is still going.";
  if (state.chosen === undefined) return "Choose an element to insert it, or use the arrow keys and press Enter.";
  return "";
}

/** The one primary control per screen: what it says, and whether it can be pressed. */
export function primary(state: PaneState, step: StepId): { label: string; disabled: boolean } {
  if (step === "problem") return { label: "Try again", disabled: false };
  return { label: "Insert an element", disabled: step !== "browse" || state.busy === true || !state.chosen };
}

/**
 * The line under the header, when there is one.
 *
 * Only for a deck that is neither 16:9 nor 4:3: "4:3 library, scaled to A4
 * slides" (`docs/DESIGN.md` sections 3 and 4). A pane that said nothing there
 * would leave the user wondering why an element does not quite fit.
 */
export function borrowedLine(state: PaneState): string | undefined {
  return state.library?.borrowed;
}

/** What the footer says about the last insert, and what it offers. */
export interface Footer {
  detail: string;
  /** Whether the deck needs the user to put something right by hand. */
  byHand: boolean;
  undo: number;
  /** True when the last insert can be repeated on the current slide. */
  again: boolean;
}

export function footerOf(state: PaneState): Footer {
  return {
    detail: state.outcome?.detail ?? "",
    byHand: state.outcome?.byHand ?? false,
    undo: state.undo,
    again: state.recent.length > 0 && state.busy !== true,
  };
}

/** The gear's own line, so the settings are visible without opening it. */
export function settingsLine(settings: Settings): string {
  const where = settings.target === "onto" ? "onto this slide" : "as a new slide";
  const how = settings.group ? "as one group" : "loose";
  return `Inserting ${where}, ${how}.`;
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

/** How many recent elements the pane remembers. Section 4: the last six inserts. */
export const RECENT_DEPTH = 6;

/** `id` at the front, with any earlier mention removed, capped at `depth`. */
export function remember(list: string[], id: string, depth: number): string[] {
  return [id, ...list.filter((other) => other !== id)].slice(0, depth);
}

/** A starred element, or one no longer starred. */
export function toggle(list: string[], id: string): string[] {
  return list.includes(id) ? list.filter((other) => other !== id) : [id, ...list];
}
