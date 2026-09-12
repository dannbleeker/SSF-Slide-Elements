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
import type { Box, Element, SlideSize } from "../core/catalogue/types.js";

export type StepId = "loading" | "browse" | "problem";

/** The heading each step draws. The manual must quote every one of these. */
export const STEP_TITLE: Record<StepId, string> = {
  loading: "Loading the library",
  browse: "Slide elements",
  problem: "The library did not load",
};

/**
 * Every step there is, DERIVED rather than listed.
 *
 * `test/docs.test.ts` sweeps this to hold the manual to the pane, so what it
 * contains decides how wide that gate is. It used to be a hand-written copy of
 * the `StepId` union, and a copy is a thing that goes stale: adding a fourth
 * step and a fourth heading left the manual gate green, because the sweep was
 * still walking the three someone had typed here. `STEP_TITLE` is a
 * `Record<StepId, …>`, so the compiler will not let it miss a step — reading
 * the list off it makes the sweep as wide as the type.
 */
export const STEPS: readonly StepId[] = Object.keys(STEP_TITLE) as StepId[];

/** What the gear holds. `docs/DESIGN.md` section 7. */
export interface Settings {
  /** Onto the slide the user is on, or as a new slide after it. */
  target: "onto" | "new";
  /** Multi-shape elements land as one group, or loose. */
  group: boolean;
  /**
   * `deck`: theme-referenced colours follow the deck the element lands in, so
   * an element takes the customer's blue. `library`: every one of them is
   * pinned to the value the library's own theme gave it.
   *
   * `deck` is the default because it is what most people want most of the time
   * — an element that looks like the deck it is in — and because it is what
   * happens by itself.
   */
  colours: "deck" | "library";
}

export const DEFAULT_SETTINGS: Settings = { target: "onto", group: true, colours: "deck" };

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
   * The category chip the user has picked while searching, if any.
   *
   * Section 8's chips narrow a search to one category. Not remembered per
   * machine and not carried across a cleared search: it is a way of reading one
   * set of results, not a preference.
   */
  category?: string;
  /**
   * True once the first-run coach marks have been dismissed.
   *
   * Remembered per machine, like the favourites and the settings. Section 4:
   * "dismissed once" — not once per session, and not once per deck.
   */
  coached?: boolean;
  /**
   * The element whose preview card is open, if any.
   *
   * In the state rather than managed beside it, so the card cannot survive a
   * re-render that removed the tile it belongs to. `docs/DESIGN.md` section 4
   * opens it after a third of a second of hover or focus; the delay lives in
   * the wiring, and by the time it reaches here the card is simply open.
   */
  previewing?: string;
  /**
   * What this deck already carries, read from the tags the add-in wrote
   * (`docs/DESIGN.md` section 4, "Used in this deck").
   *
   * `undefined` means NOT ASKED — not "nothing", which is the empty array. The
   * two are different states on the screen and the difference is the whole
   * reason the field is optional: reading it means reading the user's entire
   * deck, and section 13's sixth open question is how long that takes on a 50 MB
   * one. So the pane asks when the user asks, and says which it is.
   */
  used?: DeckUsage[];
  /** True while that read is running. */
  reading?: boolean;
  /**
   * Whether a slide number in that list can be a link.
   *
   * Set once at boot from the host: `setSelectedSlides` is PowerPointApi 1.5.
   * Below it the numbers stay text, because a control that might do nothing is
   * worse than a sentence that says where the element is (`docs/DESIGN.md`
   * section 4).
   */
  canJump?: boolean;
  /**
   * The tile whose right-click menu is open, as `tileKey` spells it.
   *
   * Not the element id alone: one element is drawn up to three times — in
   * Favourites, in Recent, and in its own category — and an id matches all of
   * them, so a single right-click opened a menu on every copy. The key names
   * the TILE. Still no coordinates: the menu is anchored to its own tile, so
   * where the pointer was stays the browser's business, and the shot audit can
   * still draw the state.
   */
  menuFor?: string;
  /**
   * What the slide the user is on already holds, for the preview card's grey
   * boxes (`docs/DESIGN.md` sections 1 and 4).
   *
   * Stamped with the slide it was read FROM, because it goes stale the moment
   * the user clicks another slide and the pane does not re-read the deck on
   * every slide change — that is a whole-presentation read, and section 13's
   * sixth open question is how long one takes. The card draws the boxes when
   * this names the slide the user is on and draws none when it does not, which
   * is the honest half of a snapshot.
   */
  onSlide?: { slide: number; boxes: Box[] };
  /**
   * The user's own slide size in EMU, when the deck has been read.
   *
   * NOT `library.width`/`library.height`, which are the LIBRARY deck's — and
   * the two are different numbers exactly when a deck borrows the nearest
   * library (A4, 16:10, anything custom). Everything measured against the
   * user's slide — the boxes the card draws in grey, and where the splice says
   * an element landed — is in these units, so mixing the two draws the right
   * rectangle in the wrong place on precisely the decks nobody tests on.
   */
  deck?: { width: number; height: number };
  /**
   * The removal the user is being asked to confirm, and how far it has got.
   *
   * `docs/DESIGN.md` section 6's deck-wide stamps. Confirmed rather than done on
   * the click, because it is the only thing this add-in does that takes
   * something OUT of somebody's deck and the pane cannot put it back: Undo is
   * one insert deep by design (section 6), and an undo of this would mean
   * holding a copy of the deck per slide touched.
   *
   * `done` counts slides finished, so a run that stops halfway can say where it
   * stopped rather than leaving the user to count.
   */
  removing?: { id: string; slides: number[]; done: number; where: string };
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
    (el) =>
      matches(el, state.query) &&
      state.tags.every((tag) => el.tags.includes(tag)) &&
      (state.category === undefined || el.category.key === state.category),
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

/** A category a search found something in, and how many tiles it left there. */
export interface CategoryHit {
  key: string;
  name: string;
  count: number;
}

/**
 * Which categories a search found something in, with counts, for the chips
 * `docs/DESIGN.md` section 8 puts above the results while searching.
 *
 * Counted WITHOUT the category the user has already picked applied, which is
 * the whole point: the chips have to keep showing the other categories and
 * their counts, or picking one would hide the way back and the way across. The
 * query and the tags DO apply, because those are what the counts are counts of.
 *
 * Counts are tiles, not elements, so they agree with the number beside the
 * search and with what a person can see: a run of six sizes is one tile.
 */
export function categoryHits(library: Library, state: PaneState): CategoryHit[] {
  const free: PaneState = { ...state, category: undefined };
  return groups(library, free).map((g) => ({ key: g.key, name: g.name, count: g.elements.length }));
}

/**
 * Whether this member of a sized run matches the search.
 *
 * Section 8: a sized tile greys out the counts that do not match. The tile
 * itself is the first member that survived the search, and its stepper still
 * offers every size — so without this, a search for "3 boxes" shows a stepper
 * of six numbers with nothing to say which one was searched for.
 */
export function stepMatches(member: Element, state: PaneState): boolean {
  return matches(member, state.query) && state.tags.every((tag) => member.tags.includes(tag));
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
/**
 * Which TILE something is open on.
 *
 * An element appears in Favourites, in Recent and in its own category, so the
 * pane draws up to three tiles for one id. Anything anchored to a tile — the
 * right-click menu, the question before a removal — has to name the tile rather
 * than the element, or one right-click opens three menus.
 */
export function tileKey(where: string, id: string): string {
  return `${where}:${id}`;
}

/**
 * The insert target the gear is NOT set to.
 *
 * `docs/DESIGN.md` section 6: right-click a tile and the pane offers the other
 * target for that one insert, without touching the setting. So there is exactly
 * one other thing to offer, and this is it.
 */
export function otherTarget(settings: Settings): Settings["target"] {
  return settings.target === "onto" ? "new" : "onto";
}

/**
 * Whether right-clicking this element offers anything.
 *
 * Only a whole-slide element. A PART — a stamp, a marker, a flowchart shape —
 * ignores the insert target entirely and always lands on the slide the user is
 * on (section 5), so a menu offering it "as a new slide" would be the pane
 * promising something the engine does not do. The preview card's landing line
 * already says as much for the same reason.
 */
export function offersOtherTarget(element: Element): boolean {
  return element.kind === "slide";
}

/** What that one menu item says, which is an action rather than a setting. */
export function otherTargetLabel(settings: Settings): string {
  return otherTarget(settings) === "new" ? "Insert as a new slide" : "Insert onto this slide";
}

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
 * Whether a tile offers "Remove from N slides", and for which slides.
 *
 * Section 4 puts it on a PART already in the deck — a stamp or a marker, the
 * things that go on many slides. A whole-slide element is one slide's worth of
 * content and "remove" for one of those is the slide's own delete key.
 *
 * Answers nothing until the deck has been read: the pane does not know what is
 * in the deck before then, and a button offering to remove something from
 * nowhere is worse than no button.
 */
export function removableFrom(element: Element, state: PaneState): number[] {
  if (element.kind !== "part" || state.used === undefined) return [];
  return state.used.find((u) => u.element === element.id)?.slides ?? [];
}

/** What that button says, counting the slides rather than the shapes. */
export function removeLabel(slides: number[]): string {
  return slides.length === 1 ? "Remove from 1 slide" : `Remove from ${slides.length} slides`;
}

/**
 * The question the pane asks before it takes anything out of the deck.
 *
 * Names the slides, because "3 slides" is not something a user can check and
 * "slides 2, 5 and 9" is. Says the pane cannot undo it, because it cannot:
 * Undo is one insert deep, and this is not an insert.
 */
export function removeQuestion(element: Element, slides: number[]): string {
  return `Take ${element.name} off ${slideList(slides)}? The pane cannot undo this.`;
}

/**
 * How far a removal got, for the footer.
 *
 * Three answers, not two. "All of them" and "some of them" are the obvious
 * pair; the third is **none to remove**, and it exists because the removal
 * re-reads which slides carry the element from the deck it is about to change
 * rather than trusting the list the question was asked about. That re-read is
 * what makes a removal honest about a deck that has moved on — and it can come
 * back empty, when the user has taken the shapes off by hand since the pane
 * last looked.
 *
 * That case used to answer `Removed from 0 slides.` and call it a success,
 * which is a sentence that reads as a glitch: nothing was removed, and the pane
 * said it had removed things. It is still not a FAILURE — there was nothing to
 * do and nothing went wrong — so it says so, and offers no by-hand advice for
 * work that is already done.
 */
export function removalOutcome(element: string, done: number, wanted: number): Outcome {
  if (wanted === 0) {
    return { ok: true, byHand: false, name: element, detail: "It is not on any slide any more, so nothing changed." };
  }
  const ok = done === wanted;
  return {
    ok,
    byHand: !ok,
    name: element,
    detail: ok
      ? `Removed from ${done === 1 ? "1 slide" : `${done} slides`}.`
      : `Removed from ${done} of ${wanted} slides. The rest are as they were — try again, or take them off by hand.`,
  };
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

export function settingsLine(settings: Settings): string {
  const where = settings.target === "onto" ? "onto this slide" : "as a new slide";
  const how = settings.group ? "as one group" : "loose";
  // The colours only get a clause when they are NOT the default, so the line
  // stays short for the setting almost everybody is on and says the surprising
  // thing out loud for the one they are not.
  const colours = settings.colours === "library" ? ", in the library's own colours" : "";
  return `Inserting ${where}, ${how}${colours}.`;
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

/**
 * How far apart two strings are, counting single-character edits.
 *
 * The ordinary Levenshtein distance, one row at a time so a 117-element library
 * costs a few thousand numbers rather than a matrix per name.
 */
function distance(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  let row = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const next = [i];
    for (let j = 1; j <= b.length; j++) {
      const substitute = (row[j - 1] as number) + (a[i - 1] === b[j - 1] ? 0 : 1);
      next[j] = Math.min(substitute, (row[j] as number) + 1, (next[j - 1] as number) + 1);
    }
    row = next;
  }
  return row[b.length] as number;
}

/**
 * What the user might have meant, for a query that found nothing.
 *
 * `docs/DESIGN.md` section 8. A search that returns an empty list is a dead end;
 * this turns it into a route, and it only ever offers names the library really
 * has, so picking one cannot fail.
 *
 * Measured against each WORD of a name as well as the whole of it, because a
 * query is usually one word and "triangel" should reach "Triangle, simple, with
 * text at the corners" — which as a whole string is 30 edits away from it.
 *
 * The threshold grows with the query: one edit for a short word, more for a
 * long one. Without that, a three-letter typo would reach half the library.
 */
export function didYouMean(library: Library, query: string, limit = 3): string[] {
  const q = query.trim().toLowerCase();
  if (q.length < 3) return [];
  const allowed = Math.max(1, Math.floor(q.length / 3));

  const scored: { name: string; cost: number }[] = [];
  for (const element of library.elements) {
    const name = element.name;
    const words = name
      .toLowerCase()
      .split(/[^\p{L}\p{N}]+/u)
      .filter(Boolean);
    let best = distance(q, name.toLowerCase());
    for (const word of words) best = Math.min(best, distance(q, word));
    if (best <= allowed) scored.push({ name, cost: best });
  }

  scored.sort((a, b) => a.cost - b.cost || a.name.localeCompare(b.name, "en"));
  const seen = new Set<string>();
  const out: string[] = [];
  for (const { name } of scored) {
    if (seen.has(name)) continue;
    seen.add(name);
    out.push(name);
    if (out.length === limit) break;
  }
  return out;
}

/**
 * The three things a first-time user is told, in the order they will need them.
 *
 * `docs/DESIGN.md` section 4: hover to preview, click to insert, Undo and the
 * gear, dismissed once.
 *
 * They are three LINES of one panel rather than three callouts pointing at the
 * controls, and that is a deliberate departure recorded in section 4: this pane
 * is 320 px at its narrowest, where three floating callouts would cover the
 * tiles, the footer and the gear they were pointing at.
 */
export const COACH: readonly string[] = [
  "Rest on a tile to see the element up close.",
  "Click it to put it on your slide.",
  "Undo takes it back, and the gear changes where things land.",
];

/** Whether the first-run coach marks should be drawn. */
export function coaching(state: PaneState): boolean {
  return state.library !== undefined && state.coached !== true;
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
