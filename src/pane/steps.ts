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
import { slideList, type DeckUsage } from "./used.js";

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
  /**
   * The element the footer offers to move to a new slide, by id.
   *
   * `docs/DESIGN.md` section 6: a whole-slide element that landed ONTO a slide
   * which already had content gets the offer beside Undo. Set only when all
   * four hold — the insert worked, the target was "onto", the element is a
   * whole slide rather than a part, and the splice counted something already
   * on the slide — because the move is an undo followed by a second insert and
   * it can offer nothing an undo cannot deliver.
   *
   * The id rather than a flag, so the pane re-inserts the element the user
   * actually placed even if the tile they are hovering has moved on. Mirrors
   * `undoable` in `main.ts` the way `undo` does: the entry itself is not in the
   * state, and the render needs to know the offer stands.
   */
  moveable?: string;
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
  /**
   * True when the last insert can be moved onto a new slide instead.
   *
   * Gated on `undo` as well as on `moveable`, and that is not belt and braces:
   * the move is an undo followed by a second insert, so an offer standing
   * after the history has gone would be a button that cannot do what it says.
   *
   * NOT gated on `busy`, unlike `again` beside it. The two differ because they
   * are about different things: Again offers a NEW insert, which there is no
   * point drawing while one is running, and the move is about the insert that
   * just happened, which is still the last one whatever the pane is doing. So
   * it stays on screen and `render.ts` disables it, the way Undo does — a
   * control that vanishes and comes back under the cursor is worse than one
   * that greys out.
   */
  move: boolean;
}

export function footerOf(state: PaneState): Footer {
  return {
    detail: state.outcome?.detail ?? "",
    byHand: state.outcome?.byHand ?? false,
    undo: state.undo,
    again: state.recent.length > 0 && state.busy !== true,
    move: state.moveable !== undefined && state.undo > 0,
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
 * Whether this element goes AROUND the selected shape rather than on top of it.
 *
 * `docs/DESIGN.md` section 5: with a shape selected, a marker wraps it — sized
 * to the shape with a little air — while a stamp, a flowchart shape or an icon
 * lands at the cursor without resizing. The splice takes it as `wraps`
 * (`src/core/splice/landing.ts`), and this is the only place that decides which
 * elements set it.
 *
 * The library says which are markers by the CATEGORY, because nothing else in
 * an element distinguishes them: `landing` is `cursor` for markers, flowchart
 * shapes and icons alike, and `kind` is `part` for all three.
 *
 * **Matched at the START of the name, not anywhere in it**, and the tightening
 * is the point of this function existing. Until 2026-09-12 the rule was
 * `key.toLowerCase().includes("mark")` inline in `main.ts` — untested, because
 * that file is the one the coverage floor exempts. The committed library's
 * Danish category key is `Markeringer`, so the substring test is right on the
 * library as it stands today and wrong on plausible additions to it:
 * `Danmarkskort` (a map of Denmark) contains "mark", and every element in such
 * a category would have started resizing itself around whatever the user had
 * selected. Both the Danish key and the English name are tried, so a deck the
 * owner renames keeps working.
 *
 * The tightening changes nothing on the library as it stands: measured over the
 * committed catalogue on 2026-09-12, the two rules agree on all 234 elements,
 * of which 6 wrap. So this is a latent bug closed, not a behaviour changed.
 */
export function wrapsSelection(element: Element): boolean {
  return isMarkers(element.category.key) || isMarkers(element.category.name);
}

/** "Markeringer", "Markers", "Markers and labels" — but not "Danmarkskort" or "Markedsandel". */
function isMarkers(name: string): boolean {
  return name.trim().toLowerCase().startsWith("marker");
}

/**
 * Whether the insert that just happened can be offered "Move to a new slide".
 *
 * `docs/DESIGN.md` section 6. Four conditions, and each one removes an offer
 * that could not be honoured:
 *
 * - **The insert worked.** There is nothing to move otherwise, and the move is
 *   an undo followed by a second insert.
 * - **It went ONTO a slide.** An element already on a slide of its own is where
 *   the offer would put it.
 * - **It is a whole-slide element.** A part ignores the target switch entirely
 *   (`offersOtherTarget` above), so "as a new slide" is not a thing the engine
 *   would do differently.
 * - **The slide already held something.** `held` is what the splice counted in
 *   the bytes it was handed, so the offer costs no second read and no host
 *   call. On an empty slide the element covers nothing and the offer is noise.
 *
 * Answers the element's ID rather than a flag, which is what `moveable` stores:
 * the second insert must place the element the user actually placed, not
 * whatever tile the pointer has wandered onto since.
 */
export function moveableAfter(
  attempt: { ok: boolean; target: Settings["target"]; held: number },
  element: Element,
): string | undefined {
  const offer = attempt.ok && attempt.target === "onto" && element.kind === "slide" && attempt.held > 0;
  return offer ? element.id : undefined;
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

export function settingsLine(settings: Settings): string {
  const where = settings.target === "onto" ? "onto this slide" : "as a new slide";
  const how = settings.group ? "as one group" : "loose";
  // The colours only get a clause when they are NOT the default, so the line
  // stays short for the setting almost everybody is on and says the surprising
  // thing out loud for the one they are not.
  const colours = settings.colours === "library" ? ", in the library's own colours" : "";
  return `Inserting ${where}, ${how}${colours}.`;
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

/**
 * What Escape shuts, given everything the pane currently has open.
 *
 * `docs/DESIGN.md` section 9: back out of what was opened last, and never clear
 * a search on the way past something else. A question the user escapes is a
 * question answered "no", so the confirm goes first.
 *
 * A LADDER rather than five independent checks, and the order is the whole
 * rule: it lived inside the pane's key handler, where nothing could reach it,
 * and getting it wrong loses somebody's search while they were only trying to
 * shut a preview card. The names are what to close; `main.ts` does the closing,
 * because two of them cancel a timer as well as clearing a field.
 */
export type EscapeTarget = "removing" | "menu" | "preview" | "gear" | "search" | undefined;

export function escapeCloses(state: PaneState): EscapeTarget {
  if (state.removing !== undefined) return "removing";
  if (state.menuFor !== undefined) return "menu";
  if (state.previewing !== undefined) return "preview";
  if (state.gear === true) return "gear";
  if (state.query !== "" || state.tags.length > 0) return "search";
  return undefined;
}

/**
 * Which tile an arrow key moves the focus to, counting from zero.
 *
 * Clamped at both ends rather than wrapping: a grid that jumps from the last
 * tile back to the first reads as a glitch, and the two ends are exactly where
 * an off-by-one hides. Answers nothing for a key that is not an arrow and for a
 * list with no tiles in it.
 *
 * `at` is where the focus is now, or `-1` when it is on none of them — the
 * first arrow press then lands on the FIRST tile whichever direction it was,
 * which is what a user pressing Down from the search box expects. That falls
 * out of the clamp (`-1 + 1` and `-1 - 1` both floor to 0) rather than needing
 * a case of its own: `main.ts` carried one, and it could not be made to go red.
 */
export function arrowTo(key: string, at: number, count: number): number | undefined {
  const step = key === "ArrowRight" || key === "ArrowDown" ? 1 : key === "ArrowLeft" || key === "ArrowUp" ? -1 : 0;
  if (step === 0 || count <= 0) return undefined;
  return Math.min(count - 1, Math.max(0, at + step));
}
