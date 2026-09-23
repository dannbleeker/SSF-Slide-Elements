/**
 * The pane's entry point, and the only file here that touches Office.js.
 *
 * Everything it shows comes from `render.ts` and everything it decides comes
 * from `steps.ts` and `src/host`, all checked by the suite without a PowerPoint
 * anywhere. `test/architecture.test.ts` holds that seam: a decision that
 * migrates into this file becomes untestable the moment it arrives.
 *
 * Excluded from coverage for the same reason — pooling it with the engine would
 * produce one number that hides both.
 */
import { occupiedBoxes } from "../core/catalogue/boxes.js";
// `Element` is aliased because the DOM has one too, and this file is full of
// both. `LibraryElement` is what the catalogue means by it.
import type { Box, Element as LibraryElement, Markup } from "../core/catalogue/types.js";
import { slideSize } from "../core/pptx/layout.js";
import { Pkg } from "../core/pptx/pkg.js";
import { usedInDeck } from "../core/pptx/tags.js";
import { removeElement, slidesHolding } from "../core/splice/remove.js";
import { onlySlide, splice } from "../core/splice/splice.js";
import { coalescing } from "../host/coalesce.js";
import {
  INSERTING,
  announcement,
  landedOn,
  mayRemove,
  outcomeOf,
  stampTargets,
  stillThere,
  undoPlan,
} from "../host/insert.js";
import { readable } from "../host/errors.js";
import { jumpOutcome } from "../host/jump.js";
import { GLOBAL_KEY, deckKey } from "../host/memory.js";
import { paneTheme } from "../host/theme.js";
import { catalogueUrl, reportUrl, siteFrom } from "../host/links.js";
import { BUDGET } from "../host/timeout.js";
import {
  currentSlide,
  deckUrl,
  hostStamp,
  hostSupports,
  selectSlide,
  insertPackage,
  onSlideChange,
  openExternal,
  ready as hostReady,
  readDeck,
  removeSlideAt,
  selectedShape,
  selectedSlides,
  countReaching,
  slideCount,
  slideIdAt,
  themeBackground,
} from "../office/powerpoint.js";
import { Store, carriedTypes, libraryFor, loadIndex, themeColours, type Index } from "./catalogue.js";
import { firstVisit, restored, shouldRestoreScroll, storedScroll, writes } from "./storage.js";
import { render } from "./render.js";
import { fractionOf, withLanded } from "./card.js";
import { elementOf, openAtFirst } from "./search.js";
import {
  EMPTY,
  RECENT_DEPTH,
  arrowTo,
  escapeCloses,
  moveableAfter,
  offersOtherTarget,
  otherTarget,
  remember,
  removableFrom,
  removalOutcome,
  removeQuestion,
  stampOutcome,
  stepFor,
  tileKey,
  toggle,
  wrapsSelection,
  type Library,
  type PaneState,
} from "./steps.js";
import { holds, renumbered, withInsert, withoutInsert } from "./used.js";

let state: PaneState = { ...EMPTY };
let index: Index | undefined;
let store: Store | undefined;

/**
 * What Undo needs to put the deck back, for the LAST insert only.
 *
 * One deep, and `docs/DESIGN.md` section 6 records why that is not the ten it
 * originally promised: undoing an insert that landed onto a slide means putting
 * the REPLACED slide back, and the only way to put a slide back is to hand
 * PowerPoint a package containing it. Ten of those is ten copies of the user's
 * presentation held in a task-pane WebView, which is the memory that killed a
 * sibling's run. PowerPoint's own Ctrl+Z reverts an insert — measured on the
 * web on 2026-09-10, `docs/host-answers/` — and that is the deeper history.
 */
interface Undoable {
  target: "onto" | "new";
  /** Which slide the insert was aimed at, counting from ZERO. */
  index: number;
  /** The deck as it was before the insert. */
  before: string;
  name: string;
  /** The catalogue id, so "Used in this deck" can drop it again when this is undone. */
  id: string;
  /** Which slide it landed on, counting from ONE, for the same reason. */
  landedOn: number;
  /**
   * Whether "Used in this deck" already had this element on that slide.
   *
   * Asked before the insert, because afterwards nothing can tell: the list
   * keeps one number per slide, so a second copy leaves the row unchanged.
   * Undo restores the slide as it was BEFORE this insert, which still holds
   * the earlier copy — so the row has to stay.
   */
  alreadyThere: boolean;
  /** What the pane knew the destination slide held BEFORE the insert. */
  onSlide: PaneState["onSlide"];
}
let undoable: Undoable | undefined;

/**
 * How many times this pane has changed the deck.
 *
 * "Used in this deck" reads the WHOLE deck, which is section 13's sixth open
 * question and unmeasured on a fifty-megabyte one — and the pane stays usable
 * while it runs, deliberately (`render.ts`, and `docs/DESIGN.md` section 4).
 * So an insert, an undo or a removal can finish DURING that read, and the read
 * then lands with an answer from before it and overwrites `used` and
 * `onSlide` — silently putting back a list that is missing the element the
 * user just watched land.
 *
 * A counter rather than a lock, because a lock is the fix the record forbids:
 * it would make the pane unusable for the length of an unmeasured read. The
 * read notes this number before it starts and throws its own answer away if it
 * has moved.
 *
 * Bumped when an operation BEGINS, not when it is confirmed to have worked,
 * and the direction is the whole reason. `CLAUDE.md` records both halves of
 * why a confirmation cannot be trusted here: a queued call that raises nothing
 * has not necessarily happened, and a call can raise and still have done the
 * work. An insert that threw halfway may have landed; a removal that broke at
 * its third slide changed two. Counting from the start over-invalidates a read
 * that overlapped an operation which turned out to do nothing, and the cost of
 * that is one more click. Counting from success under-invalidates exactly the
 * cases the host is documented to produce, and the cost of that is the bug.
 */
let deckEdits = 0;

function root(): HTMLElement {
  const node = document.getElementById("pane");
  if (!node) throw new Error("the pane's root element is missing");
  return node;
}

/**
 * The live region, made once and never rebuilt.
 *
 * `render` empties `#pane` and builds fresh elements on every draw, and a live
 * region CREATED with its content in it does not announce — the region has to
 * exist first and have text put into it. So this one lives outside the pane, is
 * made on the first draw, and is only ever written to.
 *
 * It is off-screen rather than `display: none`, which would take it out of the
 * accessibility tree along with everything in it.
 */
function liveRegion(): HTMLElement {
  const existing = document.getElementById("announcer");
  if (existing) return existing;
  const node = document.createElement("p");
  node.id = "announcer";
  node.className = "visually-hidden";
  // `polite`, never `assertive`: nothing here is urgent enough to cut across
  // what the user is already being told.
  node.setAttribute("role", "status");
  node.setAttribute("aria-live", "polite");
  document.body.append(node);
  return node;
}

/** The last thing announced, so the same sentence is not said twice. */
let announced = "";

function announce(text: string): void {
  // Only on a CHANGE: writing the same string back into a live region makes
  // some screen readers say it again.
  if (text === announced) return;
  announced = text;
  liveRegion().textContent = text;
}

/** Where the search caret was, so a redraw does not throw it away. */
/**
 * A selector that finds the same control again after a redraw.
 *
 * `render` empties the pane and builds fresh elements, so nothing survives by
 * reference — a control has to be found again by what it IS. The three data
 * attributes are what distinguish one: the action, the element it is for, and
 * which of the up-to-three tiles for that element this one is.
 *
 * Answers undefined for anything with no action, which is everything the pane
 * does not own.
 */
function focusKey(el: Element): string | undefined {
  if (!(el instanceof HTMLElement)) return undefined;
  const action = el.dataset["action"];
  if (action === undefined) return undefined;
  const parts = [`[data-action="${CSS.escape(action)}"]`];
  for (const name of ["id", "where", "value"] as const) {
    const value = el.dataset[name];
    if (value !== undefined) parts.push(`[data-${name}="${CSS.escape(value)}"]`);
  }
  // The three attributes do not always tell two controls apart. The gear is
  // drawn TWICE — the ⚙ button at the top and the settings line in the footer —
  // and both carry `data-action="gear"` and nothing else, so pressing the
  // footer line put the focus on the button at the top of the pane, past the
  // search box, the chips and the whole list. `:nth-of-type` cannot help,
  // because the two are not siblings; the position among the matches is
  // appended instead, and only when there is more than one to tell apart.
  const selector = parts.join("");
  const all = [...root().querySelectorAll(selector)];
  const at = all.indexOf(el);
  return all.length > 1 && at >= 0 ? `${selector}\u0000${at}` : selector;
}

/**
 * The control a `focusKey` names, if the redraw still has it.
 *
 * The position after the `\u0000` is only there when `focusKey` found more than
 * one match, so an ordinary key is still a plain selector.
 */
function focusedBy(key: string): HTMLElement | null {
  const [selector, at] = key.split("\u0000");
  if (selector === undefined) return null;
  const all = [...root().querySelectorAll<HTMLElement>(selector)];
  return (at === undefined ? all[0] : all[Number(at)]) ?? null;
}

/** Whether `draw` is putting the focus back, rather than the user moving it. */
let restoringFocus = false;

/**
 * The control that OPENED the tile menu, and the one that opened the gear.
 *
 * A `focusKey`, taken at the moment the surface opens, so closing it can put
 * the focus back where the user was rather than on `<body>`.
 *
 * `draw`'s restore cannot do this on its own: the control the focus was on is
 * INSIDE the surface — a menu item, a gear choice — and the redraw that closes
 * it legitimately removes that control, which is exactly the case `draw` hands
 * to the browser's fallback. That fallback is right for a tile a search
 * filtered away, which has no owner to go back to. It is wrong for a dismissed
 * surface, which has exactly one and it is still on screen. Measured in jsdom:
 * Shift+F10 on a tile, Tab into the menu, Escape — `document.activeElement`
 * came back `<body>`, so the next Tab restarted at the top of the pane.
 *
 * The gear needs one of its own because it is drawn TWICE, as the ⚙ above the
 * list and as the settings line in the footer, and both carry
 * `data-action="gear"` and nothing else; `focusKey` is what tells them apart.
 */
let menuOwner: string | undefined;
let gearOwner: string | undefined;

/**
 * A control the NEXT redraw should focus, rather than restoring what was there.
 *
 * For the one case where the right answer is not "put it back": opening the
 * removal question removes the button that opened it from every tile, so there
 * is nothing to put it back on. Consumed by `draw` and cleared whether or not
 * the control turns up.
 */
let focusAfterDraw: string | undefined;

/**
 * A control the restore FOUND but could not focus yet, kept for the next draw.
 *
 * `render` disables every tile while `state.busy`, and `focus()` on a disabled
 * button is a no-op — measured on jsdom 30 and true in every browser. `insert`
 * sets `busy` synchronously, so the very first redraw of an insert rebuilt the
 * tile the user had just pressed Enter on, disabled, and the restore quietly
 * did nothing. The old node was already detached by `render`, so focus fell to
 * `<body>` — and stayed there, because the NEXT draw found `activeElement` on
 * body, outside `root()`, and so had no `held` to restore at all.
 *
 * The result was that `docs/DESIGN.md` section 9's "the arrow keys move between
 * the tiles" stopped working after the first insert of a session: the next
 * ArrowDown reached `arrowTo` with an index of -1, which clamps to 0, and the
 * user was back at the top of the library. The `held` machinery exists to stop
 * a redraw throwing focus on the floor, and it was failing for the one redraw a
 * user causes most.
 *
 * So a target that exists but cannot take focus is REMEMBERED rather than
 * dropped, and the next draw that can focus it does. Cleared the moment it is
 * used or the control goes away, so it cannot pull the focus back to a tile the
 * user has since left.
 */
let deferredFocus: string | undefined;

function draw(): void {
  const active = document.activeElement;
  const wasSearch = active instanceof HTMLInputElement && active.dataset["action"] === "search";
  const caret = wasSearch ? active.selectionStart : null;
  /**
   * What else held the focus, so the redraw does not throw it on the floor.
   *
   * `docs/DESIGN.md` section 9 promises the arrow keys move between the tiles
   * and Enter inserts the one you are on, and none of it worked: focus landing
   * on a tile calls `set({chosen})`, `set` redraws, and `render` starts by
   * emptying the pane — so the button that had just taken focus was removed
   * from the document and focus fell back to `<body>`. Tab could not get past
   * the first tile, and every arrow after that landed on tile 0 forever,
   * because `indexOf(document.activeElement)` was -1.
   *
   * The caret restore above is the same repair for the search box, written
   * when the same redraw ate the caret. This is the rest of the pane.
   */
  const held = !wasSearch && active && root().contains(active) ? focusKey(active) : undefined;

  render(root(), state, stepFor(state));
  liveRegion();
  // The notice when there IS one, and otherwise the region is left alone.
  //
  // This announced `state.notice ?? ""`, so every redraw with no notice — which
  // is nearly all of them — wiped the region. An outcome announced just before
  // one of those was gone before anything could read it, which is why section
  // 9's "a live region announces every outcome" held only for an outcome that
  // happened to be the last word the pane said. Blanking bought nothing
  // either: a live region is read on CHANGE, so stale text sits there unread.
  if (state.notice !== undefined) announce(state.notice);

  if (wasSearch) {
    const search = root().querySelector<HTMLInputElement>('[data-action="search"]');
    if (search) {
      search.focus();
      if (caret !== null) search.setSelectionRange(caret, caret);
    }
  } else if (focusAfterDraw !== undefined) {
    const key = focusAfterDraw;
    focusAfterDraw = undefined;
    // Flagged like the restore below: `focus()` raises `focusin`, and `onFocus`
    // reads that as the user arriving at a tile.
    restoringFocus = true;
    try {
      focusedBy(key)?.focus();
    } finally {
      restoringFocus = false;
    }
  } else if (held !== undefined || deferredFocus !== undefined) {
    // Only when it is still there: a control the redraw legitimately removed —
    // a tile a search filtered away — is not something to hunt for, and the
    // browser's own fallback is right for it. A DISMISSED SURFACE is not that
    // case, which is what `menuOwner` and `gearOwner` above are for: a menu
    // item or a gear choice is removed by the redraw too, but it has one owner
    // and the owner is still on screen, so the close hands `focusAfterDraw`
    // that owner rather than letting the fallback stand.
    //
    // Flagged, because `focus()` raises `focusin` and `onFocus` treats that as
    // the USER arriving at a tile: it marks the tile chosen and arms the
    // preview. Putting the focus back where it already was is neither. Without
    // this the restore armed a third-of-a-second timer on every redraw that
    // had a tile focused, so the pane kept redrawing itself long after
    // anything had happened — measured as an insert's whole result being
    // painted over by a later draw, outcome and Undo and all.
    // `held` when the user was on something; otherwise the one the last draw
    // could not focus. `deferredFocus` is only consulted when nothing held the
    // focus, which is exactly the state the failed restore leaves behind.
    const key = held ?? deferredFocus;
    const target = key === undefined ? null : focusedBy(key);
    const blocked = target instanceof HTMLButtonElement && target.disabled;
    deferredFocus = blocked ? key : undefined;
    if (target !== null && !blocked) {
      restoringFocus = true;
      try {
        target.focus();
      } finally {
        restoringFocus = false;
      }
    }
  }
  restoreScroll();
}

function set(changes: Partial<PaneState>): void {
  state = { ...state, ...changes };
  draw();
}

// ---------------------------------------------------------------------------
// What the pane remembers: some of it per machine, the rest per deck.
// ---------------------------------------------------------------------------

/**
 * The deck's own bucket, decided once at boot.
 *
 * Once rather than per write, because the answer cannot change while the pane
 * is open — a task pane is torn down and rebuilt when the user opens another
 * presentation — and because `Office.context.document.url` is a host call, and
 * one per keystroke in the search box is a cost for nothing.
 */
let deckBucket = GLOBAL_KEY;

/**
 * One bucket, or nothing when it cannot be had.
 *
 * Wrapped because storage is not always there: a WebView with site data blocked
 * throws on the accessor itself rather than answering empty, and a pane that
 * will not open because it could not remember a search box is worse than one
 * that forgets. WHAT is in a bucket and which half it belongs to is
 * `src/pane/storage.ts`; this is the reading, which is all that has to be here.
 */
function read(key: string): Partial<PaneState> {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return {};
    return JSON.parse(raw) as Partial<PaneState>;
  } catch {
    return {};
  }
}

function remembered(): Partial<PaneState> {
  deckBucket = deckKey(deckUrl());
  const machine = read(GLOBAL_KEY);
  // The same object when the host would not name a deck: one bucket holds both
  // halves, and `restored` treats that as the ordinary case rather than one to
  // branch on.
  const deck = deckBucket === GLOBAL_KEY ? machine : read(deckBucket);
  // Not part of the state, so it is taken here rather than returned.
  keptScroll = storedScroll(deck);
  // Read HERE and kept, because `restored` cannot carry it: the answer is
  // whether a field was absent, and what it returns is a state where that field
  // is a list either way. `load` uses it once, when the library arrives.
  newDeck = firstVisit(deck);
  return restored(machine, deck);
}

/**
 * Whether this deck has never been seen, so the first category can be opened.
 *
 * Alongside `keptScroll` rather than in the state, and for the same reason: it
 * is a fact about the READ, true once, and a state field would have to be kept
 * correct for the rest of the session by everything that touches it.
 */
let newDeck = false;

/**
 * How far down the list the user had scrolled.
 *
 * Kept OUTSIDE `PaneState` on purpose: an offset in the state means a re-render
 * per scroll event, and re-rendering the whole list while it moves under the
 * user's finger is the one thing this feature must not cost. WHEN it goes back
 * is `shouldRestoreScroll` in `src/pane/storage.ts`; the DOM question it needs
 * answering — are there tiles yet — is the one thing only this file can ask.
 */
let keptScroll = 0;
let scrollRestored = false;

function restoreScroll(): void {
  const hasTiles = root().querySelector('[data-action="tile"]') !== null;
  if (!shouldRestoreScroll({ restored: scrollRestored, kept: keptScroll, hasTiles })) return;
  scrollRestored = true;
  window.scrollTo(0, keptScroll);
}

/**
 * Remember where the list is, without a render and without a write per event.
 *
 * A scroll fires tens of times a second and `keep` is a `setItem`, so the write
 * trails by a quarter of a second. A pane torn down inside that window loses up
 * to that much scrolling, which is a few pixels of where the user was — the
 * only thing this stores.
 */
let scrollTimer: ReturnType<typeof setTimeout> | undefined;

function onScroll(): void {
  keptScroll = window.scrollY;
  // Once the user has scrolled, the restore has had its chance: putting the old
  // offset back after that would undo the scroll that just happened.
  scrollRestored = true;
  if (scrollTimer !== undefined) clearTimeout(scrollTimer);
  scrollTimer = setTimeout(keep, 250);
}

function write(key: string, value: unknown): void {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // A pane that cannot remember is still a pane that works.
  }
}

function keep(): void {
  for (const [key, value] of writes(deckBucket, state, keptScroll)) write(key, value);
}

// ---------------------------------------------------------------------------
// Loading the library.
// ---------------------------------------------------------------------------

/**
 * The deck's slide size, read from the file, because no API answers it — and,
 * from the same read, what the slide the user is on already holds.
 *
 * Both out of ONE deck read. The size is what the pane needs to pick a library;
 * the boxes are what the preview card draws in grey (`docs/DESIGN.md` sections
 * 1 and 4). Reading the deck twice for them would be paying the sixth open
 * question's unmeasured cost twice on every open.
 */
async function deckShape(): Promise<{ width: number; height: number; slide?: number; boxes?: Box[] } | undefined> {
  try {
    const deck = await readDeck();
    const pkg = await Pkg.open(deck.base64);
    const size = await slideSize(pkg);
    const held = await heldBy(pkg, size);
    return { ...size, ...held };
  } catch {
    return undefined;
  }
}

/**
 * What the slide the user is on holds, out of an already-open package.
 *
 * Answers nothing rather than throwing, and nothing rather than guessing when
 * the host will not say which slide that is: grey boxes for the wrong slide
 * would be a wrong answer to the only question they exist to answer.
 */
async function heldBy(pkg: Pkg, size: { width: number; height: number }): Promise<{ slide?: number; boxes?: Box[] }> {
  try {
    const current = await currentSlide();
    if (current == null) return {};
    const path = (await pkg.slidePaths())[current.index];
    if (path === undefined) return {};
    return { slide: current.index + 1, boxes: occupiedBoxes(await pkg.doc(path), size.width, size.height) };
  } catch {
    return {};
  }
}

async function load(): Promise<void> {
  set({ problem: undefined, library: undefined });
  try {
    index = await loadIndex();
  } catch (e) {
    set({
      problem: `The library did not load: ${readable(e)}. That is usually the network rather than a fault in the add-in.`,
    });
    return;
  }
  // The index first, so tiles are on screen while the deck is being measured.
  // 16:9 is what almost every modern deck is, and the shape read below corrects
  // it — visibly, through the line under the header — when it is not.
  const provisional = libraryFor(index, 12192000, 6858000);
  // On a deck the pane has never seen, open the first category, so the first
  // screen carries elements rather than a column of shut headings. Once only,
  // and never against a deck that remembered something — including a deck whose
  // user shut everything, which `firstVisit` is what tells apart.
  const firstSight = newDeck;
  set({
    library: provisional,
    ...(firstSight ? { open: openAtFirst(provisional, state) } : {}),
  });
  newDeck = false;
  store = new Store(provisional.size);

  const shape = await deckShape();
  if (!shape || !index) {
    // The deck read failing is not a reason to stop following the SELECTION.
    // The two are independent — one reads the file, the other subscribes to an
    // event — and this early return took the subscription with it for the rest
    // of the session, so the line under the header stayed on whatever it said
    // at boot however much the user clicked. Silently: nothing on screen
    // distinguishes a line that is right from one that stopped being asked.
    // `follow` is idempotent, which is what lets this stand beside the call at
    // the end without a second subscription.
    void follow();
    return;
  }
  const library: Library = libraryFor(index, shape.width, shape.height);
  if (library.size !== provisional.size || library.borrowed !== provisional.borrowed) {
    store = new Store(library.size);
    // Choose the open category AGAIN, against the library that is real. The one
    // chosen above came from the provisional 16:9 library, and the two libraries
    // do not carry the same categories — so on a deck of another shape that key
    // can name a category this library has not got, and the first screen goes
    // back to the column of shut headings the open exists to prevent, with
    // nothing on screen saying why. Today both libraries happen to start with
    // the same category, which is luck and not a rule.
    set({ library, ...(firstSight ? { open: openAtFirst(library, state) } : {}) });
  }
  const current = await currentSlide();
  set({
    slide: current == null ? undefined : current.index + 1,
    deck: { width: shape.width, height: shape.height },
    ...(shape.slide !== undefined && shape.boxes ? { onSlide: { slide: shape.slide, boxes: shape.boxes } } : {}),
  });
  void follow();
}

/**
 * Keep the line under the header naming the slide the user is actually on.
 *
 * Coalesced rather than guarded by a flag: a flag DROPS the event that arrives
 * while a read is running, and on the web that read takes long enough for a
 * second click to fall inside it. Measured on 2026-09-11 — slides 2, 3 and 4
 * clicked a quarter of a second apart left the pane saying "Slide 3." with no
 * later event coming to put it right. `src/host/coalesce.ts` carries the rule.
 *
 * Nothing is read while an insert is running: the host has a context open and
 * a read for a line nobody is looking at is a full presentation save for
 * nothing. `insert` and `undo` each ask for one when they finish instead.
 */
const followSelection = coalescing(async () => {
  if (state.busy === true) return;
  // A glance, not a read: `src/host/timeout.ts` says why the budget is its own.
  const current = await currentSlide(BUDGET.glance);
  // Only a host that ANSWERED may change this line. `null` is a read that ran
  // out of time, and the last number the host gave is a better answer to that
  // than "PowerPoint did not say".
  if (current === null) return;
  const slide = current === undefined ? undefined : current.index + 1;
  // Only when the number changed. `set` redraws the pane, and a redraw the
  // user did not ask for is a redraw that can take the focus off whatever they
  // were on.
  if (slide !== state.slide) set({ slide });
});

/** Subscribe once; a host that refuses simply keeps the line it has. */
async function follow(): Promise<void> {
  if (followed) return;
  followed = await onSlideChange(followSelection);
}
let followed = false;

/**
 * Go to a slide named in "Used in this deck".
 *
 * The number is turned into an id positionally, at the last moment
 * (`slideIdAt`, the same read the insert aims with), the host is asked to
 * select it, and the selection is read straight back. `jumpOutcome` decides
 * what that read means: the pane says "Slide N" only when the host was seen
 * there, and otherwise says which slide to click. Nothing about the deck
 * changes either way, so there is nothing to confirm by counting.
 */
async function jumpTo(slide: number): Promise<void> {
  if (state.busy === true || state.reading === true) return;
  if (!Number.isInteger(slide) || slide < 1) return;
  set({ notice: undefined });
  const wanted = await slideIdAt(slide - 1);
  if (wanted === undefined) {
    set({ notice: `PowerPoint would not name slide ${slide}. Click it in the strip.` });
    return;
  }
  const seen = await selectSlide(wanted);
  const outcome = jumpOutcome({ slide, wanted, ...seen });
  if (!outcome.ok) {
    set({ notice: outcome.detail });
    return;
  }
  // Seen there, so the line under the header may say so now; the selection
  // event, where the host raises one, will agree with it.
  set({ slide });
  announce(outcome.detail);
}

// ---------------------------------------------------------------------------
// Inserting.
// ---------------------------------------------------------------------------

/**
 * Insert an element, optionally onto the target the gear is NOT set to.
 *
 * `once` is section 6's right-click: the other target for this one insert,
 * without touching the setting. It is threaded through everything downstream
 * that asks where the element went — the splice, whether the replaced slide is
 * removed, what the footer says, and what Undo puts back — because those four
 * have to agree, and the setting is no longer the answer for this call.
 */
async function insert(id: string, once?: "onto" | "new"): Promise<void> {
  const library = state.library;
  const element = elementOf(library, id);
  if (!library || !element || !store || !index || state.busy === true) return;
  /**
   * The store this insert will use, bound HERE, beside the library the element
   * was read from.
   *
   * `store` is a module-level `let` and `load()` REASSIGNS it, after
   * `deckShape()` has answered, to the size the deck turned out to be. That
   * await is the window in which the tiles are already on screen and clickable
   * — deliberately, so the pane is usable while the deck is measured — and
   * `deckShape` is a whole `getFileAsync`: `timeout.ts` records 874 ms on a
   * healthy web session and 40 s on a degraded one.
   *
   * Both uses below read the module-level binding, and the second is several
   * awaits deep inside the splice. So an insert begun against the provisional
   * 16:9 library could fetch its carried parts from `catalogue/4x3/parts/…`
   * once the swap landed. 47 part paths exist under BOTH sizes, and 44 of them
   * hold DIFFERENT bytes — measured on the committed catalogue, 2026-09-23 —
   * so the user's deck got the other library's chart, workbook or picture with
   * no error at all; the paths that exist under one size only answered 404 and
   * raised "the catalogue has no part …", naming a part the catalogue has.
   *
   * Bound once, the insert is CONSISTENT: markup and parts come from the same
   * library. Whether the tiles should be clickable against the provisional
   * library at all is a separate question and the owner's.
   */
  const parts = store;
  // A PART ignores the insert target, whether the target came from the gear or
  // from the right-click. `docs/DESIGN.md` section 5: "A part ignores the
  // insert target: it always lands on the slide the user is on"; section 7
  // says it again for the menu. `offersOtherTarget` keeps the menu to that
  // rule and nothing kept the SETTING to it, so a gear left on "As a new
  // slide" sent a stamp through the splice's blanking path: the stamp landed
  // alone on an empty slide after the user's, the user's own slide kept
  // nothing, the removal was skipped because the target was not "onto", and
  // the footer reported plain success over it.
  const target = element.kind === "part" ? "onto" : (once ?? state.settings.target);

  // `noQuestion`: an insert rewrites `state.recent` through `remember`, which
  // DROPS the oldest id once the list is six long — so a question open on that
  // element's Recent tile was left drawn nowhere, with the Remove button
  // suppressed on every tile because one was notionally open and nothing on
  // screen saying why. The rule `noQuestion` states is the general one, and an
  // insert is a state change like any of the others that spread it.
  set({
    busy: true,
    busyWith: "insert",
    chosen: id,
    notice: INSERTING,
    outcome: undefined,
    menuFor: undefined,
    ...noQuestion,
  });
  // Before the first await: from here on, any deck read running underneath this
  // is reading a deck this pane is in the middle of changing.
  deckEdits += 1;
  /** Whether `insertSlidesFromBase64` was reached. See the catch at the end. */
  let asked = false;
  try {
    const markup = await parts.markup(element);
    const deck = await readDeck();
    // Section 5's several-slide stamp. Read only for a PART, because that is
    // the only kind the record gives it to, and only after the two expensive
    // fetches above so an ordinary one-slide insert pays nothing extra for a
    // selection read it will not use. `stampTargets` answers the empty list for
    // anything under two slides, which is what hands the ordinary path back.
    const many = element.kind === "part" ? stampTargets(await selectedSlides()) : [];
    if (many.length > 1) {
      await stampEvery(element, { markup, deck: deck.base64, library, index, parts }, many);
      return;
    }
    const current = await currentSlide();
    const at = current?.index ?? 0;
    const before = await slideCount();
    const selection = element.landing === "cursor" ? await selectedShape() : undefined;

    const report = await splice({
      deck: deck.base64,
      slide: at,
      element: {
        id: element.id,
        name: element.name,
        kind: element.kind,
        box: element.box,
        landing: element.landing,
        ...(wrapsSelection(element) ? { wraps: true } : {}),
        markup: { xml: markup.xml, rels: markup.rels },
      },
      options: { ...state.settings, target },
      catalogue: {
        version: library.version,
        carried: carriedTypes(index, library.size),
        theme: themeColours(index, library.size),
        // The LIBRARY's own slide, not this deck's: the element's shapes are
        // drawn in those units, and the splice rebases them onto whatever this
        // deck measures. `libraryFor` answers the catalogue's dimensions, so
        // these are the library's even when the deck borrowed it.
        width: library.width,
        height: library.height,
      },
      store: (path) => parts.part(path),
      ...(selection ? { selection } : {}),
    });

    const targetId = current?.id;
    if (targetId === undefined) {
      set({
        busy: false,
        busyWith: undefined,
        notice: undefined,
        outcome: {
          ok: false,
          byHand: false,
          name: element.name,
          detail:
            "PowerPoint would not say which slide you are on, so nothing was inserted. Click a slide and try again.",
        },
      });
      return;
    }

    // From here the deck may have changed, whatever happens next. The catch
    // below needs to tell that apart from a failure before this line, because
    // the two want opposite things said and opposite things done with the undo
    // that is already armed.
    asked = true;
    const error = await insertPackage(report.base64, targetId);
    // Asked again until it agrees, because the count lags the insert: see
    // `countReaching`. One read here would report a landed insert as a no-op.
    const inserted = await countReaching(before + 1);
    let removed: number | undefined;
    let moved = false;
    if (target === "onto" && mayRemove({ before, inserted })) {
      // `mayRemove` asks only about the count, and a count cannot see a
      // REORDER. `at` was read before the host calls and the insert can take up
      // to `BUDGET.insert`; the pane locks itself, not PowerPoint, so the user
      // can drag a slide in the strip in that window and the count will not
      // move. The insert aims by id and survives it; this delete aims by
      // position and does not. So the id is read back and compared before
      // anything is deleted, and a mismatch — or a read that does not answer —
      // leaves the copy standing, which is the failure mode `CLAUDE.md` asks
      // for: a duplicate the user can delete rather than a slide they lost.
      moved = !stillThere(current?.id, await slideIdAt(at));
    }
    if (target === "onto" && !moved && mayRemove({ before, inserted })) {
      // The rebuilt slide landed AFTER the original, so the original is still
      // at its own index. Positional, never by id: a slide next to one the run
      // has just added is exactly where an id read is not to be trusted.
      await removeSlideAt(at);
      // The DELTA decides, in both arms, and the raise decides nothing.
      //
      // This read the raise instead: a refusal set `removed = inserted`, so a
      // delete that raised AND LANDED — `CLAUDE.md`'s "a call can raise and
      // still have done the work", measured on a sibling's insert that timed
      // out with both slides in — came out as "The deck grew by one but the
      // copy could not be removed: delete slide N by hand." over a deck that
      // was already correct. A user who follows that instruction deletes the
      // slide the element is now on, which is their own content, and the pane
      // told them to.
      //
      // `outcomeOf` already answers `byHand` correctly when the count really
      // is still `before + 1`, so the honest failure keeps its sentence and
      // only the false one goes.
      removed = await countReaching(before);
    }

    const outcome = outcomeOf({
      target,
      slide: at + 1,
      before,
      inserted,
      ...(removed === undefined ? {} : { removed }),
      ...(moved ? { moved } : {}),
      ...(error === undefined ? {} : { error }),
    });
    // Where the element ended up, counting from one. `landedOn` is `undoPlan`
    // read from the other end, and lives beside it for that reason.
    const landed = landedOn({ target, index: at });
    undoable = outcome.ok
      ? {
          target,
          index: at,
          before: deck.base64,
          name: element.name,
          id: element.id,
          landedOn: landed,
          alreadyThere: target === "onto" && holds(state.used, element.id, landed),
          onSlide: state.onSlide,
        }
      : undefined;
    state = {
      ...state,
      busy: false,
      busyWith: undefined,
      // Section 6's "Move to a new slide". `report.held` is what the slide the
      // user was on already carried, counted by the splice out of the bytes it
      // was already holding, so the offer costs no second read and no host call.
      moveable: moveableAfter({ ok: outcome.ok, target, held: report.held }, element),
      recent: outcome.ok ? remember(state.recent, element.id, RECENT_DEPTH) : state.recent,
      undo: outcome.ok ? 1 : 0,
      outcome: { ...outcome, name: element.name },
      // "Used in this deck" is updated rather than re-read: the pane knows
      // exactly what it just put where, and the read is the expensive half of
      // this feature. Untouched when nothing has been read — an insert is not a
      // reason to start claiming the deck has been looked at.
      // Renumbered FIRST when the deck grew, then the new row added at the
      // number it actually landed on. "As a new slide" pushes every element at
      // or after that point one slide along, and those numbers are jump
      // controls — a row left saying "slide 5" sends the user to whatever
      // slide 5 has become. "Onto this slide" is net zero and shifts nothing.
      used: outcome.ok
        ? withInsert(target === "new" ? renumbered(state.used, landed, 1) : state.used, element.id, landed)
        : state.used,
      // The card's grey boxes keep up the same way: the splice says where the
      // element landed, in EMU **on the user's slide**, so it is the USER's
      // slide size that turns it into a fraction. The library's size is a
      // different number whenever a deck borrowed the nearest library, and
      // dividing by that one draws the right rectangle in the wrong place on
      // exactly the decks nobody tests on.
      onSlide: outcome.ok
        ? withLanded(state.onSlide, landed, fractionOf(report.landed, state.deck), target === "new")
        : state.onSlide,
    };
    delete state.notice;
    keep();
    draw();
    announce(announcement(outcome, element.name));
    // The insert may have moved the selection, and nothing was read while it
    // ran. Ask once now.
    followSelection();
  } catch (e) {
    // Two different failures reach here, and only one of them leaves the deck
    // alone. Before `insertSlidesFromBase64` was reached — the markup fetch,
    // the deck read, the splice — nothing was asked of the host, "refused" is
    // honest, and an undo armed by an EARLIER insert still describes the deck
    // and must survive.
    //
    // After it, the slide may be in the deck: the two calls that can throw
    // there are count reads, and `withTimeout` rejects on the budget as
    // readily as on a host raise. Saying "refused" then tells the user nothing
    // changed while a rebuilt slide sits in their deck with the original still
    // beside it — and the armed undo now points at the wrong insert, so
    // pressing it would put a slide back against a deck this code has already
    // misread. It is disarmed rather than left to do that.
    if (asked) {
      undoable = undefined;
      state = {
        ...state,
        busy: false,
        busyWith: undefined,
        undo: 0,
        moveable: undefined,
        outcome: {
          ok: false,
          byHand: true,
          name: element.name,
          detail: `The insert did not confirm: ${readable(e)}. Check the end of the deck.`,
        },
      };
    } else {
      state = {
        ...state,
        busy: false,
        busyWith: undefined,
        outcome: { ok: false, byHand: false, name: element.name, detail: `The insert was refused: ${readable(e)}` },
      };
    }
    delete state.notice;
    draw();
    // `docs/DESIGN.md` section 9: "A live region announces every outcome." The
    // success path a few lines up does; neither failure path did — so the one
    // case where the deck may be holding a slide too many was the one a screen
    // reader was told nothing about.
    announce(state.outcome?.detail ?? "");
  }
}

/**
 * Put the deck back the way it was.
 *
 * Count-checked at every step, and positional throughout. Nothing here reads an
 * id: `CLAUDE.md` records that a slide the run just added does not resolve by
 * one on the web, and undo is working right next to one.
 */
async function undo(): Promise<boolean> {
  const entry = undoable;
  if (!entry || state.busy === true) return false;
  set({ busy: true, busyWith: "undo", notice: "Undoing…" });
  deckEdits += 1;
  const plan = undoPlan(entry);
  /**
   * Whether the host has been asked to change the deck yet. See the catch.
   *
   * The same flag `insert` keeps, for the same reason and one step worse here:
   * this function puts the user's original slide back FIRST and takes the
   * rebuilt one away second, so a failure between the two leaves the deck
   * holding both.
   */
  let asked = false;
  try {
    const before = await slideCount();

    if (plan.after !== undefined) {
      // Put the user's own slide back first, aimed at the rebuilt one so it
      // lands immediately after it whatever the selection is now.
      const targetId = await slideIdAt(plan.after);
      if (targetId === undefined) {
        throw new Error(`PowerPoint would not name slide ${plan.after + 1}, so the original could not be put back`);
      }
      const original = await onlySlide(entry.before, entry.index);
      // From here the deck may hold the user's original slide again, whatever
      // happens next.
      asked = true;
      const refused = await insertPackage(original.base64, targetId);
      const grown = await countReaching(plan.grownTo(before));
      if (grown !== plan.grownTo(before)) {
        throw new Error(
          refused ?? `the deck went ${before} → ${grown}, which is not what putting one slide back looks like`,
        );
      }
    }

    asked = true;
    const refused = await removeSlideAt(plan.remove);
    const want = before - (plan.after === undefined ? 1 : 0);
    const after = await countReaching(want);
    // The count decides and the raise only supplies the MESSAGE, which is what
    // the insert half ten lines above already does. This half tested `refused`
    // first, so a delete that raised and landed threw over a deck that was
    // correctly restored: the pane said "Undo did not work", `undoable` was
    // never cleared, and "Move to a new slide" — which runs this undo and then
    // inserts again — stopped after the undo. The element the user asked to
    // MOVE was gone from the deck altogether, under a message saying the undo
    // had failed.
    if (after !== want) {
      throw new Error(refused ?? `the deck has ${after} slides, which is not what was expected`);
    }

    undoable = undefined;
    state = {
      ...state,
      busy: false,
      busyWith: undefined,
      // `moveable` is deliberately NOT cleared here. `footerOf` gates the offer
      // on `undo` as well, so zeroing this is what makes it go, and a second
      // line saying the same thing would be one nothing could catch — a line
      // the suite cannot make go red is a line that rots quietly.
      undo: 0,
      outcome: { ok: true, byHand: false, name: entry.name, detail: `Undone. The deck has ${after} slides.` },
      // The slide it was on is the user's own again, so whatever the insert put
      // there went with it — both in the list and in the card's grey boxes.
      // The mirror: the row for the slide that went, then everything after it
      // back down by one. Only for a new slide — undoing an "onto" puts the
      // user's own slide back in place and moves nothing.
      used:
        entry.target === "new"
          ? renumbered(withoutInsert(state.used, entry.id, entry.landedOn), entry.landedOn + 1, -1)
          : entry.alreadyThere
            ? state.used
            : withoutInsert(state.used, entry.id, entry.landedOn),
      onSlide: entry.onSlide,
    };
    delete state.notice;
    draw();
    announce(`${entry.name} taken back.`);
    followSelection();
    return true;
  } catch (e) {
    // Disarmed once the host has been asked, because the entry no longer
    // describes the deck.
    //
    // `undoPlan` is built from the insert's own index, and this function
    // restores BEFORE it removes — so a failure in between leaves the deck
    // holding the original and the rebuilt slide both. Pressing Undo again
    // then re-runs the whole plan against that deck: the original goes in a
    // second time, the rebuilt one comes out, the counts agree, and the pane
    // reports "Undone." over two copies of the user's own slide with the
    // element gone.
    //
    // Before the host was asked — `slideIdAt` refusing, `onlySlide` throwing —
    // nothing changed and the entry is still exactly right, so it stays.
    if (asked) {
      undoable = undefined;
      state = {
        ...state,
        busy: false,
        busyWith: undefined,
        undo: 0,
        outcome: {
          ok: false,
          byHand: true,
          name: entry.name,
          detail: `Undo did not finish: ${readable(e)}. Check the deck around slide ${entry.landedOn}.`,
        },
      };
    } else {
      state = {
        ...state,
        busy: false,
        busyWith: undefined,
        outcome: { ok: false, byHand: true, name: entry.name, detail: `Undo did not work: ${readable(e)}` },
      };
    }
    delete state.notice;
    draw();
    // Announced, for the reason the insert's catch gives.
    announce(state.outcome?.detail ?? "");
    return false;
  }
}

/**
 * Put the last insert on a new slide instead (`docs/DESIGN.md` section 6).
 *
 * A whole-slide element that landed on a slide which already had something on
 * it covers that something. The offer is the way back out, and it is built from
 * the two operations the pane already has rather than a third: take the insert
 * back, then make it again with the other target. So it inherits both of their
 * guarantees — the undo is positional and count-checked, the second insert
 * reads the deck fresh and proves the delta — and it introduces no new way for
 * the deck to end up somewhere neither of them can describe.
 *
 * If the undo does not work the move stops there, with the undo's own sentence
 * on screen. The alternative is inserting a second copy beside the first, which
 * is the one outcome a user asking to MOVE something cannot have meant.
 */
async function moveToNewSlide(): Promise<void> {
  const id = state.moveable;
  if (id === undefined || state.busy === true) return;
  if (!(await undo())) return;
  await insert(id, "new");
}

// ---------------------------------------------------------------------------
// Wiring.
// ---------------------------------------------------------------------------

/**
 * Which control a click landed on, given whatever it actually hit.
 *
 * `Element`, not `HTMLElement`, and that is the whole of this function's
 * history. **An SVG element is not an HTMLElement**, and most of a tile IS an
 * SVG: the ghost drawing fills it. So a click anywhere on the picture — which
 * is where a user clicks — started the walk at `undefined` and resolved to no
 * action at all, and the pane sat there doing nothing. Every test passed: jsdom
 * dispatched its clicks on the button, and the audit never clicks anything.
 * Found by opening the add-in in PowerPoint and pressing a tile.
 *
 * `closest` rather than a hand-rolled walk, because it is defined on `Element`
 * and therefore crosses from the SVG into its HTML ancestors without caring
 * which is which.
 */
function actionOf(target: EventTarget | null): { action: string; el: HTMLElement } | undefined {
  const start = target instanceof Element ? target : undefined;
  const node = start?.closest("[data-action]");
  if (!(node instanceof HTMLElement)) return undefined;
  const action = node.dataset["action"];
  return action ? { action, el: node } : undefined;
}

/**
 * Open one of the site's own pages in the user's browser.
 *
 * The site is the one the PANE was served from, so a dev build's links point at
 * the dev origin and nothing here can send somebody to an origin the add-in did
 * not come from.
 *
 * When the host opens nothing — no `openBrowserWindow`, and a blocked
 * `window.open` — the pane says so and names the address, because a click that
 * silently does nothing is the version of this the user cannot work around.
 */
function leave(urlFor: (site: { origin: string }) => string, what: string): void {
  const site = siteFrom(window.location.href);
  if (!site) {
    set({ notice: `This pane was not opened from a web address, so it cannot open ${what}.` });
    return;
  }
  const url = urlFor(site);
  if (openExternal(url)) {
    set({ notice: undefined, gear: false });
    return;
  }
  set({
    notice: `PowerPoint would not open a browser window. You can reach ${what} at ${url.replace(/^https?:\/\//, "")}`,
  });
}

/**
 * Read the deck and say which library elements are already in it.
 *
 * `docs/DESIGN.md` section 4's "Used in this deck", off the tags an insert
 * writes — measured surviving `insertSlidesFromBase64` on 2026-09-11 (section
 * 15), which is what makes the feature buildable at all.
 *
 * On request rather than on open. It reads the WHOLE deck, and how long that
 * takes on a fifty-megabyte one is section 13's sixth open question: unanswered.
 * A pane that answered it on everybody's behalf, every time it opened, would be
 * spending an unmeasured cost on the people who never look at this list.
 */
async function readUsed(): Promise<void> {
  if (state.reading === true || state.busy === true) return;
  set({ reading: true, notice: undefined });
  const started = deckEdits;
  try {
    const deck = await readDeck();
    const pkg = await Pkg.open(deck.base64);
    const used = await usedInDeck(pkg);
    // The same read answers what the current slide holds, so the card's grey
    // boxes come back into step with the deck for free.
    const size = await slideSize(pkg);
    const held = await heldBy(pkg, size);
    if (deckEdits !== started) {
      // The pane changed the deck while this was running, so this answer is
      // about a deck that no longer exists. Writing it would take the element
      // the user just inserted back OUT of the list, and roll the card's grey
      // boxes back to the slide as it was — both silently, both wrong.
      //
      // Thrown away rather than merged. What is held now came from the
      // operation itself, which knew exactly what it did; merging a stale
      // sweep into it would be guessing which half to believe.
      set({ reading: false, notice: "The deck changed while it was being read. Ask again for an up-to-date list." });
      return;
    }
    set({
      reading: false,
      used: used.map((u) => ({ element: u.element, slides: u.slides })),
      ...(held.slide !== undefined && held.boxes ? { onSlide: { slide: held.slide, boxes: held.boxes } } : {}),
    });
  } catch (e) {
    // The list stays as it was — including "never asked" — rather than becoming
    // an empty one, because an empty list is a claim about the deck and this is
    // a failure to look at it.
    set({ reading: false, notice: `This deck could not be read: ${readable(e)}` });
  }
}

/**
 * Take an element off every slide it is on (`docs/DESIGN.md` sections 4 and 6).
 *
 * The same insert-then-remove cycle as an insert, once per slide, because that
 * is the only sequence this family has measured: hand PowerPoint a package
 * holding the rebuilt slide, prove the deck grew, then take the original away
 * and prove it shrank back. Nothing is deleted through the shape collection.
 *
 * **Every cycle is confirmed by the deck's own size before the next one
 * starts, and the first refusal stops the run.** A loop that pressed on after a
 * step it could not verify would be a loop editing a deck whose shape it has
 * already misread — and the user is told how far it got rather than left to
 * count slides.
 *
 * One deck read for the whole run. Each package is built from the bytes read at
 * the start, which is sound because a cycle only rewrites the slide it targets
 * and leaves every other slide's bytes as they were.
 *
 * NOT MEASURED: no round has run this against a real PowerPoint. The mechanism
 * is the insert's, which has been measured on the web and on Windows, but a
 * sequence of them has not. `docs/DESIGN.md` section 15 says so.
 */
/**
 * A stamp onto every selected slide (`docs/DESIGN.md` section 5).
 *
 * ONE CYCLE PER SLIDE, not one insert. The record said "in one insert" until
 * 2026-09-23 and that cannot be built: `insertSlidesFromBase64` puts every
 * slide of its package CONTIGUOUSLY after one `targetSlideId` — `CLAUDE.md`
 * records a real run that put 37 generated slides ahead of a title slide — so
 * rebuilt copies of slides 2, 5 and 9 would arrive in a block and the deck's
 * own order would be gone. Aiming each copy at its own slide keeps the order,
 * and that is one insert each.
 *
 * Which makes this `removeEverywhere` with a different payload, and everything
 * that path learned applies unchanged:
 *
 * - **One deck read for the whole run.** A cycle only rewrites the slide it
 *   targets, so every package is built from the bytes read at the start. The
 *   slides this run has already replaced are not among the ones still to come.
 * - **Net zero per cycle**, so the indices of later slides do not move: the
 *   copy lands after the original and the original is taken away. `many` is
 *   sorted for that reason.
 * - **The raise is not consulted; the DELTA is.** A call that raises can still
 *   have done the work, and a call that raises nothing may not have.
 * - **The positional delete is guarded by an id read back.** `at` was computed
 *   before the host calls and the pane locks itself rather than PowerPoint, so
 *   a user can drag a slide in the strip in that window. A mismatch leaves the
 *   copy standing, which is the failure `CLAUDE.md` asks for: a duplicate the
 *   user can delete rather than a slide they lost.
 *
 * **The pane's Undo is disarmed**, for the reason `removeEverywhere` gives:
 * `undoable` holds the whole deck as it was before an EARLIER insert, one
 * insert deep and positional, and it stops describing this deck at the first
 * cycle that lands. The footer says so rather than leaving the button silently
 * gone. PowerPoint's own Ctrl+Z does revert an insert — question 5, measured on
 * the web and on Windows — which is the route that does exist.
 */
async function stampEvery(
  element: LibraryElement,
  from: { markup: Markup; deck: string; library: Library; index: Index; parts: Store },
  many: number[],
): Promise<void> {
  set({ notice: `Stamping ${element.name} onto ${many.length} slides…` });
  let done = 0;
  /** Whether a cycle left its copy behind, which changes what may be said. */
  let stranded = false;
  try {
    for (const at of many) {
      const before = await slideCount();
      const targetId = await slideIdAt(at);
      if (targetId === undefined) break;
      const report = await splice({
        deck: from.deck,
        slide: at,
        element: {
          id: element.id,
          name: element.name,
          kind: element.kind,
          box: element.box,
          landing: element.landing,
          ...(wrapsSelection(element) ? { wraps: true } : {}),
          markup: { xml: from.markup.xml, rels: from.markup.rels },
        },
        // A part ignores the insert target, and with several slides selected
        // there is no "the slide you are on" to ignore it in favour of.
        options: { ...state.settings, target: "onto" },
        catalogue: {
          version: from.library.version,
          carried: carriedTypes(from.index, from.library.size),
          theme: themeColours(from.index, from.library.size),
          width: from.library.width,
          height: from.library.height,
        },
        store: (path) => from.parts.part(path),
      });
      await insertPackage(report.base64, targetId);
      if ((await countReaching(before + 1)) !== before + 1) break;
      if (!stillThere(targetId, await slideIdAt(at))) {
        // The copy landed and the slide it was aimed at has moved, so the
        // positional delete below would take somebody else's slide. Leaving the
        // copy is the whole point of insert-then-remove.
        stranded = true;
        break;
      }
      await removeSlideAt(at);
      if ((await countReaching(before)) !== before) {
        stranded = true;
        break;
      }
      done += 1;
    }
  } catch {
    // `done` is the number of cycles seen through to the end, and the outcome
    // below reports it. Whatever raised, the counts above are the evidence.
  }

  const outcome = stampOutcome(element.name, done, many.length, stranded);
  undoable = undefined;
  state = {
    ...state,
    busy: false,
    busyWith: undefined,
    moveable: undefined,
    recent: done > 0 ? remember(state.recent, element.id, RECENT_DEPTH) : state.recent,
    undo: 0,
    outcome,
    // The deck has changed under the pane on several slides at once, and both
    // of these describe one slide. Dropped rather than guessed: the next "See
    // what this deck already uses" is what puts them back.
    used: undefined,
    onSlide: undefined,
  };
  delete state.notice;
  keep();
  draw();
  announce(outcome.detail);
  followSelection();
}

async function removeEverywhere(id: string): Promise<void> {
  const plan = state.removing;
  const element = elementOf(state.library, id);
  if (!plan || !element || plan.id !== id || state.busy === true) return;
  set({
    busy: true,
    busyWith: "remove",
    notice: `Taking ${element.name} off ${plan.slides.length} slide(s)…`,
    outcome: undefined,
  });
  deckEdits += 1;

  let done = 0;
  /** Whether a cycle left its copy behind, which changes what may be said. */
  let stranded = false;
  let wanted = plan.slides;
  try {
    const deck = await readDeck();
    // Which slides carry it, asked of the bytes THIS run is working from rather
    // than of the list the question was asked about. Between the two the user
    // may have added a slide, deleted one, or moved them: the list is 1-based
    // positions, and a position that has moved names a different slide. The
    // read is already paid for here, and `slidesHolding` is the same sweep
    // "Used in this deck" uses.
    wanted = (await slidesHolding(await Pkg.open(deck.base64), id)).map((i) => i + 1);
    for (const slide of wanted) {
      // Counting from one in the state, from zero in the engine and the host.
      const at = slide - 1;
      const before = await slideCount();
      const targetId = await slideIdAt(at);
      if (targetId === undefined) break;
      const report = await removeElement({ deck: deck.base64, slide: at, element: id });
      // The raise is not consulted in either half of the cycle: the count on
      // the line after each call already is, and it is the evidence. Breaking
      // on the raise ahead of it stopped a cycle whose work had LANDED — the
      // rebuilt slide in the deck, the original never taken away, `done` still
      // zero, and the user told "The rest are as they were" over a deck now
      // one slide longer with the element on both. A raise that really did
      // nothing still stops the run, because the count then does not move.
      await insertPackage(report.base64, targetId);
      if ((await countReaching(before + 1)) !== before + 1) break;
      await removeSlideAt(at);
      if ((await countReaching(before)) !== before) {
        // The insert landed and the delete did not, so this slide's ORIGINAL is
        // still there with the element on it and an element-free copy sits
        // beside it. The other break above leaves the deck untouched; this one
        // does not, and the two cannot share a sentence.
        stranded = true;
        break;
      }
      done += 1;
    }
  } catch {
    // Whatever went wrong, `done` is the number of slides that were seen
    // through to the end, and the outcome below reports it.
  }

  const outcome = removalOutcome(element.name, done, wanted.length, stranded);
  // The armed Undo goes with it, whether or not a single slide was changed.
  //
  // `undoable` holds `before`: the WHOLE deck as it was before an earlier
  // insert. Undoing after a removal replays `undoPlan` against a deck that has
  // moved on, putting back a slide from bytes that predate the removal — so
  // the element the user just took off the deck comes back on one slide, and
  // every count check agrees, because each removal cycle is insert-then-remove
  // and is net zero on the slide count. The pane then printed "Undone."
  //
  // Disarmed on the way out rather than in the catch above: the entry stops
  // describing the deck at the first cycle that lands, and `done` is not known
  // until here. The same rule as the insert's `asked` flag — an undo entry that
  // no longer describes the deck is not an undo.
  undoable = undefined;
  state = {
    ...state,
    busy: false,
    busyWith: undefined,
    removing: undefined,
    outcome,
    undo: 0,
    // What the deck holds has changed under the pane, and the snapshot of the
    // current slide with it. Both are dropped rather than guessed: the next
    // "See what this deck already uses" is what puts them back.
    used: undefined,
    onSlide: undefined,
  };
  delete state.notice;
  draw();
  announce(outcome.detail);
}

/**
 * The right-click menu on a tile (`docs/DESIGN.md` section 6).
 *
 * Opened by the `contextmenu` event, which is the right mouse button AND the
 * keyboard's own menu key or Shift+F10 — so this is not a mouse-only feature by
 * accident. `preventDefault` only when a menu of ours actually opens: a user who
 * right-clicks the search box should still get the browser's own menu, with
 * paste in it.
 */
function onContextMenu(event: MouseEvent): void {
  const found = actionOf(event.target);
  const id = found?.el.dataset["id"];
  const element = elementOf(state.library, id);
  if (
    !found ||
    found.action !== "tile" ||
    !element ||
    !offersOtherTarget(element) ||
    state.busy === true ||
    // While a removal question is open, `render` will not draw a tile menu
    // (`state.removing === undefined` is one of its terms) — so opening one
    // here set a menu nothing would draw. The gesture did nothing at all, and
    // the menu arrived later out of nowhere: the Escape that answers the
    // question redraws with `menuFor` still set, and the menu opens on a tile
    // the user right-clicked long before. Refusing here is also the honest
    // answer to `docs/DESIGN.md` section 6 — no `preventDefault`, so the
    // browser's own menu stands rather than a gesture being swallowed by a
    // menu of ours that never appears.
    state.removing !== undefined
  ) {
    // A part ignores the insert target, so there is nothing to offer on one.
    if (state.menuFor !== undefined) set({ menuFor: undefined });
    return;
  }
  event.preventDefault();
  menuOwner = focusKey(found.el);
  set({ menuFor: tileKey(found.el.dataset["where"] ?? "", element.id) });
}

/**
 * Long-press is the same menu, for touch (section 6).
 *
 * Pointer events rather than touch events: one code path for a finger and a
 * pen, and `pointerType` is what tells them from a mouse — a mouse already has
 * the right button and a press-and-hold on one would be a surprise. The press
 * is abandoned as soon as the finger moves, because a hold that travels is a
 * scroll, and a menu that opened mid-scroll would be under the finger when it
 * lifted.
 */
const LONG_PRESS = 500;
let pressing: ReturnType<typeof setTimeout> | undefined;

/**
 * Set when a long press has OPENED the menu, and consumed by the click the same
 * gesture then produces.
 *
 * A finger lifting raises `pointerup` and then a `click`, and nothing cancelled
 * it: `onClick` closed the menu the press had just opened and fell straight
 * into `case "tile": void insert(id)` — inserting onto the slide the user is
 * on, which is the very target the menu exists to override. The menu flashed
 * up at 500 ms and was gone on the lift with an element in the deck.
 *
 * The mouse path never had this, because `contextmenu` fires with no left-button
 * click after it and `onContextMenu` calls `preventDefault`.
 *
 * A flag consumed once, the way `restoringFocus` gates `onFocus`, rather than
 * `preventDefault` on the pointer event: cancelling `pointerup` does not
 * reliably suppress the click on every engine, and this is about one click
 * rather than about the gesture.
 */
let pressOpened = false;

function cancelPress(): void {
  if (pressing !== undefined) clearTimeout(pressing);
  pressing = undefined;
}

function onPointerDown(event: PointerEvent): void {
  cancelPress();
  // A fresh gesture, so nothing is owed to the previous one. Cleared HERE
  // rather than on a timer: a long press whose click never arrives — which is
  // what Windows touch does, raising `contextmenu` instead — must not leave the
  // flag set to swallow the next tap.
  pressOpened = false;
  if (event.pointerType === "mouse") return;
  const found = actionOf(event.target);
  const element = elementOf(state.library, found?.el.dataset["id"]);
  if (!found || found.action !== "tile" || !element || !offersOtherTarget(element)) return;
  const key = tileKey(found.el.dataset["where"] ?? "", element.id);
  const owner = focusKey(found.el);
  pressing = setTimeout(() => {
    pressing = undefined;
    if (state.busy !== true) {
      pressOpened = true;
      menuOwner = owner;
      set({ menuFor: key });
    }
  }, LONG_PRESS);
}

function onClick(event: MouseEvent): void {
  const found = actionOf(event.target);
  // The click a long press produces when the finger lifts. It is the tail of
  // the gesture that opened the menu, not a new one, so it is swallowed whole:
  // without this it closed the menu and inserted with the default target.
  if (pressOpened) {
    pressOpened = false;
    event.preventDefault();
    return;
  }
  // Any click that is not ON the menu closes it, which is what every other
  // menu on every other platform does.
  if (state.menuFor !== undefined && found?.action !== "other-target") set({ menuFor: undefined });
  if (!found) return;
  const { action, el } = found;
  const id = el.dataset["id"];
  const value = el.dataset["value"];
  switch (action) {
    case "tile":
    case "step":
      if (id) void insert(id);
      break;
    case "insert":
      if (state.chosen) void insert(state.chosen);
      break;
    case "undo":
      void undo();
      break;
    case "retry":
      void load();
      break;
    case "star":
      if (id) {
        // `noQuestion` for the same reason the search and the categories carry
        // it: un-starring empties the Favourites section, and a question open
        // on a tile THERE goes with it — leaving the Remove button suppressed
        // on every other tile and nothing on screen saying why.
        set({ favourites: toggle(state.favourites, id), ...noQuestion });
        keep();
      }
      break;
    case "gear":
      gearOwner = state.gear === true ? undefined : focusKey(el);
      set({ gear: state.gear !== true });
      break;
    // The chevron `docs/DESIGN.md` section 4 has always described. Until
    // 2026-09-23 nothing drew it, and the tag line's one-row clip was lifted by
    // `state.gear` instead — so the line unfolded when the user opened the
    // OPTIONS panel and could not be opened on purpose at all.
    case "tags-open":
      set({ tagsOpen: state.tagsOpen !== true });
      break;
    case "target":
      if (value === "onto" || value === "new") {
        set({ settings: { ...state.settings, target: value } });
        keep();
      }
      break;
    case "group":
      set({ settings: { ...state.settings, group: value === "group" } });
      keep();
      break;
    case "colours":
      if (value === "deck" || value === "library") {
        set({ settings: { ...state.settings, colours: value } });
        keep();
      }
      break;
    // Section 7's two links. Both open the site the PANE was served from, in
    // the user's browser, and neither navigates the pane.
    case "report":
      // The build stamp, the host and the platform go with it — nothing else
      // can, because `reportUrl` takes nothing else (`src/host/links.ts`).
      leave((site) => reportUrl(site, { build: buildStamp(), ...hostStamp() }), "the support page");
      break;
    case "catalogue":
      leave(catalogueUrl, "the catalogue page");
      break;
    // Section 4: every category at once, beside the count. Remembered like
    // any other open category, because that is the same field.
    case "open-all":
      if (state.library) {
        set({ open: state.library.categories.map((category) => category.key) });
        keep();
      }
      break;
    // Section 6: the whole-slide element that landed on a busy slide, moved off
    // it. Undo then insert again, which is why it lives next to the two.
    case "move":
      void moveToNewSlide();
      break;
    case "used":
      void readUsed();
      break;
    // Section 4: a slide number in "Used in this deck" goes to that slide.
    case "jump":
      if (value !== undefined) void jumpTo(Number(value));
      break;
    // Sections 4 and 6: a part already in the deck, off every slide it is on —
    // asked first, because the pane cannot put it back.
    case "remove":
      if (id) {
        const element = elementOf(state.library, id);
        const slides = element ? removableFrom(element, state) : [];
        // The question is asked ON the tile that asked it: an element is drawn
        // in Favourites, in Recent and in its category, and a question keyed by
        // id alone appears on all three.
        const where = el.dataset["where"] ?? "";
        if (slides.length > 0 && element) {
          // Focus goes INTO the question, and the question is announced.
          // Neither happened: the redraw took the Remove button off every tile,
          // so the restore had nothing to find and focus fell to `<body>` —
          // leaving a keyboard user to Tab from the top of the document to
          // reach a confirmation they had opened one keystroke earlier, and a
          // screen-reader user told nothing at all. The question is the notice
          // rather than a summary of it, because it names the element and the
          // slides and says the pane cannot undo it.
          focusAfterDraw = `[data-action="remove-ask"][data-id="${CSS.escape(id)}"][data-where="${CSS.escape(where)}"]`;
          set({
            removing: { id, slides, done: 0, where },
            menuFor: undefined,
            notice: removeQuestion(element, slides),
          });
        }
      }
      break;
    case "remove-cancel": {
      // Same hand-off as the Escape rung. "Keep them" carries no `data-id` of
      // its own, so its `focusKey` is the bare `[data-action="remove-cancel"]`
      // — unfindable after the redraw that removes it, which put the focus on
      // `<body>` for a user who had just declined a destructive action.
      const asking = state.removing;
      if (asking) {
        focusAfterDraw =
          `[data-action="remove"][data-id="${CSS.escape(asking.id)}"]` + `[data-where="${CSS.escape(asking.where)}"]`;
      }
      set({ removing: undefined });
      break;
    }
    case "remove-go":
      if (id) void removeEverywhere(id);
      break;
    // Section 6: the other insert target, for this one insert, without touching
    // the setting.
    case "other-target":
      if (id) void insert(id, otherTarget(state.settings));
      break;
    case "tag":
      if (value) {
        set({ tags: toggle(state.tags, value), ...noQuestion });
        keep();
      }
      break;
    case "category":
      if (el.dataset["key"]) {
        set({ open: toggle(state.open, el.dataset["key"]), ...noQuestion });
        keep();
      }
      break;
    // Section 4: dismissed once, and remembered per machine like the
    // favourites are.
    case "coached":
      set({ coached: true });
      keep();
      break;
    case "clear":
      set({ query: "", tags: [], category: undefined, ...noQuestion });
      // Kept, like every keystroke that FILLED the box. Without it the deck's
      // bucket holds the old search and the next open restores one the user
      // explicitly got rid of — `docs/DESIGN.md` section 4 asks for the search
      // back, not for a cleared search back.
      keep();
      break;
    // Section 8's chips narrow a search to one category; picking the one
    // already picked widens it again.
    case "category-chip":
      if (el.dataset["key"]) {
        const key = el.dataset["key"];
        set({ category: state.category === key ? undefined : key, previewing: undefined, ...noQuestion });
      }
      break;
    // A "Did you mean" suggestion is a name the library really has, so putting
    // it in the search box is a search that will find something.
    case "guess":
      if (value) set({ query: value, previewing: undefined });
      break;
    default:
      break;
  }
}

/**
 * The question a tile is asking, dropped.
 *
 * Spread into every change that can take a tile OFF the screen, because the
 * question is drawn on the tile and nowhere else — and while one is open the
 * Remove button is suppressed on every tile. So a question whose tile was
 * filtered away left no question on screen, no Remove button anywhere either,
 * and only Escape to get out of a state nothing on screen was describing. A
 * question about a tile does not outlive the tile.
 */
const noQuestion = { removing: undefined } as const;

function onInput(event: Event): void {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.dataset["action"] !== "search") return;
  // A category chip narrows a SEARCH, and section 8 draws the chips only while
  // there is one. So a category left picked over a box the user backspaced to
  // empty goes on filtering the whole library down to that one category, with
  // nothing on screen saying so and no chip left to unpick it. Clear and Escape
  // already drop it; emptying the box by hand is the same gesture typed out.
  const query = target.value;
  set(query.trim() === "" ? { query, category: undefined, ...noQuestion } : { query, ...noQuestion });
  // Per keystroke, and deliberately not debounced: `keep` is one small
  // synchronous `setItem`, and a debounce would mean the pane forgetting
  // whatever was typed in the last moment before it was closed — which is
  // exactly the search a user is most likely to want back.
  keep();
}

/**
 * The keyboard, as `docs/DESIGN.md` section 9 specifies it.
 *
 * `/` focuses search, Esc clears it, arrows move between tiles, Enter inserts
 * the one the focus is on. Nothing here fights the browser: a tile is a real
 * button, so Enter and Space already work on it, and the arrows only move
 * focus.
 */
function onKey(event: KeyboardEvent): void {
  const inSearch = event.target instanceof HTMLInputElement;
  if (event.key === "/" && !inSearch) {
    const search = root().querySelector<HTMLInputElement>('[data-action="search"]');
    if (search) {
      event.preventDefault();
      search.focus();
    }
    return;
  }
  if (event.key === "Escape") {
    // The ladder is `escapeCloses` in `steps.ts`; this only does what it says.
    // `preview` is not a `set` like the others — it cancels a pending timer as
    // well — which is why the rule answers a name rather than a state patch.
    switch (escapeCloses(state)) {
      case "removing": {
        // Back to the Remove button that asked. It is not on screen while the
        // question is up — `render` suppresses it on every tile — so this is a
        // hand-off to the redraw rather than something `draw` could restore.
        const asking = state.removing;
        if (asking) {
          focusAfterDraw =
            `[data-action="remove"][data-id="${CSS.escape(asking.id)}"]` + `[data-where="${CSS.escape(asking.where)}"]`;
        }
        set({ removing: undefined });
        break;
      }
      case "menu":
        focusAfterDraw = menuOwner;
        menuOwner = undefined;
        set({ menuFor: undefined });
        break;
      case "preview":
        closePreview();
        break;
      case "gear":
        focusAfterDraw = gearOwner;
        gearOwner = undefined;
        set({ gear: false });
        break;
      case "search":
        // The CATEGORY goes too, and is kept, which is what the `clear` action
        // a few lines up already does. `steps.ts` states the rule — the picked
        // category is not carried across a cleared search — and this rung is
        // the half a user actually presses.
        //
        // What it left behind was invisible: `render` draws the chips only
        // while there IS a search, so a category still picked after Escape
        // narrowed the library to one section with no control on screen to
        // lift it and nothing saying why.
        set({ query: "", tags: [], category: undefined });
        keep();
        break;
    }
    return;
  }
  // LEFT and RIGHT belong to the search box while the caret is in it. The
  // branch below ran whatever the event target was, and `arrowTo` answers a
  // tile for an `at` of -1 — which is what the focus is when it is in the
  // input — so ArrowLeft was `preventDefault`ed and the focus thrown onto tile
  // 0. A user correcting a typo could not move the caret at all, the pane
  // redrew with a preview card open, and the next Enter activated the tile,
  // which is a real button: it inserted.
  //
  // DOWN is left alone deliberately. `arrowTo`'s own docstring justifies it —
  // "a user pressing Down from the search box" — and it is how the keyboard
  // reaches the tiles at all. Up with it, for symmetry: neither does anything
  // in a one-line input that Home and End do not.
  if (inSearch && (event.key === "ArrowLeft" || event.key === "ArrowRight")) return;
  const tiles = [...root().querySelectorAll<HTMLElement>('[data-action="tile"]')];
  const at = tiles.indexOf(document.activeElement as HTMLElement);
  // FROM A TILE, or out of the search box, and nowhere else. `arrowTo` clamps
  // an `at` of -1 to 0, so every arrow pressed on any other control — the
  // gear, a category heading, a tag chip, the size stepper, the star, a jump
  // button, the primary button, the menu item — was `preventDefault`ed and
  // threw the focus to the first tile at the top of the list. Two costs, and
  // the second is the one a mouse user feels: the focus surprise, and the
  // cancelled key, because in a 400 px pane ArrowDown is how you scroll and it
  // did nothing but jump to tile 0. The stepper is the sharpest case — a row
  // of numbers where left and right are the obvious gesture.
  //
  // The search box keeps its exception, which `arrowTo`'s own docstring
  // justifies: Down and Up out of it are how the keyboard reaches the tiles.
  if (at < 0 && !inSearch) return;
  const to = arrowTo(event.key, at, tiles.length);
  const next = to === undefined ? undefined : tiles[to];
  if (next) {
    event.preventDefault();
    // `focus` raises `focusin`, and `onFocus` marks the tile chosen — so the
    // `set` that used to follow this line was a second redraw for a state that
    // had already been set by the first.
    next.focus();
  }
}

/**
 * The preview card's delay: a third of a second of hover or focus
 * (`docs/DESIGN.md` section 4).
 *
 * The delay is the whole point. Without it, dragging the pointer across a grid
 * of tiles opens and closes a card per tile, which is a strobe rather than a
 * preview.
 */
const PREVIEW_DELAY = 333;
let previewTimer: ReturnType<typeof setTimeout> | undefined;

function cancelPreview(): void {
  if (previewTimer !== undefined) clearTimeout(previewTimer);
  previewTimer = undefined;
}

/** Open the card for `id` after the delay; any earlier pending open is dropped. */
function previewAfterDelay(id: string): void {
  cancelPreview();
  if (state.previewing === id) return;
  previewTimer = setTimeout(() => {
    previewTimer = undefined;
    set({ previewing: id });
  }, PREVIEW_DELAY);
}

/** Shut the card, and stop one that was about to open. */
function closePreview(): void {
  cancelPreview();
  if (state.previewing !== undefined) set({ previewing: undefined });
}

function onOver(event: MouseEvent): void {
  const found = actionOf(event.target);
  const id = found?.action === "tile" ? found.el.dataset["id"] : undefined;
  if (id) previewAfterDelay(id);
  else closePreview();
}

function onFocus(event: FocusEvent): void {
  // The focus `draw` just put back is not the user arriving anywhere.
  if (restoringFocus) return;
  const found = actionOf(event.target);
  if (found?.action === "tile" && found.el.dataset["id"]) {
    const id = found.el.dataset["id"];
    // Only when it CHANGED. `set` redraws, `draw` puts the focus back where it
    // was, and putting it back raises `focusin` again — so setting
    // unconditionally here is a redraw calling itself for as long as the stack
    // allows. It is also what `render`'s own comment asks for a line away: a
    // redraw the user did not ask for is a redraw that can take the focus off
    // whatever they were on.
    if (state.chosen !== id) set({ chosen: id });
    // Focus previews too, so the card is not a thing only a mouse can reach.
    previewAfterDelay(id);
    return;
  }
  closePreview();
}

/**
 * Follow PowerPoint's theme, not the browser's.
 *
 * The read is `themeBackground` and the decision is `paneTheme`; this is the
 * one line neither of them can be, which is stamping the answer on the
 * document. An unreadable or absent colour leaves the attribute UNSET rather
 * than guessing — `taskpane.css` carries a `prefers-color-scheme` fallback for
 * exactly that case, and a guess would override it with a worse answer.
 */
function applyTheme(): void {
  const theme = paneTheme(themeBackground());
  if (theme) document.documentElement.setAttribute("data-theme", theme);
}

/** The commit this pane was built from, or undefined outside a build. */
function buildStamp(): string | undefined {
  return typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : undefined;
}

/**
 * The build this pane was served from, on the document, before anything is run.
 *
 * PowerPoint caches the pane's HTML for about ten minutes, so opening it too
 * soon after a deploy tests code the host never fetched, and the result reads
 * as a clean run of the wrong build. The sibling projects record whole rounds
 * lost to it, and this project used it twice on 2026-09-14 to know the site had
 * caught up with `main` before trusting a round.
 *
 * It used to be seven characters PAINTED in the header. The owner asked for
 * them out of the AppSource screenshot on 2026-09-15, and a stamp that is in
 * the picture cannot be left out of it honestly — `docs/LISTING.md` forbids
 * retouching. So it moved rather than went: `data-build` on the root element is
 * invisible to a user and to a screenshot, and is still there for anyone
 * reading the DOM — devtools, a support request, or a driver over CDP, which is
 * how this project reads it. The diagnostic survives; only the paint is gone.
 *
 * "Report a problem" carries the same value into its prefilled URL and is
 * unaffected, so a user who cannot open devtools still has a way to send it.
 */
function stampBuild(): void {
  const build = buildStamp();
  if (!build) return;
  document.documentElement.setAttribute("data-build", build);
}

void Office.onReady(() => {
  applyTheme();
  // Before the floor check: a host that cannot run the add-in is exactly the
  // case where somebody needs to say which build refused them.
  stampBuild();
  const check = hostReady();
  if (!check.ok) {
    // Said out loud rather than swallowed: a pane that renders a dead UI on an
    // unsupported host gives the user nothing to report.
    const node = root();
    node.textContent = "";
    const p = document.createElement("p");
    p.className = "blocked";
    p.textContent = check.detail;
    node.append(p);
    return;
  }
  // Whether a slide number can be a link, decided once: the call behind it is
  // PowerPointApi 1.5, and `src/host/jump.ts` says why it may be made at all.
  state = { ...state, ...remembered(), canJump: hostSupports("1.5") };
  document.addEventListener("click", onClick);
  document.addEventListener("contextmenu", onContextMenu);
  // Long-press for touch, and every way a press can end without becoming one.
  document.addEventListener("pointerdown", onPointerDown);
  document.addEventListener("pointerup", cancelPress);
  document.addEventListener("pointermove", cancelPress);
  document.addEventListener("pointercancel", cancelPress);
  document.addEventListener("input", onInput);
  // On `window`, because the pane scrolls the document rather than a box of its
  // own: nothing in `taskpane.css` sets `overflow` on a container.
  window.addEventListener("scroll", onScroll, { passive: true });
  document.addEventListener("keydown", onKey);
  document.addEventListener("focusin", onFocus);
  document.addEventListener("mouseover", onOver);
  // Leaving the pane entirely, which no mouseover over a tile will report.
  document.addEventListener("mouseleave", closePreview);
  draw();
  void load();
});
