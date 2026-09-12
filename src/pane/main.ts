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
import type { Box } from "../core/catalogue/types.js";
import { slideSize } from "../core/pptx/layout.js";
import { Pkg } from "../core/pptx/pkg.js";
import { usedInDeck } from "../core/pptx/tags.js";
import { removeElement, slidesHolding } from "../core/splice/remove.js";
import { onlySlide, splice } from "../core/splice/splice.js";
import { coalescing } from "../host/coalesce.js";
import { INSERTING, announcement, mayRemove, outcomeOf, undoPlan } from "../host/insert.js";
import { readable } from "../host/errors.js";
import { jumpOutcome } from "../host/jump.js";
import { catalogueUrl, reportUrl, siteFrom } from "../host/links.js";
import { BUDGET } from "../host/timeout.js";
import {
  currentSlide,
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
  countReaching,
  slideCount,
  slideIdAt,
} from "../office/powerpoint.js";
import { Store, carriedTypes, libraryFor, loadIndex, themeColours, type Index } from "./catalogue.js";
import { render } from "./render.js";
import {
  DEFAULT_SETTINGS,
  EMPTY,
  RECENT_DEPTH,
  elementOf,
  fractionOf,
  offersOtherTarget,
  otherTarget,
  remember,
  removableFrom,
  removalOutcome,
  stepFor,
  tileKey,
  toggle,
  withInsert,
  withLanded,
  withoutInsert,
  type Library,
  type PaneState,
} from "./steps.js";

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
  /** What the pane knew the destination slide held BEFORE the insert. */
  onSlide: PaneState["onSlide"];
}
let undoable: Undoable | undefined;

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
function draw(): void {
  const active = document.activeElement;
  const wasSearch = active instanceof HTMLInputElement && active.dataset["action"] === "search";
  const caret = wasSearch ? active.selectionStart : null;

  render(root(), state, stepFor(state));
  liveRegion();
  announce(state.notice ?? "");

  if (wasSearch) {
    const search = root().querySelector<HTMLInputElement>('[data-action="search"]');
    if (search) {
      search.focus();
      if (caret !== null) search.setSelectionRange(caret, caret);
    }
  }
}

function set(changes: Partial<PaneState>): void {
  state = { ...state, ...changes };
  draw();
}

// ---------------------------------------------------------------------------
// Settings, kept per machine.
// ---------------------------------------------------------------------------

const KEY = "ssf-slide-elements";

/**
 * The pane reopens where you left it (`docs/DESIGN.md` section 4).
 *
 * Wrapped, because storage is not always there: a WebView with site data
 * blocked throws on the accessor itself rather than answering empty, and a pane
 * that will not open because it could not remember a search box is worse than
 * one that forgets.
 */
function remembered(): Partial<PaneState> {
  try {
    const raw = window.localStorage.getItem(KEY);
    if (!raw) return {};
    const held = JSON.parse(raw) as Partial<PaneState>;
    return {
      settings: { ...DEFAULT_SETTINGS, ...(held.settings ?? {}) },
      favourites: held.favourites ?? [],
      coached: held.coached === true,
      recent: held.recent ?? [],
      open: held.open ?? [],
    };
  } catch {
    return {};
  }
}

function keep(): void {
  try {
    window.localStorage.setItem(
      KEY,
      JSON.stringify({
        settings: state.settings,
        favourites: state.favourites,
        coached: state.coached === true,
        recent: state.recent,
        open: state.open,
      }),
    );
  } catch {
    // A pane that cannot remember is still a pane that works.
  }
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
  set({ library: provisional });
  store = new Store(provisional.size);

  const shape = await deckShape();
  if (!shape || !index) return;
  const library: Library = libraryFor(index, shape.width, shape.height);
  if (library.size !== provisional.size || library.borrowed !== provisional.borrowed) {
    store = new Store(library.size);
    set({ library });
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
 * Whether a selection read is already in flight.
 *
 * PowerPoint fires the selection event for every shape a user touches, and each
 * read is a `PowerPoint.run`. One at a time, and none at all while an insert
 * is running: a read that overlaps the insert tells the user nothing they need
 * and costs the host a context it is already using.
 */
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
  const target = once ?? state.settings.target;

  set({ busy: true, chosen: id, notice: INSERTING, outcome: undefined, menuFor: undefined });
  try {
    const markup = await store.markup(element);
    const deck = await readDeck();
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
        ...(element.category.key.toLowerCase().includes("mark") ? { wraps: true } : {}),
        markup: { xml: markup.xml, rels: markup.rels },
      },
      options: { ...state.settings, target },
      catalogue: {
        version: library.version,
        carried: carriedTypes(index, library.size),
        theme: themeColours(index, library.size),
      },
      store: (path) => (store as Store).part(path),
      ...(selection ? { selection } : {}),
    });

    const targetId = current?.id;
    if (targetId === undefined) {
      set({
        busy: false,
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

    const error = await insertPackage(report.base64, targetId);
    // Asked again until it agrees, because the count lags the insert: see
    // `countReaching`. One read here would report a landed insert as a no-op.
    const inserted = await countReaching(before + 1);
    let removed: number | undefined;
    if (target === "onto" && mayRemove({ before, inserted })) {
      // The rebuilt slide landed AFTER the original, so the original is still
      // at its own index. Positional, never by id: a slide next to one the run
      // has just added is exactly where an id read is not to be trusted.
      const failure = await removeSlideAt(at);
      removed = failure === undefined ? await countReaching(before) : inserted;
    }

    const outcome = outcomeOf({
      target,
      slide: at + 1,
      before,
      inserted,
      ...(removed === undefined ? {} : { removed }),
      ...(error === undefined ? {} : { error }),
    });
    // Where the element ended up, counting from one. "Onto this slide" rebuilt
    // the slide the user was on and took the original away, so it is that
    // slide; "as a new slide" put one after it. The same arithmetic `undoPlan`
    // does, from the other end.
    const landedOn = target === "new" ? at + 2 : at + 1;
    undoable = outcome.ok
      ? {
          target,
          index: at,
          before: deck.base64,
          name: element.name,
          id: element.id,
          landedOn,
          onSlide: state.onSlide,
        }
      : undefined;
    state = {
      ...state,
      busy: false,
      recent: outcome.ok ? remember(state.recent, element.id, RECENT_DEPTH) : state.recent,
      undo: outcome.ok ? 1 : 0,
      outcome: { ...outcome, name: element.name },
      // "Used in this deck" is updated rather than re-read: the pane knows
      // exactly what it just put where, and the read is the expensive half of
      // this feature. Untouched when nothing has been read — an insert is not a
      // reason to start claiming the deck has been looked at.
      used: outcome.ok ? withInsert(state.used, element.id, landedOn) : state.used,
      // The card's grey boxes keep up the same way: the splice says where the
      // element landed, in EMU **on the user's slide**, so it is the USER's
      // slide size that turns it into a fraction. The library's size is a
      // different number whenever a deck borrowed the nearest library, and
      // dividing by that one draws the right rectangle in the wrong place on
      // exactly the decks nobody tests on.
      onSlide: outcome.ok
        ? withLanded(state.onSlide, landedOn, fractionOf(report.landed, state.deck), target === "new")
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
    state = {
      ...state,
      busy: false,
      outcome: { ok: false, byHand: false, name: element.name, detail: `The insert was refused: ${readable(e)}` },
    };
    delete state.notice;
    draw();
  }
}

/**
 * Put the deck back the way it was.
 *
 * Count-checked at every step, and positional throughout. Nothing here reads an
 * id: `CLAUDE.md` records that a slide the run just added does not resolve by
 * one on the web, and undo is working right next to one.
 */
async function undo(): Promise<void> {
  const entry = undoable;
  if (!entry || state.busy === true) return;
  set({ busy: true, notice: "Undoing…" });
  const plan = undoPlan(entry);
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
      const refused = await insertPackage(original.base64, targetId);
      const grown = await countReaching(plan.grownTo(before));
      if (grown !== plan.grownTo(before)) {
        throw new Error(
          refused ?? `the deck went ${before} → ${grown}, which is not what putting one slide back looks like`,
        );
      }
    }

    const refused = await removeSlideAt(plan.remove);
    const after = await countReaching(before - (plan.after === undefined ? 1 : 0));
    if (refused !== undefined || after !== before - (plan.after === undefined ? 1 : 0)) {
      throw new Error(refused ?? `the deck has ${after} slides, which is not what was expected`);
    }

    undoable = undefined;
    state = {
      ...state,
      busy: false,
      undo: 0,
      outcome: { ok: true, byHand: false, name: entry.name, detail: `Undone. The deck has ${after} slides.` },
      // The slide it was on is the user's own again, so whatever the insert put
      // there went with it — both in the list and in the card's grey boxes.
      used: withoutInsert(state.used, entry.id, entry.landedOn),
      onSlide: entry.onSlide,
    };
    delete state.notice;
    draw();
    announce(`${entry.name} taken back.`);
    followSelection();
  } catch (e) {
    state = {
      ...state,
      busy: false,
      outcome: { ok: false, byHand: true, name: entry.name, detail: `Undo did not work: ${readable(e)}` },
    };
    delete state.notice;
    draw();
  }
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
  try {
    const deck = await readDeck();
    const pkg = await Pkg.open(deck.base64);
    const used = await usedInDeck(pkg);
    // The same read answers what the current slide holds, so the card's grey
    // boxes come back into step with the deck for free.
    const size = await slideSize(pkg);
    const held = await heldBy(pkg, size);
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
 * The right-click menu on a tile (`docs/DESIGN.md` section 6).
 *
 * Opened by the `contextmenu` event, which is the right mouse button AND the
 * keyboard's own menu key or Shift+F10 — so this is not a mouse-only feature by
 * accident. `preventDefault` only when a menu of ours actually opens: a user who
 * right-clicks the search box should still get the browser's own menu, with
 * paste in it.
 */
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
async function removeEverywhere(id: string): Promise<void> {
  const plan = state.removing;
  const element = elementOf(state.library, id);
  if (!plan || !element || plan.id !== id || state.busy === true) return;
  set({ busy: true, notice: `Taking ${element.name} off ${plan.slides.length} slide(s)…`, outcome: undefined });

  let done = 0;
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
      const refused = await insertPackage(report.base64, targetId);
      if (refused !== undefined) break;
      if ((await countReaching(before + 1)) !== before + 1) break;
      const failed = await removeSlideAt(at);
      if (failed !== undefined) break;
      if ((await countReaching(before)) !== before) break;
      done += 1;
    }
  } catch {
    // Whatever went wrong, `done` is the number of slides that were seen
    // through to the end, and the outcome below reports it.
  }

  const outcome = removalOutcome(element.name, done, wanted.length);
  state = {
    ...state,
    busy: false,
    removing: undefined,
    outcome,
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

function onContextMenu(event: MouseEvent): void {
  const found = actionOf(event.target);
  const id = found?.el.dataset["id"];
  const element = elementOf(state.library, id);
  if (!found || found.action !== "tile" || !element || !offersOtherTarget(element) || state.busy === true) {
    // A part ignores the insert target, so there is nothing to offer on one.
    if (state.menuFor !== undefined) set({ menuFor: undefined });
    return;
  }
  event.preventDefault();
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

function cancelPress(): void {
  if (pressing !== undefined) clearTimeout(pressing);
  pressing = undefined;
}

function onPointerDown(event: PointerEvent): void {
  cancelPress();
  if (event.pointerType === "mouse") return;
  const found = actionOf(event.target);
  const element = elementOf(state.library, found?.el.dataset["id"]);
  if (!found || found.action !== "tile" || !element || !offersOtherTarget(element)) return;
  const key = tileKey(found.el.dataset["where"] ?? "", element.id);
  pressing = setTimeout(() => {
    pressing = undefined;
    if (state.busy !== true) set({ menuFor: key });
  }, LONG_PRESS);
}

function onClick(event: MouseEvent): void {
  const found = actionOf(event.target);
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
    case "again":
      if (state.recent[0]) void insert(state.recent[0]);
      break;
    case "undo":
      void undo();
      break;
    case "retry":
      void load();
      break;
    case "star":
      if (id) {
        set({ favourites: toggle(state.favourites, id) });
        keep();
      }
      break;
    case "gear":
      set({ gear: state.gear !== true });
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
        if (slides.length > 0) set({ removing: { id, slides, done: 0, where }, menuFor: undefined });
      }
      break;
    case "remove-cancel":
      set({ removing: undefined });
      break;
    case "remove-go":
      if (id) void removeEverywhere(id);
      break;
    // Section 6: the other insert target, for this one insert, without touching
    // the setting.
    case "other-target":
      if (id) void insert(id, otherTarget(state.settings));
      break;
    case "tag":
      if (value) set({ tags: toggle(state.tags, value) });
      break;
    case "category":
      if (el.dataset["key"]) {
        set({ open: toggle(state.open, el.dataset["key"]) });
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
      set({ query: "", tags: [], category: undefined });
      break;
    // Section 8's chips narrow a search to one category; picking the one
    // already picked widens it again.
    case "category-chip":
      if (el.dataset["key"]) {
        const key = el.dataset["key"];
        set({ category: state.category === key ? undefined : key, previewing: undefined });
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

function onInput(event: Event): void {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.dataset["action"] !== "search") return;
  set({ query: target.value });
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
    // The card first: it is the most recently opened thing and the one the user
    // is most likely to mean, and shutting it must not also clear their search.
    // The confirm first, then the tile menu, then the card, then the gear: back
    // out of what was opened last, and never clear a search on the way past
    // something else. A question the user escapes is a question answered "no".
    if (state.removing !== undefined) set({ removing: undefined });
    else if (state.menuFor !== undefined) set({ menuFor: undefined });
    else if (state.previewing !== undefined) closePreview();
    else if (state.gear === true) set({ gear: false });
    else if (state.query !== "" || state.tags.length > 0) set({ query: "", tags: [] });
    return;
  }
  if (event.key !== "ArrowRight" && event.key !== "ArrowLeft" && event.key !== "ArrowDown" && event.key !== "ArrowUp") {
    return;
  }
  const tiles = [...root().querySelectorAll<HTMLElement>('[data-action="tile"]')];
  if (tiles.length === 0) return;
  const at = tiles.findIndex((t) => t === document.activeElement);
  if (at < 0) {
    event.preventDefault();
    tiles[0]?.focus();
    return;
  }
  const step = event.key === "ArrowRight" || event.key === "ArrowDown" ? 1 : -1;
  const next = tiles[Math.min(tiles.length - 1, Math.max(0, at + step))];
  if (next) {
    event.preventDefault();
    next.focus();
    const id = next.dataset["id"];
    if (id) set({ chosen: id });
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
  const found = actionOf(event.target);
  if (found?.action === "tile" && found.el.dataset["id"]) {
    const id = found.el.dataset["id"];
    set({ chosen: id });
    // Focus previews too, so the card is not a thing only a mouse can reach.
    previewAfterDelay(id);
    return;
  }
  closePreview();
}

/**
 * Follow PowerPoint's theme, not the browser's.
 *
 * The pane lives inside PowerPoint, which can be dark while the OS is light, so
 * `prefers-color-scheme` is the wrong question. `officeTheme` answers the right
 * one. Outside a host it is undefined — which is the case every time this pane
 * is opened in a browser to look at it — and the stylesheet's media query
 * carries that fallback.
 *
 * Read ONCE, on ready. There is no theme-change event for a PowerPoint pane:
 * the typings put `OfficeThemeChanged` on Outlook's `Mailbox` and nowhere else,
 * so switching PowerPoint's theme mid-session needs the pane reopened.
 */
function applyTheme(): void {
  const body = Office.context?.officeTheme?.bodyBackgroundColor;
  if (!body) return;
  const hex = body.replace("#", "");
  const n = Number.parseInt(hex.length === 3 ? [...hex].map((c) => c + c).join("") : hex, 16);
  if (Number.isNaN(n)) return;
  const luminance = ((n >> 16) & 255) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114;
  document.documentElement.setAttribute("data-theme", luminance < 128 ? "dark" : "light");
}

/**
 * The build this pane was served from, in the header, before anything is run.
 *
 * PowerPoint caches the pane's HTML for about ten minutes, so opening it too
 * soon after a deploy tests code the host never fetched, and the result reads
 * as a clean run of the wrong build. The sibling projects record whole rounds
 * lost to it. Seven characters in the header is how the two are told apart.
 */
/** The commit this pane was built from, or undefined outside a build. */
function buildStamp(): string | undefined {
  return typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : undefined;
}

function showBuild(): void {
  const build = buildStamp() ?? "unknown";
  const header = document.querySelector("header");
  if (!header || build === "unknown") return;
  const span = document.createElement("span");
  span.className = "build";
  span.textContent = build;
  span.title = `SSF Slide Elements was built from commit ${build}`;
  header.append(span);
}

void Office.onReady(() => {
  applyTheme();
  // Before the floor check: a host that cannot run the add-in is exactly the
  // case where somebody needs to say which build refused them.
  showBuild();
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
  document.addEventListener("keydown", onKey);
  document.addEventListener("focusin", onFocus);
  document.addEventListener("mouseover", onOver);
  // Leaving the pane entirely, which no mouseover over a tile will report.
  document.addEventListener("mouseleave", closePreview);
  draw();
  void load();
});
