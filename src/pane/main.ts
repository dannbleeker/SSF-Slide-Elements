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
import { slideSize } from "../core/pptx/layout.js";
import { Pkg } from "../core/pptx/pkg.js";
import { onlySlide, splice } from "../core/splice/splice.js";
import { coalescing } from "../host/coalesce.js";
import { INSERTING, announcement, mayRemove, outcomeOf, undoPlan } from "../host/insert.js";
import { readable } from "../host/errors.js";
import { catalogueUrl, reportUrl, siteFrom } from "../host/links.js";
import { BUDGET } from "../host/timeout.js";
import {
  currentSlide,
  hostStamp,
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
  remember,
  stepFor,
  toggle,
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

/** The deck's slide size, read from the file, because no API answers it. */
async function deckShape(): Promise<{ width: number; height: number } | undefined> {
  try {
    const deck = await readDeck();
    return await slideSize(await Pkg.open(deck.base64));
  } catch {
    return undefined;
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
  set({ slide: current == null ? undefined : current.index + 1 });
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

// ---------------------------------------------------------------------------
// Inserting.
// ---------------------------------------------------------------------------

async function insert(id: string): Promise<void> {
  const library = state.library;
  const element = elementOf(library, id);
  if (!library || !element || !store || !index || state.busy === true) return;

  set({ busy: true, chosen: id, notice: INSERTING, outcome: undefined });
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
      options: state.settings,
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
    if (state.settings.target === "onto" && mayRemove({ before, inserted })) {
      // The rebuilt slide landed AFTER the original, so the original is still
      // at its own index. Positional, never by id: a slide next to one the run
      // has just added is exactly where an id read is not to be trusted.
      const failure = await removeSlideAt(at);
      removed = failure === undefined ? await countReaching(before) : inserted;
    }

    const outcome = outcomeOf({
      target: state.settings.target,
      slide: at + 1,
      before,
      inserted,
      ...(removed === undefined ? {} : { removed }),
      ...(error === undefined ? {} : { error }),
    });
    undoable = outcome.ok
      ? { target: state.settings.target, index: at, before: deck.base64, name: element.name }
      : undefined;
    state = {
      ...state,
      busy: false,
      recent: outcome.ok ? remember(state.recent, element.id, RECENT_DEPTH) : state.recent,
      undo: outcome.ok ? 1 : 0,
      outcome: { ...outcome, name: element.name },
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

function onClick(event: MouseEvent): void {
  const found = actionOf(event.target);
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
    if (state.previewing !== undefined) closePreview();
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
  state = { ...state, ...remembered() };
  document.addEventListener("click", onClick);
  document.addEventListener("input", onInput);
  document.addEventListener("keydown", onKey);
  document.addEventListener("focusin", onFocus);
  document.addEventListener("mouseover", onOver);
  // Leaving the pane entirely, which no mouseover over a tile will report.
  document.addEventListener("mouseleave", closePreview);
  draw();
  void load();
});
