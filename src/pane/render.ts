/**
 * The pane's DOM, built from the decisions in `steps.ts`.
 *
 * Every string a user sees reaches the page through `textContent`; nothing here
 * assigns markup, so a name out of the catalogue or a reason out of the host
 * cannot become an element. `test/security.test.ts` holds that.
 *
 * The layout rules the SSF pane was approved on — one column, one orange tick,
 * one primary control drawn last — are what this file holds, and
 * `scripts/pane-shots.mjs` is what measures them. It imports this function
 * directly and hands it a state, so `render` must stay synchronous and must
 * decide nothing: everything it draws is either in the state or answered by
 * `steps.ts`.
 *
 * Every control carries a `data-action`, because that is what the shot audit
 * reads to check that a state shows what it claims to.
 */
import type { Box, Element } from "../core/catalogue/types.js";
import { previewUrl } from "./catalogue.js";
import { landingLine, occupiedFor, slideLine } from "./card.js";
import {
  categoryHits,
  didYouMean,
  elementOf,
  groups,
  isOpen,
  offersOpenAll,
  runOf,
  stepMatches,
  tagsOf,
  tileCount,
} from "./search.js";
import {
  blockedReason,
  borrowedLine,
  COACH,
  coaching,
  footerOf,
  offersOtherTarget,
  otherTargetLabel,
  primary,
  removableFrom,
  removeLabel,
  removeQuestion,
  settingsLine,
  STEP_TITLE,
  tileKey,
  type Library,
  type PaneState,
  type StepId,
} from "./steps.js";
import { slideParts, usedHeading, usedRows } from "./used.js";

const SVG_NS = "http://www.w3.org/2000/svg";

/** How many tags the line shows before the chevron is worth having. */
const TAGS_SHOWN = 12;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function button(action: string, className: string, text: string): HTMLButtonElement {
  const node = el("button", className, text);
  node.type = "button";
  node.dataset["action"] = action;
  return node;
}

/**
 * The little slide with the element's box on it.
 *
 * The FALLBACK now rather than the only thing a tile can show: `picture` below
 * puts the element's own photograph on the tile, cut from the deck's print, and
 * this draws when there is none. That happens in a tree where the previews have
 * not been built — they are generated at deploy, not committed — and it would
 * happen for a new element whose deck has been merged but not yet deployed.
 *
 * What it draws is what the catalogue KNOWS without a picture: the slide's
 * shape, and where on it the element sits and how much of it the element
 * covers. It is the same ghost the preview card uses for the landing, at tile
 * size.
 *
 * Built with `createElementNS`, not markup, for the same reason as everything
 * else here.
 */
function ghost(element: Element, aspect: number, occupied: Box[] = []): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, "svg");
  svg.setAttribute("viewBox", `0 0 ${Math.round(100 * aspect)} 100`);
  svg.setAttribute("class", "ghost");
  // Decorative: the name beside it is the accessible label, and a second
  // announcement of the same thing is noise in a screen reader.
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");

  const slide = document.createElementNS(SVG_NS, "rect");
  slide.setAttribute("x", "0.5");
  slide.setAttribute("y", "0.5");
  slide.setAttribute("width", String(Math.round(100 * aspect) - 1));
  slide.setAttribute("height", "99");
  slide.setAttribute("class", "ghost-slide");
  svg.appendChild(slide);

  // What the destination slide already holds, drawn UNDER the element's own
  // frame (`docs/DESIGN.md` sections 1 and 4): grey, in z-order, so the user can
  // see whether the element is about to land on top of something. Empty on a
  // tile, and on a card whose snapshot is of another slide — `occupiedFor`
  // decides that, not this.
  for (const held of occupied) {
    const grey = document.createElementNS(SVG_NS, "rect");
    grey.setAttribute("x", String((held.x * 100 * aspect).toFixed(1)));
    grey.setAttribute("y", String((held.y * 100).toFixed(1)));
    grey.setAttribute("width", String(Math.max(1, held.w * 100 * aspect).toFixed(1)));
    grey.setAttribute("height", String(Math.max(1, held.h * 100).toFixed(1)));
    grey.setAttribute("class", "ghost-held");
    svg.appendChild(grey);
  }

  const box = document.createElementNS(SVG_NS, "rect");
  box.setAttribute("x", String((element.box.x * 100 * aspect).toFixed(1)));
  box.setAttribute("y", String((element.box.y * 100).toFixed(1)));
  box.setAttribute("width", String(Math.max(2, element.box.w * 100 * aspect).toFixed(1)));
  box.setAttribute("height", String(Math.max(2, element.box.h * 100).toFixed(1)));
  box.setAttribute("class", "ghost-box");
  svg.appendChild(box);
  return svg;
}

/**
 * The element's own photograph, with the landing ghost behind it.
 *
 * PowerPoint's rendering of the element, cut out of the deck's print
 * (`docs/DESIGN.md` section 3). The picture is what a person recognises an
 * element BY, which a diagram of its box is not.
 *
 * The ghost is drawn first and the image laid over it, rather than swapping one
 * for the other on load. A tile that draws nothing until a PNG arrives is a
 * list of empty boxes on a slow connection, and an `onload` swap would reflow
 * every tile in the category as the images land. If the picture never arrives —
 * a tree where the previews have not been built, a deploy that has not caught
 * up with a new element — the ghost is simply what stays visible.
 */
function picture(element: Element, library: Library, occupied: Box[] = [], aspect?: number): HTMLElement {
  const frame = el("span", "tile-shot");
  frame.appendChild(ghost(element, aspect ?? library.width / library.height, occupied));

  const img = document.createElement("img");
  img.className = "tile-img";
  // Decorative: the name beside it is the accessible label, and the ghost
  // underneath already carries aria-hidden for the same reason.
  img.alt = "";
  // Attributes rather than properties: a category can hold 22 tiles and only a
  // few are on screen, so `lazy` is doing real work — and it is only testable
  // as an attribute, which is also how it reaches the parser.
  img.setAttribute("loading", "lazy");
  img.setAttribute("decoding", "async");
  img.src = previewUrl(library.size, element.id, library.version);
  img.addEventListener("error", () => img.remove(), { once: true });
  frame.appendChild(img);
  return frame;
}

/**
 * The preview card: the element at full width, its name, and where it lands.
 *
 * `docs/DESIGN.md` section 4 opens it after a third of a second of hover or
 * focus, pinned over the top of the list; at 512px and wider it docks beside
 * the list and hides no tiles. Both of those are the stylesheet's job — this
 * decides only what is IN it.
 *
 * The little slide in it shows where the element lands AND, in grey, what the
 * slide the user is on already holds — read out of the deck rather than through
 * the API, so it needs no host capability the insert does not already use. The
 * boxes are drawn only while the snapshot is of the slide the user is actually
 * on; `occupiedFor` in `steps.ts` is what decides that.
 *
 * Inert to the pointer. A card that opens under the cursor and then swallows
 * the click would make the tile it describes unpickable.
 */
function card(element: Element, library: Library, state: PaneState): HTMLElement {
  const box = el("aside", "card");
  box.dataset["id"] = element.id;
  // Announced by the tile it belongs to, not by itself: the tile already
  // carries "Insert <name>", and a live card would interrupt a reader mid-word.
  box.setAttribute("aria-hidden", "true");
  // The card's little slide is the USER's slide, so it is drawn in the user's
  // own shape when the pane has read it: the grey boxes are fractions of that
  // slide, and on a deck that borrowed the nearest library — A4, 16:10, custom
  // — the library's shape is a different rectangle. The element's own frame is
  // in LIBRARY fractions and is therefore approximate on such a deck, which is
  // what the borrowed line under the header already says out loud.
  const aspect = state.deck && state.deck.height > 0 ? state.deck.width / state.deck.height : undefined;
  box.appendChild(picture(element, library, occupiedFor(state), aspect));
  box.appendChild(el("strong", "card-name", element.name));
  box.appendChild(el("p", "card-landing", landingLine(element, state.settings)));
  return box;
}

/** One element's tile: where it lands, what it is called, and whether it is starred. */
function tile(state: PaneState, library: Library, element: Element, where: string): HTMLElement {
  const key = tileKey(where, element.id);
  const item = el("li", "tile");
  const pick = button("tile", "tile-pick", "");
  pick.dataset["id"] = element.id;
  // Which of the up-to-three tiles for this element this one is. The menu and
  // the removal question are anchored to a TILE, and the click handler reads
  // this back to know which one was asked.
  pick.dataset["where"] = where;
  pick.setAttribute("aria-label", `Insert ${element.name}`);
  if (state.chosen === element.id) pick.setAttribute("aria-current", "true");
  if (state.busy === true) pick.disabled = true;
  pick.appendChild(picture(element, library));
  pick.appendChild(el("span", "tile-name", element.name));
  if (state.busy === true && state.chosen === element.id) {
    pick.appendChild(el("span", "tile-busy", "Inserting…"));
  }
  item.appendChild(pick);

  const star = button("star", state.favourites.includes(element.id) ? "star on" : "star", "★");
  star.dataset["id"] = element.id;
  star.setAttribute(
    "aria-label",
    state.favourites.includes(element.id)
      ? `Remove ${element.name} from favourites`
      : `Add ${element.name} to favourites`,
  );
  star.setAttribute("aria-pressed", state.favourites.includes(element.id) ? "true" : "false");
  item.appendChild(star);

  // A sized element is one tile with a stepper naming what it counts:
  // "boxes 1 2 3 4 5 6" (`docs/DESIGN.md` section 2).
  const run = runOf(library, element);
  if (run.length > 1 && element.run) {
    const stepper = el("div", "stepper");
    stepper.appendChild(el("span", "stepper-noun", element.run.noun));
    for (const member of run) {
      // Section 8: a sized tile greys out the counts that do not match. Still
      // pickable — the size exists and somebody may want it — but the search
      // says which one it was looking for.
      const hit = stepMatches(member, state);
      const classes = [state.chosen === member.id ? "step on" : "step", hit ? "" : "off"].filter(Boolean).join(" ");
      const step = button("step", classes, String(member.run?.count ?? ""));
      step.dataset["id"] = member.id;
      step.setAttribute("aria-label", hit ? member.name : `${member.name}, not a match`);
      if (state.busy === true) step.disabled = true;
      stepper.appendChild(step);
    }
    item.appendChild(stepper);
  }

  // Section 4: a part already in the deck can be taken off every slide it is
  // on. Only once the deck has been READ — before that the pane does not know
  // what is in it, and `removableFrom` is what decides.
  const removable = removableFrom(element, state);
  if (removable.length > 0 && state.removing === undefined) {
    const take = button("remove", "tile-remove", removeLabel(removable));
    take.dataset["id"] = element.id;
    take.dataset["where"] = where;
    if (state.busy === true) take.disabled = true;
    item.appendChild(take);
  }

  // The confirm, in the same place the right-click menu goes: it is the only
  // thing here that takes something OUT of the user's deck, so it is asked
  // rather than done, and the question names the slides.
  if (state.removing !== undefined && tileKey(state.removing.where, state.removing.id) === key) {
    const ask = el("div", "tile-menu");
    ask.setAttribute("role", "group");
    ask.setAttribute("aria-label", removeLabel(state.removing.slides));
    ask.appendChild(el("p", "tile-ask", removeQuestion(element, state.removing.slides)));
    const go = button("remove-go", "tile-menu-item danger", "Remove");
    go.dataset["id"] = element.id;
    const no = button("remove-cancel", "tile-menu-item", "Keep them");
    if (state.busy === true) {
      go.disabled = true;
      no.disabled = true;
    }
    ask.appendChild(go);
    ask.appendChild(no);
    item.appendChild(ask);
  }

  // Section 6: right-click (or the keyboard's own menu key, which fires the
  // same event) offers the OTHER insert target for this one insert, without
  // touching the gear. Anchored to the tile rather than to the pointer, so the
  // menu cannot outlive the thing it belongs to and the audit can draw it.
  if (state.menuFor === key && state.removing === undefined && offersOtherTarget(element)) {
    const menu = el("div", "tile-menu");
    menu.setAttribute("role", "menu");
    menu.setAttribute("aria-label", element.name);
    const other = button("other-target", "tile-menu-item", otherTargetLabel(state.settings));
    other.dataset["id"] = element.id;
    other.setAttribute("role", "menuitem");
    if (state.busy === true) other.disabled = true;
    menu.appendChild(other);
    item.appendChild(menu);
  }
  return item;
}

/** The gear's panel: the settings, and the links that leave the pane. */
function gearPanel(state: PaneState): HTMLElement {
  const panel = el("div", "gear-panel");
  panel.setAttribute("role", "group");
  panel.setAttribute("aria-label", "Options");

  const target = el("div", "option");
  target.appendChild(el("span", "option-name", "Insert"));
  for (const [value, label] of [
    ["onto", "onto this slide"],
    ["new", "as a new slide"],
  ] as const) {
    const choice = button("target", state.settings.target === value ? "choice on" : "choice", label);
    choice.dataset["value"] = value;
    choice.setAttribute("aria-pressed", state.settings.target === value ? "true" : "false");
    target.appendChild(choice);
  }
  panel.appendChild(target);

  const group = el("div", "option");
  group.appendChild(el("span", "option-name", "Shapes"));
  for (const [value, label] of [
    ["group", "as one group"],
    ["loose", "loose"],
  ] as const) {
    const on = (value === "group") === state.settings.group;
    const choice = button("group", on ? "choice on" : "choice", label);
    choice.dataset["value"] = value;
    choice.setAttribute("aria-pressed", on ? "true" : "false");
    group.appendChild(choice);
  }
  panel.appendChild(group);

  const colours = el("div", "option");
  colours.appendChild(el("span", "option-name", "Colours"));
  for (const [value, label] of [
    ["deck", "this deck's theme"],
    ["library", "as in the library"],
  ] as const) {
    const on = state.settings.colours === value;
    const choice = button("colours", on ? "choice on" : "choice", label);
    choice.dataset["value"] = value;
    choice.setAttribute("aria-pressed", on ? "true" : "false");
    colours.appendChild(choice);
  }
  panel.appendChild(colours);

  // The two links that leave the pane (`docs/DESIGN.md` section 7). Buttons
  // rather than anchors, deliberately: a task pane is a WebView with no browser
  // around it, so an `<a href>` either does nothing or — worse — navigates the
  // pane away from itself, and the user's way back is to close and reopen it.
  // `openExternal` in `src/office` is what actually opens them.
  const links = el("div", "option links");
  for (const [action, label] of [
    ["report", "Report a problem"],
    ["catalogue", "Browse the catalogue on the site"],
  ] as const) {
    links.appendChild(button(action, "link", label));
  }
  panel.appendChild(links);
  return panel;
}

/**
 * "Used in this deck": what the deck already carries, and where.
 *
 * `docs/DESIGN.md` section 4, with one deviation the record now carries: it is
 * drawn CLOSED, as a button, until the user asks. Section 4 describes the list
 * as simply being there, and that would mean reading the user's whole deck
 * every time the pane opens — which is section 13's sixth open question, still
 * unanswered, on a deck that can be fifty megabytes. Asking costs one click and
 * the pane is usable in the meantime; reading on open costs everybody the
 * unknown, including the people who never look at this list.
 *
 * Three states, and the difference between two of them is the point: not asked
 * yet, asked and empty, asked and answered. "Nothing from the library is in
 * this deck yet" is a fact; a blank section is not.
 */
function usedSection(state: PaneState, library: Library): HTMLElement {
  const section = el("section", "used");
  section.setAttribute("aria-label", "Used in this deck");

  if (state.used === undefined) {
    const ask = button("used", "link", "See what this deck already uses");
    if (state.reading === true || state.busy === true) ask.disabled = true;
    ask.textContent = state.reading === true ? "Reading this deck…" : "See what this deck already uses";
    section.appendChild(ask);
    return section;
  }

  section.appendChild(el("h2", "used-head", usedHeading(state)));
  const rows = usedRows(library, state.used);
  if (rows.length > 0) {
    const list = el("ul", "used-list");
    for (const row of rows) {
      const item = el("li", row.known ? "used-row" : "used-row unknown");
      item.appendChild(el("span", "used-name", row.name));
      // Section 4: a number jumps to that slide. The call behind it is
      // PowerPointApi 1.5 and the pane knows at boot whether the host has it;
      // without it the numbers stay TEXT, because a control that might do
      // nothing is worse than a sentence that says where the element is.
      const where = el("span", "used-where");
      if (state.canJump === true) {
        for (const part of slideParts(row.slides)) {
          if ("text" in part) {
            where.appendChild(document.createTextNode(part.text));
          } else {
            const go = button("jump", "jump", String(part.slide));
            go.dataset["value"] = String(part.slide);
            go.setAttribute("aria-label", `Go to slide ${part.slide}`);
            if (state.reading === true || state.busy === true) go.disabled = true;
            where.appendChild(go);
          }
        }
      } else {
        where.textContent = row.where;
      }
      item.appendChild(where);
      list.appendChild(item);
    }
    section.appendChild(list);
  }
  // Asked again, because the answer is a snapshot: anything done to the deck
  // outside this pane since the read is not in it.
  const again = button("used", "link", "Read this deck again");
  if (state.reading === true || state.busy === true) again.disabled = true;
  section.appendChild(again);
  return section;
}

/** The footer: what happened last, and what can still be done about it. */
function footer(state: PaneState): HTMLElement {
  const bar = el("footer", "footer");
  const report = footerOf(state);
  if (report.detail !== "") {
    bar.appendChild(el("p", report.byHand ? "outcome by-hand" : "outcome", report.detail));
  }
  const actions = el("div", "actions");
  if (report.move) {
    // First, which is the order `docs/DESIGN.md` section 6 lists the three in.
    // Not for the wrap: at 320 px the row wraps to two lines with Undo alone on
    // the second whichever way round these two go — measured on 2026-09-12,
    // both orders, in `/tmp/pane-shots`. It is the record's order and nothing
    // more, which is the honest reason to have picked it.
    const move = button("move", "secondary", "Move to a new slide");
    if (state.busy === true) move.disabled = true;
    actions.appendChild(move);
  }
  if (report.again) {
    const again = button("again", "secondary", "Again");
    if (state.busy === true) again.disabled = true;
    actions.appendChild(again);
  }
  if (report.undo > 0) {
    const undo = button("undo", "secondary", `Undo (${report.undo})`);
    if (state.busy === true) undo.disabled = true;
    actions.appendChild(undo);
  }
  if (actions.childNodes.length > 0) bar.appendChild(actions);
  bar.appendChild(button("gear", "settings-line", settingsLine(state.settings)));
  return bar;
}

/** The browse screen: search, tags, favourites, recent, and the categories. */
function browse(main: HTMLElement, state: PaneState, library: Library): void {
  // Section 4: on the first open ever, three coach marks, dismissed once.
  // Drawn before the search so it is the first thing read, and it is a panel of
  // three lines rather than three callouts — see COACH for why.
  if (coaching(state)) {
    const coach = el("section", "coach");
    coach.setAttribute("aria-label", "Getting started");
    const list = el("ul", "coach-list");
    for (const line of COACH) list.appendChild(el("li", "coach-line", line));
    coach.appendChild(list);
    coach.appendChild(button("coached", "secondary", "Got it"));
    main.appendChild(coach);
  }

  const tools = el("div", "tools");
  const search = el("input", "search");
  search.type = "search";
  search.value = state.query;
  search.placeholder = "Search";
  search.dataset["action"] = "search";
  search.setAttribute("aria-label", "Search the library");
  tools.appendChild(search);
  const gear = button("gear", "gear", "⚙");
  gear.setAttribute("aria-label", "Options");
  gear.setAttribute("aria-expanded", state.gear === true ? "true" : "false");
  tools.appendChild(gear);
  main.appendChild(tools);

  if (state.gear === true) main.appendChild(gearPanel(state));

  const tags = tagsOf(library);
  if (tags.length > 0) {
    const line = el("div", state.gear === true ? "tags open" : "tags");
    // A picked tag moves to the front so it stays visible when the line is
    // closed (`docs/DESIGN.md` section 4).
    const ordered = [...state.tags, ...tags.filter((t) => !state.tags.includes(t))].slice(0, TAGS_SHOWN);
    for (const tag of ordered) {
      const chip = button("tag", state.tags.includes(tag) ? "chip on" : "chip", tag);
      chip.dataset["value"] = tag;
      chip.setAttribute("aria-pressed", state.tags.includes(tag) ? "true" : "false");
      line.appendChild(chip);
    }
    main.appendChild(line);
  }

  // Section 8: while searching, the categories that have hits appear as chips
  // with counts. Only while SEARCHING — with no query the categories are
  // already the list's own headings, and a row of chips repeating them would
  // be the same information twice.
  if (state.query.trim() !== "") {
    const hits = categoryHits(library, state);
    if (hits.length > 1 || state.category !== undefined) {
      const line = el("div", "cats");
      for (const hit of hits) {
        const on = state.category === hit.key;
        const chip = button("category-chip", on ? "chip on" : "chip", `${hit.name} ${hit.count}`);
        chip.dataset["key"] = hit.key;
        chip.setAttribute("aria-pressed", on ? "true" : "false");
        chip.setAttribute("aria-label", `${hit.name}, ${hit.count} ${hit.count === 1 ? "match" : "matches"}`);
        line.appendChild(chip);
      }
      main.appendChild(line);
    }
  }

  main.appendChild(usedSection(state, library));

  const previewed = elementOf(library, state.previewing);
  if (previewed) {
    // The stylesheet opens a gutter for the docked card at 512 and wider, and
    // only while there is one to dock.
    main.classList.add("has-card");
    main.appendChild(card(previewed, library, state));
  }

  const found = groups(library, state);
  const count = tileCount(found);
  const summary = el("p", "count", count === 0 ? "Nothing matches that." : `${count} of ${library.elements.length}`);
  if (offersOpenAll(state, library)) {
    // Section 4 puts it beside the count, so the two share a row rather than
    // the button taking a line of its own in a 320px pane.
    const row = el("div", "count-row");
    row.appendChild(summary);
    const all = button("open-all", "link open-all", "Open all");
    row.appendChild(all);
    main.appendChild(row);
  } else {
    main.appendChild(summary);
  }

  if (count === 0) {
    // A search that found nothing is a dead end; section 8 turns it into a
    // route. Every suggestion is a name the library really has, so each one is
    // a search that will find something.
    const meant = didYouMean(library, state.query);
    if (meant.length > 0) {
      const did = el("div", "meant");
      did.appendChild(el("span", "meant-lead", "Did you mean"));
      for (const name of meant) {
        const guess = button("guess", "chip", name);
        guess.dataset["value"] = name;
        did.appendChild(guess);
      }
      main.appendChild(did);
    }
    const clear = button("clear", "secondary", "Clear the search");
    main.appendChild(clear);
    return;
  }

  for (const list of [
    { key: "favourites", name: "Favourites", ids: state.favourites },
    { key: "recent", name: "Recent", ids: state.recent },
  ]) {
    const elements = list.ids.map((id) => elementOf(library, id)).filter((e): e is Element => e !== undefined);
    if (elements.length === 0) continue;
    const section = el("section", "category");
    section.appendChild(el("h2", "category-name", list.name));
    const tiles = el("ul", "tiles");
    for (const element of elements) tiles.appendChild(tile(state, library, element, list.key));
    section.appendChild(tiles);
    main.appendChild(section);
  }

  for (const group of found) {
    const section = el("section", "category");
    const head = button("category", "category-head", `${group.name} (${group.elements.length})`);
    head.dataset["key"] = group.key;
    const open = isOpen(state, group.key);
    head.setAttribute("aria-expanded", open ? "true" : "false");
    section.appendChild(head);
    if (open) {
      const tiles = el("ul", "tiles");
      for (const element of group.elements) tiles.appendChild(tile(state, library, element, group.key));
      section.appendChild(tiles);
    }
    main.appendChild(section);
  }
}

/**
 * Draw the pane into `root`, replacing whatever was there.
 *
 * Replaced, never appended: a redraw that appended would show two panes, and a
 * redraw is the ordinary way state reaches the screen.
 */
export function render(root: HTMLElement, state: PaneState, step: StepId): void {
  root.textContent = "";
  const main = el("main");
  main.append(el("span", "tick"));
  main.append(el("h1", undefined, STEP_TITLE[step]));

  const borrowed = borrowedLine(state);
  if (borrowed !== undefined) main.append(el("p", "borrowed", borrowed));

  if (step === "browse" && state.library) {
    main.append(el("p", "slide", slideLine(state)));
    browse(main, state, state.library);
  }

  const reason = blockedReason(state, step);
  if (reason !== "") main.append(el("p", "blocked", reason));
  if (state.notice) main.append(el("p", "notice", state.notice));

  if (step === "browse") main.append(footer(state));

  const action = primary(state, step);
  const control = el("button", "primary", action.label);
  control.type = "button";
  control.dataset["action"] = step === "problem" ? "retry" : "insert";
  control.disabled = action.disabled;
  // Last, always: one thing to press per screen, and it is at the bottom.
  main.append(control);
  root.append(main);
}
