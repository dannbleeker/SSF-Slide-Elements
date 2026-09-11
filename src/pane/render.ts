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
import type { Element } from "../core/catalogue/types.js";
import { previewUrl } from "./catalogue.js";
import {
  STEP_TITLE,
  blockedReason,
  borrowedLine,
  elementOf,
  footerOf,
  groups,
  isOpen,
  landingLine,
  primary,
  runOf,
  settingsLine,
  slideLine,
  tagsOf,
  tileCount,
  type Library,
  type PaneState,
  type StepId,
} from "./steps.js";

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
function ghost(element: Element, aspect: number): SVGSVGElement {
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
function picture(element: Element, library: Library): HTMLElement {
  const frame = el("span", "tile-shot");
  frame.appendChild(ghost(element, library.width / library.height));

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
 * NOT here yet, and named so it is not mistaken for done: the grey boxes for
 * what the destination slide already has. The pane does not read the slide's
 * shapes, and adding that read is host work this could not verify without a
 * round against a real PowerPoint. The ghost frame for the landing is drawn
 * from the catalogue, which needs no host at all.
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
  box.appendChild(picture(element, library));
  box.appendChild(el("strong", "card-name", element.name));
  box.appendChild(el("p", "card-landing", landingLine(element, state.settings)));
  return box;
}

/** One element's tile: where it lands, what it is called, and whether it is starred. */
function tile(state: PaneState, library: Library, element: Element): HTMLElement {
  const item = el("li", "tile");
  const pick = button("tile", "tile-pick", "");
  pick.dataset["id"] = element.id;
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
      const step = button("step", state.chosen === member.id ? "step on" : "step", String(member.run?.count ?? ""));
      step.dataset["id"] = member.id;
      step.setAttribute("aria-label", `${member.name}`);
      if (state.busy === true) step.disabled = true;
      stepper.appendChild(step);
    }
    item.appendChild(stepper);
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
  return panel;
}

/** The footer: what happened last, and what can still be done about it. */
function footer(state: PaneState): HTMLElement {
  const bar = el("footer", "footer");
  const report = footerOf(state);
  if (report.detail !== "") {
    bar.appendChild(el("p", report.byHand ? "outcome by-hand" : "outcome", report.detail));
  }
  const actions = el("div", "actions");
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
  main.appendChild(summary);

  if (count === 0) {
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
    for (const element of elements) tiles.appendChild(tile(state, library, element));
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
      for (const element of group.elements) tiles.appendChild(tile(state, library, element));
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
