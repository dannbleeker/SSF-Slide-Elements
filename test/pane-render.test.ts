/**
 * @vitest-environment jsdom
 */
import { beforeEach, describe, expect, it } from "vitest";
import type { Element } from "../src/core/catalogue/types.js";
import { render } from "../src/pane/render.js";
import { EMPTY, STEP_TITLE, blockedReason, primary, type Library, type PaneState } from "../src/pane/steps.js";

/**
 * The pane's DOM, in the one environment the suite has for it.
 *
 * jsdom has no layout and no colour, so nothing here can say the pane LOOKS
 * right — that is what `scripts/pane-shots.mjs` is for. What it can pin is what
 * is drawn, in what order, that every control carries the `data-action` the
 * shot audit reads, and that a string from outside reaches the page as text and
 * never as an element.
 */
let root: HTMLElement;

beforeEach(() => {
  document.body.innerHTML = '<header><b>SSF</b><span>Slide Elements</span></header><div id="pane"></div>';
  root = document.getElementById("pane") as HTMLElement;
});

function element(over: Partial<Element> & { id: string }): Element {
  return {
    key: over.key ?? over.id,
    name: over.name ?? over.id,
    category: over.category ?? { key: "boxes", name: "White boxes" },
    slide: 1,
    kind: "slide",
    box: over.box ?? { x: 0.1, y: 0.2, w: 0.5, h: 0.5 },
    landing: "layout",
    shapes: 1,
    tags: over.tags ?? ["boxes"],
    markup: { xml: "", rels: [], parts: [] },
    ...over,
  };
}

const LIBRARY: Library = {
  size: "16:9",
  width: 12192000,
  height: 6858000,
  version: "v1",
  categories: [{ key: "boxes", name: "White boxes" }],
  elements: [
    element({ id: "one-box", name: "One box" }),
    element({
      id: "flow-1",
      name: "Process flow, 1 box",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 1 },
    }),
    element({
      id: "flow-2",
      name: "Process flow, 2 boxes",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 2 },
    }),
  ],
};

const browsing: PaneState = { ...EMPTY, library: LIBRARY, open: ["boxes"] };

const actions = (): string[] =>
  [...root.querySelectorAll<HTMLElement>("[data-action]")].map((n) => n.dataset["action"] ?? "");

describe("every screen", () => {
  it("spends the orange budget once", () => {
    for (const state of [EMPTY, browsing, { ...EMPTY, problem: "offline" }]) {
      render(root, state, state.problem ? "problem" : state.library ? "browse" : "loading");
      expect(root.querySelectorAll(".tick").length).toBe(1);
    }
  });

  it("draws the heading for the step it was given", () => {
    render(root, EMPTY, "loading");
    expect(root.querySelector("h1")?.textContent).toBe(STEP_TITLE.loading);
    render(root, browsing, "browse");
    expect(root.querySelector("h1")?.textContent).toBe(STEP_TITLE.browse);
  });

  it("puts the primary control last and labels it from the decision", () => {
    render(root, browsing, "browse");
    const main = root.querySelector("main") as HTMLElement;
    const button = main.lastElementChild as HTMLButtonElement;
    expect(button.className).toBe("primary");
    expect(button.textContent).toBe(primary(browsing, "browse").label);
    expect(button.disabled).toBe(true);
    expect(button.dataset["action"]).toBe("insert");
  });

  it("gives every control a data-action, which is what the shot audit reads", () => {
    render(root, browsing, "browse");
    const controls = [...root.querySelectorAll("button, input")];
    expect(controls.length).toBeGreaterThan(3);
    for (const control of controls) {
      expect((control as HTMLElement).dataset["action"], control.outerHTML.slice(0, 60)).toBeTruthy();
    }
  });

  it("replaces rather than appends on a redraw", () => {
    render(root, browsing, "browse");
    render(root, browsing, "browse");
    expect(root.querySelectorAll("main").length).toBe(1);
  });

  it("puts a notice on screen as TEXT, never as markup", () => {
    // A notice comes from the host, and a host echoes its argument back. If
    // that argument ever contains markup it must not become an element.
    render(root, { ...EMPTY, notice: '<img src=x onerror="alert(1)">' }, "loading");
    expect(root.querySelector("img")).toBeNull();
    expect(root.querySelector(".notice")?.textContent).toBe('<img src=x onerror="alert(1)">');
  });

  it("puts an element's NAME on screen as text too", () => {
    // The name comes out of a JSON file fetched over the network. It is the
    // owner's own catalogue, and it still reaches the page as text.
    const nasty = { ...browsing, library: { ...LIBRARY, elements: [element({ id: "x", name: "<b>bold</b>" })] } };
    render(root, nasty, "browse");
    expect(root.querySelector("b")).toBeNull();
    expect(root.textContent).toContain("<b>bold</b>");
  });
});

describe("the problem screen", () => {
  it("says what happened and offers a retry", () => {
    const failed = { ...EMPTY, problem: "The library did not load: offline." };
    render(root, failed, "problem");
    expect(root.querySelector("p.blocked")?.textContent).toBe(blockedReason(failed, "problem"));
    const button = root.querySelector("button.primary") as HTMLButtonElement;
    expect(button.textContent).toBe("Try again");
    expect(button.disabled).toBe(false);
    expect(button.dataset["action"]).toBe("retry");
  });
});

describe("browsing", () => {
  it("draws a search box, a gear, the tags and the categories", () => {
    render(root, browsing, "browse");
    expect(actions()).toContain("search");
    expect(actions()).toContain("gear");
    expect(actions()).toContain("tag");
    expect(actions()).toContain("category");
  });

  it("keeps the typed query in the box, so a redraw does not swallow it", () => {
    render(root, { ...browsing, query: "flow" }, "browse");
    const search = root.querySelector<HTMLInputElement>('[data-action="search"]');
    expect(search?.value).toBe("flow");
  });

  it("draws a tile per element in an open category, with a ghost and a name", () => {
    render(root, browsing, "browse");
    const tiles = root.querySelectorAll('[data-action="tile"]');
    // One box, plus the first member of the flow run: the second is behind the
    // stepper rather than a tile of its own.
    expect(tiles.length).toBe(2);
    const first = tiles[0] as HTMLElement;
    expect(first.dataset["id"]).toBe("one-box");
    expect(first.querySelector("svg.ghost")).not.toBeNull();
    expect(first.querySelector(".tile-name")?.textContent).toBe("One box");
  });

  it("hides the tiles of a closed category but keeps its header", () => {
    render(root, { ...browsing, open: [] }, "browse");
    expect(root.querySelectorAll('[data-action="tile"]').length).toBe(0);
    const head = root.querySelector('[data-action="category"]');
    expect(head?.getAttribute("aria-expanded")).toBe("false");
    expect(head?.textContent).toContain("White boxes");
  });

  it("gives a sized element a stepper naming what it counts", () => {
    render(root, browsing, "browse");
    const stepper = root.querySelector(".stepper");
    expect(stepper?.querySelector(".stepper-noun")?.textContent).toBe("boxes");
    expect([...root.querySelectorAll('[data-action="step"]')].map((n) => n.textContent)).toEqual(["1", "2"]);
  });

  it("marks the chosen tile so the keyboard has somewhere to be", () => {
    render(root, { ...browsing, chosen: "one-box" }, "browse");
    const chosen = root.querySelector('[data-action="tile"][aria-current="true"]') as HTMLElement;
    expect(chosen.dataset["id"]).toBe("one-box");
  });

  it("says Inserting on the tile that is going, and disables the rest", () => {
    render(root, { ...browsing, chosen: "one-box", busy: true }, "browse");
    expect(root.textContent).toContain("Inserting…");
    for (const tile of root.querySelectorAll<HTMLButtonElement>('[data-action="tile"]')) {
      expect(tile.disabled).toBe(true);
    }
  });

  it("draws a star per tile that says what pressing it would do", () => {
    render(root, browsing, "browse");
    const star = root.querySelector('[data-action="star"]') as HTMLElement;
    expect(star.getAttribute("aria-pressed")).toBe("false");
    expect(star.getAttribute("aria-label")).toContain("Add");
    render(root, { ...browsing, favourites: ["one-box"] }, "browse");
    const starred = root.querySelector('[data-action="star"]') as HTMLElement;
    expect(starred.getAttribute("aria-pressed")).toBe("true");
    expect(starred.getAttribute("aria-label")).toContain("Remove");
  });

  it("shows favourites and recent above the categories, and only when there are any", () => {
    render(root, browsing, "browse");
    expect(root.textContent).not.toContain("Favourites");
    render(root, { ...browsing, favourites: ["one-box"], recent: ["flow-1"] }, "browse");
    const headings = [...root.querySelectorAll("h2")].map((h) => h.textContent);
    expect(headings).toEqual(["Favourites", "Recent"]);
  });

  it("ignores a favourite the catalogue no longer has", () => {
    // A star kept from an older catalogue names an element that may be gone.
    render(root, { ...browsing, favourites: ["vanished"] }, "browse");
    expect(root.textContent).not.toContain("Favourites");
  });

  it("says so when a search found nothing, and offers to clear it", () => {
    render(root, { ...browsing, query: "nothing matches this" }, "browse");
    expect(root.querySelector(".count")?.textContent).toMatch(/nothing matches/i);
    expect(actions()).toContain("clear");
    expect(root.querySelectorAll('[data-action="tile"]').length).toBe(0);
  });

  it("opens the gear into a panel of pressable options", () => {
    render(root, browsing, "browse");
    expect(root.querySelector(".gear-panel")).toBeNull();
    render(root, { ...browsing, gear: true }, "browse");
    const panel = root.querySelector(".gear-panel") as HTMLElement;
    expect(panel.getAttribute("role")).toBe("group");
    const targets = [...panel.querySelectorAll<HTMLElement>('[data-action="target"]')];
    expect(targets.map((t) => t.dataset["value"])).toEqual(["onto", "new"]);
    expect(targets[0]?.getAttribute("aria-pressed")).toBe("true");
    const groups = [...panel.querySelectorAll<HTMLElement>('[data-action="group"]')];
    expect(groups[0]?.getAttribute("aria-pressed")).toBe("true");
  });

  it("draws the footer with the outcome, Again and Undo", () => {
    const after = {
      ...browsing,
      recent: ["one-box"],
      undo: 1,
      outcome: { ok: true, byHand: false, name: "One box", detail: "3 → 4 → 3 slides, slide 2 replaced." },
    };
    render(root, after, "browse");
    expect(root.querySelector(".outcome")?.textContent).toBe("3 → 4 → 3 slides, slide 2 replaced.");
    expect(actions()).toContain("again");
    expect(actions()).toContain("undo");
    expect(root.querySelector('[data-action="undo"]')?.textContent).toBe("Undo (1)");
  });

  it("marks an outcome the user has to put right by hand", () => {
    const stuck = {
      ...browsing,
      outcome: { ok: false, byHand: true, name: "One box", detail: "delete slide 2 by hand." },
    };
    render(root, stuck, "browse");
    expect(root.querySelector(".outcome.by-hand")).not.toBeNull();
  });

  it("draws the settings line as a way into the gear", () => {
    render(root, browsing, "browse");
    const line = root.querySelector(".settings-line") as HTMLElement;
    expect(line.textContent).toBe("Inserting onto this slide, as one group.");
    expect(line.dataset["action"]).toBe("gear");
  });

  it("draws the borrowed-library line only when a library was borrowed", () => {
    render(root, browsing, "browse");
    expect(root.querySelector(".borrowed")).toBeNull();
    render(root, { ...browsing, library: { ...LIBRARY, borrowed: "4:3 library, scaled to A4 slides." } }, "browse");
    expect(root.querySelector(".borrowed")?.textContent).toBe("4:3 library, scaled to A4 slides.");
  });
});

describe("the ghost on a tile", () => {
  it("draws the slide's shape and the element's box on it, and hides it from a reader", () => {
    render(root, browsing, "browse");
    const svg = root.querySelector("svg.ghost") as SVGSVGElement;
    // 16:9 at a hundred tall.
    expect(svg.getAttribute("viewBox")).toBe("0 0 178 100");
    expect(svg.getAttribute("aria-hidden")).toBe("true");
    const box = svg.querySelector(".ghost-box") as SVGRectElement;
    // The element's box is 10% across and 20% down, so on a 178-wide viewBox
    // that is 17.8 and 20.
    expect(box.getAttribute("x")).toBe("17.8");
    expect(box.getAttribute("y")).toBe("20.0");
  });

  it("never draws a box too small to see", () => {
    const hair = {
      ...browsing,
      library: { ...LIBRARY, elements: [element({ id: "hair", box: { x: 0, y: 0, w: 0.001, h: 0.001 } })] },
    };
    render(root, hair, "browse");
    const box = root.querySelector(".ghost-box") as SVGRectElement;
    expect(Number(box.getAttribute("width"))).toBeGreaterThanOrEqual(2);
    expect(Number(box.getAttribute("height"))).toBeGreaterThanOrEqual(2);
  });
});

describe("the element's photograph on a tile", () => {
  const shotOf = (): HTMLImageElement => {
    render(root, browsing, "browse");
    return root.querySelector(".tile-img") as HTMLImageElement;
  };

  it("points at the preview cut from the deck's print, for this size and build", () => {
    const img = shotOf();
    expect(img.getAttribute("src")).toBe(`./catalogue/16x9/previews/one-box.png?v=${LIBRARY.version}`);
  });

  it("carries the catalogue version, so a rebuilt preview is not served from cache", () => {
    // The file name is derived from the element id rather than hashed, so it
    // does not change when the picture does; the version is the only thing that
    // tells a browser to look again.
    expect(shotOf().getAttribute("src")).toContain(`?v=${LIBRARY.version}`);
  });

  it("is decorative, because the name beside it is the label", () => {
    const img = shotOf();
    expect(img.getAttribute("alt")).toBe("");
    expect(img.getAttribute("loading")).toBe("lazy");
  });

  it("sits ON the ghost rather than instead of it, so nothing reflows when it lands", () => {
    render(root, browsing, "browse");
    const frame = root.querySelector(".tile-shot") as HTMLElement;
    expect(frame.querySelector("svg.ghost")).not.toBeNull();
    expect(frame.querySelector(".tile-img")).not.toBeNull();
  });

  it("takes itself out when there is no picture, leaving the ghost", () => {
    // A tree where the previews have not been built, or a deploy that has not
    // caught up with a new element.
    render(root, browsing, "browse");
    const tile = root.querySelector('[data-action="tile"]') as HTMLElement;
    const img = tile.querySelector(".tile-img") as HTMLImageElement;
    img.dispatchEvent(new Event("error"));
    expect(tile.querySelector(".tile-img")).toBeNull();
    expect(tile.querySelector("svg.ghost")).not.toBeNull();
    // and only that tile's: a preview missing for one element says nothing
    // about the others
    expect(root.querySelectorAll(".tile-img").length).toBe(1);
  });
});

describe("the preview card", () => {
  const previewing = { ...browsing, previewing: "one-box" };

  it("is not there until something is being previewed", () => {
    render(root, browsing, "browse");
    expect(root.querySelector(".card")).toBeNull();
    expect(root.querySelector("main")?.classList.contains("has-card")).toBe(false);
  });

  it("shows the element at full width, its name, and where it lands", () => {
    render(root, previewing, "browse");
    const card = root.querySelector(".card") as HTMLElement;
    expect(card.dataset["id"]).toBe("one-box");
    expect(card.querySelector(".tile-img")).not.toBeNull();
    expect(card.querySelector(".card-name")?.textContent).toBe("One box");
    expect(card.querySelector(".card-landing")?.textContent).toContain("Lands");
  });

  it("follows the gear, so the line matches what the insert would do", () => {
    render(root, { ...previewing, settings: { target: "new", group: true } }, "browse");
    expect(root.querySelector(".card-landing")?.textContent).toBe("Lands as a new slide after this one.");
  });

  it("opens the gutter the dock sits in, and only while it is open", () => {
    render(root, previewing, "browse");
    expect(root.querySelector("main")?.classList.contains("has-card")).toBe(true);
  });

  it("is hidden from a reader, because the tile it describes already announces itself", () => {
    render(root, previewing, "browse");
    expect(root.querySelector(".card")?.getAttribute("aria-hidden")).toBe("true");
  });

  it("does not appear for an id the library does not have", () => {
    // A stale id — the category closed, or the search narrowed — must not draw
    // an empty card.
    render(root, { ...browsing, previewing: "no-such-element" }, "browse");
    expect(root.querySelector(".card")).toBeNull();
  });
});

describe("a search that found nothing", () => {
  const missed = { ...browsing, query: "trianglee" };

  it("offers what the user might have meant, as chips", () => {
    render(root, { ...browsing, query: "bax" }, "browse");
    // the query really did find nothing …
    expect(root.querySelector(".count")?.textContent).toBe("Nothing matches that.");
    // … and the dead end has a way out of it
    const guesses = [...root.querySelectorAll('[data-action="guess"]')];
    expect(guesses.length).toBeGreaterThan(0);
    expect(guesses.map((g) => g.textContent)).toContain("One box");
  });

  it("offers nothing when the query is near nothing, and still offers to clear", () => {
    render(root, { ...browsing, query: "xylophone" }, "browse");
    expect(root.querySelector(".meant")).toBeNull();
    expect(root.querySelector('[data-action="clear"]')).not.toBeNull();
  });

  it("puts a name the library really has on each chip", () => {
    render(root, missed, "browse");
    const guesses = [...root.querySelectorAll('[data-action="guess"]')] as HTMLElement[];
    const names = new Set(LIBRARY.elements.map((e) => e.name));
    for (const guess of guesses) {
      expect(names.has(guess.dataset["value"] ?? "")).toBe(true);
      expect(guess.textContent).toBe(guess.dataset["value"]);
    }
  });
});

describe("the category chips", () => {
  // The library above has one category, so it could never show a chip — the
  // chips exist to choose BETWEEN categories. This one has two.
  const TWO: Library = {
    ...LIBRARY,
    categories: [
      { key: "boxes", name: "White boxes" },
      { key: "stamps", name: "Stamps and labels" },
    ],
    elements: [
      element({ id: "one-box", name: "One box" }),
      element({
        id: "draft",
        name: "Draft box stamp",
        kind: "part",
        category: { key: "stamps", name: "Stamps and labels" },
      }),
    ],
  };
  const searching: PaneState = { ...EMPTY, library: TWO, open: ["boxes", "stamps"], query: "box" };

  it("are not drawn until something is being searched for", () => {
    // With no query the categories are already the list's own headings, and a
    // row of chips repeating them is the same information twice.
    render(root, { ...searching, query: "" }, "browse");
    expect(root.querySelector(".cats")).toBeNull();
  });

  it("are not drawn when only one category has hits", () => {
    // A chip narrowing to the only category on screen chooses nothing.
    render(root, { ...searching, query: "stamp" }, "browse");
    expect(root.querySelector(".cats")).toBeNull();
  });

  it("appear while searching, each naming a category and its count", () => {
    render(root, searching, "browse");
    const chips = [...root.querySelectorAll('[data-action="category-chip"]')] as HTMLElement[];
    expect(chips.length).toBe(2);
    for (const chip of chips) {
      expect(chip.dataset["key"]).toBeTruthy();
      expect(chip.textContent).toMatch(/\s\d+$/);
    }
  });

  it("marks the picked one pressed, and keeps the others reachable", () => {
    render(root, { ...searching, category: "stamps" }, "browse");
    const chips = [...root.querySelectorAll('[data-action="category-chip"]')] as HTMLElement[];
    expect(chips.length).toBe(2);
    const picked = chips.filter((c) => c.getAttribute("aria-pressed") === "true");
    expect(picked).toHaveLength(1);
    expect(picked[0]?.dataset["key"]).toBe("stamps");
  });
});

describe("the stepper while searching", () => {
  it("greys the sizes the search did not ask for, and leaves them pickable", () => {
    render(root, { ...browsing, query: "1 box" }, "browse");
    const steps = [...root.querySelectorAll('[data-action="step"]')] as HTMLButtonElement[];
    expect(steps.length).toBeGreaterThan(1);
    const off = steps.filter((s) => s.className.includes("off"));
    expect(off.length).toBeGreaterThan(0);
    for (const step of off) {
      expect(step.disabled).toBe(false);
      expect(step.getAttribute("aria-label")).toContain("not a match");
    }
  });

  it("greys nothing when nothing is being searched for", () => {
    render(root, browsing, "browse");
    const steps = [...root.querySelectorAll('[data-action="step"]')] as HTMLElement[];
    expect(steps.filter((s) => s.className.includes("off"))).toHaveLength(0);
  });
});

describe("the first-run coach marks", () => {
  const first: PaneState = { ...browsing, coached: undefined };

  it("say three things and offer one way out", () => {
    render(root, first, "browse");
    expect(root.querySelectorAll(".coach-line")).toHaveLength(3);
    expect(root.querySelectorAll('[data-action="coached"]')).toHaveLength(1);
  });

  it("are gone once dismissed", () => {
    render(root, { ...first, coached: true }, "browse");
    expect(root.querySelector(".coach")).toBeNull();
  });

  it("are read before the search, because they explain what the search is for", () => {
    render(root, first, "browse");
    const order = [...root.querySelectorAll(".coach, [data-action='search']")];
    expect(order[0]?.className).toContain("coach");
  });

  it("carry a name, so a reader knows what the panel is", () => {
    render(root, first, "browse");
    expect(root.querySelector(".coach")?.getAttribute("aria-label")).toBe("Getting started");
  });
});
