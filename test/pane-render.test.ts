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

  it("offers the two ways out of the pane, as buttons rather than links", () => {
    // `docs/DESIGN.md` section 7. Buttons on purpose: an <a href> in a task
    // pane either does nothing or navigates the pane away from itself, and a
    // user whose pane has become a web page has to close and reopen it.
    render(root, { ...browsing, gear: true }, "browse");
    const panel = root.querySelector(".gear-panel") as HTMLElement;
    for (const action of ["report", "catalogue"]) {
      const link = panel.querySelector(`[data-action="${action}"]`);
      expect(link, action).not.toBeNull();
      expect(link?.tagName).toBe("BUTTON");
    }
    expect(panel.querySelector("a")).toBeNull();
  });

  it("offers the colour switch, with the setting the user is on pressed", () => {
    // `docs/DESIGN.md` section 7's third option. The pressed one is the
    // EVIDENCE half: a panel that drew both choices and pressed neither would
    // pass a test that only counted them, and would leave the user unable to
    // see which way the switch is set.
    render(root, { ...browsing, gear: true }, "browse");
    const choices = [...root.querySelectorAll<HTMLElement>('[data-action="colours"]')];
    expect(choices.map((c) => c.dataset["value"])).toEqual(["deck", "library"]);
    expect(choices.map((c) => c.getAttribute("aria-pressed"))).toEqual(["true", "false"]);

    const pinned = { ...browsing, gear: true, settings: { ...browsing.settings, colours: "library" as const } };
    render(root, pinned, "browse");
    expect(
      [...root.querySelectorAll<HTMLElement>('[data-action="colours"]')].map((c) => c.getAttribute("aria-pressed")),
    ).toEqual(["false", "true"]);
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

  it("draws Move to a new slide first, ahead of Again and Undo", () => {
    // The order `docs/DESIGN.md` section 6 lists the three in.
    const covered = {
      ...browsing,
      recent: ["one-box"],
      undo: 1,
      moveable: "one-box",
      outcome: { ok: true, byHand: false, name: "One box", detail: "3 → 4 → 3 slides, slide 2 replaced." },
    };
    render(root, covered, "browse");
    const drawn = actions();
    expect(drawn).toContain("move");
    expect(root.querySelector('[data-action="move"]')?.textContent).toBe("Move to a new slide");
    expect(drawn.indexOf("move")).toBeLessThan(drawn.indexOf("again"));
    expect(drawn.indexOf("again")).toBeLessThan(drawn.indexOf("undo"));
  });

  it("draws no Move to a new slide when the element did not land on anything", () => {
    const plain = {
      ...browsing,
      recent: ["one-box"],
      undo: 1,
      outcome: { ok: true, byHand: false, name: "One box", detail: "3 → 4 → 3 slides, slide 2 replaced." },
    };
    render(root, plain, "browse");
    expect(actions()).not.toContain("move");
  });

  it("disables Move to a new slide while an insert is going", () => {
    const busy = { ...browsing, recent: ["one-box"], undo: 1, moveable: "one-box", busy: true };
    render(root, busy, "browse");
    expect(root.querySelector<HTMLButtonElement>('[data-action="move"]')?.disabled).toBe(true);
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
    render(root, { ...previewing, settings: { target: "new", group: true, colours: "deck" } }, "browse");
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

describe("the tile's right-click menu", () => {
  it("draws the other target, on the tile it belongs to", () => {
    render(root, { ...browsing, menuFor: "boxes:one-box" }, "browse");
    const tile = root.querySelector('.tile [data-action="other-target"]') as HTMLElement;
    expect(tile).not.toBeNull();
    expect(tile.textContent).toBe("Insert as a new slide");
    expect(tile.dataset["id"]).toBe("one-box");
    // Anchored to its own tile, so it cannot outlive the thing it belongs to.
    expect(tile.closest(".tile")?.querySelector(".tile-name")?.textContent).toBe("One box");
    // One menu, on one tile.
    expect(root.querySelectorAll('[data-action="other-target"]').length).toBe(1);
  });

  it("follows the gear, so it always offers the one you are not on", () => {
    const asNew = { ...browsing, menuFor: "boxes:one-box", settings: { ...browsing.settings, target: "new" as const } };
    render(root, asNew, "browse");
    expect(root.querySelector('[data-action="other-target"]')?.textContent).toBe("Insert onto this slide");
  });

  it("draws nothing on a part, because a part ignores the target", () => {
    const stamps = { key: "stamps", name: "Stamps and labels" };
    const withStamp: Library = {
      ...LIBRARY,
      categories: [...LIBRARY.categories, stamps],
      elements: [
        ...LIBRARY.elements,
        element({ id: "approved", name: "Approved stamp", kind: "part", landing: "top-right", category: stamps }),
      ],
    };
    render(root, { ...browsing, library: withStamp, menuFor: "approved", open: ["stamps"] }, "browse");
    // The part's tile IS on screen — otherwise this would pass by drawing
    // nothing at all, which is the vacuous version of the same assertion.
    const tile = [...root.querySelectorAll(".tile")].find(
      (t) => t.querySelector(".tile-name")?.textContent === "Approved stamp",
    );
    expect(tile, "the part's tile is not drawn, so this case proves nothing").toBeTruthy();
    expect(root.querySelector('[data-action="other-target"]')).toBeNull();
  });

  it("is a menu to a screen reader, not a stray button", () => {
    render(root, { ...browsing, menuFor: "boxes:one-box" }, "browse");
    const menu = root.querySelector(".tile-menu") as HTMLElement;
    expect(menu.getAttribute("role")).toBe("menu");
    expect(menu.getAttribute("aria-label")).toBe("One box");
    expect(menu.querySelector('[role="menuitem"]')).not.toBeNull();
  });
});

describe("the preview card's grey boxes", () => {
  // The card is open on a tile, which is what makes there be a card at all.
  const previewing = { ...browsing, previewing: "one-box", open: ["boxes"] };
  const onSlide = {
    slide: 2,
    boxes: [
      { x: 0.1, y: 0.1, w: 0.3, h: 0.2 },
      { x: 0.5, y: 0.6, w: 0.4, h: 0.3 },
    ],
  };

  it("draws what the slide already holds, under the element's own frame", () => {
    render(root, { ...previewing, slide: 2, onSlide }, "browse");
    const card = root.querySelector(".card") as HTMLElement;
    const held = card.querySelectorAll(".ghost-held");
    expect(held.length).toBe(2);
    // In z-order and under the blue frame: a held box drawn over the landing
    // would say the element goes behind what is already there.
    const shapes = [...card.querySelectorAll("rect")].map((r) => r.getAttribute("class"));
    expect(shapes).toEqual(["ghost-slide", "ghost-held", "ghost-held", "ghost-box"]);
    expect(held[0]?.getAttribute("x")).toBe("17.8");
  });

  it("draws none when the snapshot is of another slide", () => {
    render(root, { ...previewing, slide: 5, onSlide }, "browse");
    expect(root.querySelector(".card")?.querySelectorAll(".ghost-held").length).toBe(0);
  });

  it("never draws them on a TILE, which is too small to say anything with them", () => {
    render(root, { ...previewing, slide: 2, onSlide }, "browse");
    const tile = root.querySelector(".tile") as HTMLElement;
    expect(tile.querySelectorAll(".ghost-held").length).toBe(0);
  });
});

describe("removing a part from the deck", () => {
  const stamps = { key: "stamps", name: "Stamps and labels" };
  const withStamp: Library = {
    ...LIBRARY,
    categories: [...LIBRARY.categories, stamps],
    elements: [
      ...LIBRARY.elements,
      element({ id: "approved", name: "Approved stamp", kind: "part", landing: "top-right", category: stamps }),
    ],
  };
  const read = {
    ...browsing,
    library: withStamp,
    open: ["boxes", "stamps"],
    used: [{ element: "approved", slides: [2, 5, 9] }],
  };

  it("puts the button on the part's tile once the deck has been read, and nowhere else", () => {
    render(root, { ...browsing, library: withStamp, open: ["boxes", "stamps"] }, "browse");
    expect(root.querySelector('[data-action="remove"]')).toBeNull();

    render(root, read, "browse");
    const buttons = [...root.querySelectorAll<HTMLElement>('[data-action="remove"]')];
    expect(buttons).toHaveLength(1);
    expect(buttons[0]?.textContent).toBe("Remove from 3 slides");
    expect(buttons[0]?.dataset["id"]).toBe("approved");
  });

  it("asks before it removes, names the slides, and offers a way out", () => {
    render(root, { ...read, removing: { id: "approved", slides: [2, 5, 9], done: 0, where: "stamps" } }, "browse");
    expect(root.querySelector(".tile-ask")?.textContent).toContain("slides 2, 5 and 9");
    expect(root.querySelector(".tile-ask")?.textContent).toContain("cannot undo");
    expect(root.querySelector('[data-action="remove-go"]')?.textContent).toBe("Remove");
    expect(root.querySelector('[data-action="remove-cancel"]')?.textContent).toBe("Keep them");
    // And the button that opened it is gone, so the question is the only thing
    // to answer.
    expect(root.querySelector('[data-action="remove"]')).toBeNull();
  });

  it("draws one question at a time, and not the right-click menu beside it", () => {
    const both = {
      ...read,
      menuFor: "boxes:one-box",
      removing: { id: "approved", slides: [2], done: 0, where: "stamps" },
    };
    render(root, both, "browse");
    expect(root.querySelectorAll(".tile-menu").length).toBe(1);
    expect(root.querySelector('[data-action="other-target"]')).toBeNull();
  });
});

describe("an element drawn in more than one place at once", () => {
  // Favourites, Recent and the element's own category are three lists, and one
  // element can be in all three. Everything anchored to a TILE has to know
  // which tile, or one right-click opens three menus.
  const twice = {
    ...browsing,
    open: ["boxes"],
    favourites: ["one-box"],
    recent: ["one-box"],
  };

  it("draws the same element in every list it belongs to", () => {
    // The vacuity guard: without this the cases below would pass by there
    // being only one tile to find.
    render(root, twice, "browse");
    const tiles = [...root.querySelectorAll<HTMLElement>('[data-action="tile"][data-id="one-box"]')];
    expect(tiles.length).toBe(3);
    expect(tiles.map((t) => t.dataset["where"])).toEqual(["favourites", "recent", "boxes"]);
  });

  it("opens the right-click menu on the tile that was clicked, not on every copy", () => {
    render(root, { ...twice, menuFor: "recent:one-box" }, "browse");
    const menus = [...root.querySelectorAll(".tile-menu")];
    expect(menus.length).toBe(1);
    const tile = menus[0]?.closest(".tile");
    expect(tile?.querySelector<HTMLElement>('[data-action="tile"]')?.dataset["where"]).toBe("recent");
  });

  it("asks the removal question once, on the tile that asked it", () => {
    const stamps = { key: "stamps", name: "Stamps and labels" };
    const withStamp: Library = {
      ...LIBRARY,
      categories: [...LIBRARY.categories, stamps],
      elements: [
        ...LIBRARY.elements,
        element({ id: "approved", name: "Approved stamp", kind: "part", landing: "top-right", category: stamps }),
      ],
    };
    const state = {
      ...browsing,
      library: withStamp,
      open: ["boxes", "stamps"],
      favourites: ["approved"],
      used: [{ element: "approved", slides: [2, 5] }],
      removing: { id: "approved", slides: [2, 5], done: 0, where: "stamps" },
    };
    render(root, state, "browse");
    expect(root.querySelectorAll(".tile-ask").length).toBe(1);
    const asked = root.querySelector(".tile-ask")?.closest(".tile");
    expect(asked?.querySelector<HTMLElement>('[data-action="tile"]')?.dataset["where"]).toBe("stamps");
  });
});

describe("the slide numbers in Used in this deck", () => {
  const read: PaneState = {
    ...browsing,
    used: [
      { element: "one-box", slides: [2] },
      { element: "two-boxes", slides: [11, 3, 5] },
    ],
  };

  it("are links when the host can go to a slide, one per number, with the words kept as words", () => {
    render(root, { ...read, canJump: true }, "browse");
    const rows = [...root.querySelectorAll(".used-row")];
    const buttons = rows[1]?.querySelectorAll<HTMLButtonElement>('[data-action="jump"]') ?? [];
    expect([...buttons].map((b) => b.dataset["value"])).toEqual(["3", "5", "11"]);
    expect([...buttons].map((b) => b.getAttribute("aria-label"))).toEqual([
      "Go to slide 3",
      "Go to slide 5",
      "Go to slide 11",
    ]);
    // The sentence still reads as one: "slides 3, 5 and 11".
    expect(rows[1]?.querySelector(".used-where")?.textContent).toBe("slides 3, 5 and 11");
    expect(rows[0]?.querySelector(".used-where")?.textContent).toBe("slide 2");
    for (const b of buttons) expect(b.disabled).toBe(false);
  });

  it("stay text on a host that cannot go to a slide", () => {
    // A control that might do nothing is worse than a sentence that says where
    // the element is (docs/DESIGN.md section 4).
    render(root, read, "browse");
    expect(root.querySelector('[data-action="jump"]')).toBeNull();
    expect(root.querySelectorAll(".used-row")[1]?.querySelector(".used-where")?.textContent).toBe("slides 3, 5 and 11");
  });

  it("wait while an insert or a read is running", () => {
    render(root, { ...read, canJump: true, busy: true }, "browse");
    const buttons = [...root.querySelectorAll<HTMLButtonElement>('[data-action="jump"]')];
    expect(buttons.length).toBe(4);
    for (const b of buttons) expect(b.disabled).toBe(true);
  });
});

describe("Open all beside the count", () => {
  it("is offered while a category is closed, and shares the count's row", () => {
    render(root, { ...browsing, open: [] }, "browse");
    const all = root.querySelector<HTMLElement>('[data-action="open-all"]');
    expect(all?.textContent).toBe("Open all");
    expect(all?.closest(".count-row")?.querySelector(".count")).not.toBeNull();
  });

  it("is gone once every category is open", () => {
    // A control that has done its job and cannot do it again is worse than no
    // control: it invites a click that changes nothing.
    const open = LIBRARY.categories.map((c) => c.key);
    render(root, { ...browsing, open }, "browse");
    expect(root.querySelector('[data-action="open-all"]')).toBeNull();
    // The count is still drawn, on its own, when the row is not.
    expect(root.querySelector(".count")).not.toBeNull();
  });

  it("is gone while searching, which opens what it finds by itself", () => {
    render(root, { ...browsing, open: [], query: "box" }, "browse");
    expect(root.querySelector('[data-action="open-all"]')).toBeNull();
  });
});
