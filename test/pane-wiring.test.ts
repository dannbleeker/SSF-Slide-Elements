/**
 * @vitest-environment jsdom
 */
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { TAG_CATALOGUE, TAG_ELEMENT, writeShapeTags } from "../src/core/pptx/tags.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The pane's entry point, wired.
 *
 * `pane-render.test.ts` calls `render` directly, so `main.ts` — the only file
 * in `src/pane` that touches Office.js, the other being `src/office/powerpoint.ts`,
 * which `office-host.test.ts` covers — would otherwise run nowhere in the suite. It runs
 * here with `Office` stubbed to the two things it reads, and with the host
 * calls and the catalogue fetch mocked, because neither a PowerPoint nor a site
 * to fetch from exists in a test runner.
 *
 * What this file is for is the WIRING, not the decisions: that the pane boots,
 * that it stops with a sentence on a host below the floor, that the live region
 * is made once and outside the pane, and that a library which does not arrive
 * ends as a screen with a retry on it rather than a blank one.
 */
type Readiness = { ok: boolean; detail: string };
let readiness: Readiness = { ok: true, detail: "fine" };
type IndexMode = "fail" | "ok" | "hang";
let indexMode: IndexMode = "fail";

vi.mock("../src/office/powerpoint.js", () => ({
  ready: () => readiness,
  hostSupports: () => host.supports15,
  // The one selection write, recorded; the read-back answers what the case set.
  selectSlide: (id: string) => {
    host.selected.push(id);
    return Promise.resolve(host.selectAnswers ?? { supported: true, selected: [id] });
  },
  // Every call the pane can make, so a missing export cannot be mistaken for a
  // pane that decided not to make it. The last case in this file holds this
  // list against the real module, because the claim above was untrue for three
  // exports before anyone noticed.
  slideCount: () => Promise.resolve(host.slides),
  // The count the pane is waiting for, unless this run is refusing.
  countReaching: (want: number) => Promise.resolve(host.cycles === host.refuseAt ? host.slides : want),
  readDeck: () =>
    deckBase64 === undefined
      ? Promise.reject(new Error("no deck in a test runner"))
      : Promise.resolve({ base64: deckBase64, bytes: 1, ms: 1 }),
  currentSlide: () => Promise.resolve(host.current),
  selectedShape: () => Promise.resolve(undefined),
  slideIdAt: () => Promise.resolve(host.namesSlides ? "256" : undefined),
  onSlideChange: () => Promise.resolve(false),
  insertPackage: () => {
    host.cycles += 1;
    return Promise.resolve(undefined);
  },
  removeSlideAt: (at: number) => {
    host.removed.push(at);
    return Promise.resolve(host.refuseRemoval ? "the host refused" : undefined);
  },
  hostStamp: () => ({ host: "PowerPoint", platform: "PC" }),
  deckUrl: () => host.url,
  openExternal: (url: string) => {
    opened.push(url);
    return externalOpens;
  },
}));

/**
 * The splice, stubbed to record what it was asked for.
 *
 * The engine is tested against real decks elsewhere; what this file needs is
 * the one thing only the wiring knows — which insert TARGET the pane passed for
 * a given click. The stub answers a report, and the pane then stops of its own
 * accord because `currentSlide` names no slide, which is a path this file
 * already covers.
 */
const spliced: { target: string }[] = [];
vi.mock("../src/core/splice/splice.js", () => ({
  splice: (request: { options: { target: string } }) => {
    spliced.push({ target: request.options.target });
    return Promise.resolve({
      base64: "",
      deckSlides: 3,
      slidePath: "ppt/slides/slide1.xml",
      landed: { x: 0, y: 0, cx: 1, cy: 1 },
      shapes: 1,
      grouped: false,
      parts: 0,
      placeholders: 0,
      pinned: 0,
      held: host.held,
    });
  },
  onlySlide: () => Promise.resolve({ base64: "", path: "ppt/slides/slide1.xml" }),
}));

/** Every URL the pane asked the host to open, and whether the host obliged. */
const opened: string[] = [];
let externalOpens = true;
/** The deck `readDeck` hands back, when a case has built one. */
let deckBase64: string | undefined;

/**
 * A host that can be made to refuse.
 *
 * `cycles` counts insert-then-remove rounds, which is what "Remove from N
 * slides" does one of per slide; `refuseAt` is the 1-based cycle whose count
 * will NOT agree, which is the failure the run has to stop on rather than press
 * through. Zero refuses nothing.
 */
const host = {
  slides: 3,
  cycles: 0,
  refuseAt: 0,
  namesSlides: true,
  supports15: true,
  /** Every id the pane asked the host to select. */
  selected: [] as string[],
  /** What `selectSlide` answers, when a case wants something other than the slide asked for. */
  selectAnswers: undefined as { supported: boolean; selected: string[] | null; error?: string } | undefined,
  /**
   * The slide the user is on, and it is UNDEFINED by default on purpose.
   *
   * Most cases in this file want the insert to stop early — the pane refuses
   * when the host will not name the slide, which is one line after the splice
   * and is what keeps them from needing a whole deck. A case that wants a
   * finished insert sets this.
   */
  current: undefined as { index: number; id: string } | undefined,
  /** What the stubbed splice reports the destination slide already held. */
  held: 0,
  /** Every position the pane asked the host to delete, in order. */
  removed: [] as number[],
  /** True when `removeSlideAt` should refuse, which is how an undo is made to fail. */
  refuseRemoval: false,
  /** The URL the host gives for the open deck; undefined is an unsaved one. */
  url: undefined as string | undefined,
};

vi.mock("../src/pane/catalogue.js", async () => {
  const actual = await vi.importActual<typeof import("../src/pane/catalogue.js")>("../src/pane/catalogue.js");
  return {
    ...actual,
    // The real store fetches an element's markup from the site, and there is no
    // site here. Everything it would have fetched is empty: what these cases
    // are about is which options the pane hands the splice, not what it splices.
    Store: class {
      markup(): Promise<{ xml: string; rels: []; parts: [] }> {
        return Promise.resolve({ xml: "", rels: [], parts: [] });
      }
      part(): Promise<undefined> {
        return Promise.resolve(undefined);
      }
    },
    loadIndex: () =>
      indexMode === "hang"
        ? new Promise(() => undefined)
        : indexMode === "fail"
          ? Promise.reject(new Error("the site did not answer"))
          : Promise.resolve({
              version: "v1",
              sizes: {
                "16:9": {
                  size: "16:9",
                  width: 12192000,
                  height: 6858000,
                  categories: [
                    { key: "boxes", name: "White boxes" },
                    { key: "stamps", name: "Stamps and labels" },
                  ],
                  elements: [
                    {
                      id: "one-box",
                      key: "En kasse",
                      name: "One box",
                      category: { key: "boxes", name: "White boxes" },
                      slide: 1,
                      kind: "slide",
                      box: { x: 0.1, y: 0.2, w: 0.5, h: 0.5 },
                      landing: "layout",
                      shapes: 1,
                      tags: ["boxes"],
                      markup: { xml: "", rels: [], parts: [] },
                    },
                    // A PART, which is what "Remove from N slides" is offered
                    // on, and the one the removal cases insert for real.
                    {
                      id: "markeringer-1",
                      key: "Markeringer 1",
                      name: "Marker, circle",
                      category: { key: "stamps", name: "Stamps and labels" },
                      slide: 2,
                      kind: "part",
                      box: { x: 0.6, y: 0.1, w: 0.2, h: 0.2 },
                      landing: "cursor",
                      shapes: 1,
                      tags: ["stamp"],
                      markup: { xml: "", rels: [], parts: [] },
                    },
                  ],
                  carried: {},
                },
              },
            }),
  };
});

async function openPane(theme?: string): Promise<HTMLElement> {
  document.body.innerHTML = '<header><b>SSF</b><span>Slide Elements</span></header><div id="pane"></div>';
  const office = {
    onReady: (cb: () => void) => {
      cb();
      return Promise.resolve();
    },
    context: theme ? { officeTheme: { bodyBackgroundColor: theme } } : {},
  };
  vi.stubGlobal("Office", office);
  vi.resetModules();
  await import("../src/pane/main.js");
  return document.getElementById("pane") as HTMLElement;
}

/** Let the loading chain settle: `Office.onReady` starts it and does not await it. */
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

/**
 * Every listener a booted pane put on `document`, so each case starts alone.
 *
 * `openPane` imports `main.ts` fresh for each case, and the pane binds its
 * click, key and focus handlers to `document` — which survives the new body.
 * Left alone, the fifth case's click reaches five panes, four of them holding
 * the state of a test that has already finished, and the last one to redraw
 * wins. That is not a flaky test, it is a test asserting about the wrong pane.
 */
const bound: [string, EventListener][] = [];
const addToDocument = document.addEventListener.bind(document);
document.addEventListener = ((type: string, fn: EventListener, options?: AddEventListenerOptions) => {
  bound.push([type, fn]);
  addToDocument(type, fn, options);
}) as typeof document.addEventListener;

/**
 * Detach every pane opened so far, so the NEXT `openPane` starts alone.
 *
 * `afterEach` does this between cases. A case that reopens the pane several
 * times — which is what "remembers per deck" has to do — needs it between the
 * opens too: a stale pane's handlers are still on `document`, still hold the
 * previous deck's bucket, and would write the new deck's search into the old
 * deck's storage. Which is the bug this file's own comment above warns about,
 * found again the first time a case reopened the pane twice.
 */
function detachPanes(): void {
  for (const [type, fn] of bound.splice(0)) document.removeEventListener(type, fn);
}

afterEach(() => {
  detachPanes();
  vi.unstubAllGlobals();
  document.documentElement.removeAttribute("data-theme");
  readiness = { ok: true, detail: "fine" };
  indexMode = "fail";
  opened.length = 0;
  externalOpens = true;
  deckBase64 = undefined;
  host.slides = 3;
  host.cycles = 0;
  host.refuseAt = 0;
  host.namesSlides = true;
  host.supports15 = true;
  host.selected.length = 0;
  host.selectAnswers = undefined;
  host.current = undefined;
  host.held = 0;
  host.removed.length = 0;
  host.refuseRemoval = false;
  host.url = undefined;
  spliced.length = 0;
  window.localStorage.clear();
});

describe("what this deck already uses", () => {
  /** A two-slide deck with one library element stamped onto the first. */
  async function deckWithOneBox(): Promise<string> {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }]));
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const shape = doc.getElementsByTagName("p:sp")[0] as unknown as Element;
    await writeShapeTags(pkg, "ppt/slides/slide1.xml", shape, [
      [TAG_ELEMENT, "one-box"],
      [TAG_CATALOGUE, "v1"],
    ]);
    return pkg.toBase64();
  }

  /**
   * Click "See what this deck already uses" and wait for the answer.
   *
   * Polled rather than settled a fixed number of times: the read opens a zip
   * and walks every slide, which is more turns of the event loop than a
   * `setTimeout(0)` or two — and a test that guessed the number would pass or
   * fail on how fast the machine is.
   */
  async function openAndAsk(): Promise<HTMLElement> {
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="used"]') as HTMLElement).click();
    for (let i = 0; i < 200; i++) {
      if (pane.querySelector(".used-head") ?? pane.querySelector(".notice")) break;
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
    return pane;
  }

  it("does not read the deck until it is asked to", async () => {
    // The whole reason the section starts closed: reading it means reading the
    // user's entire presentation, and how long that takes on a big one is still
    // an open question in the design.
    indexMode = "ok";
    deckBase64 = await deckWithOneBox();
    const pane = await openPane();
    await settle();
    expect(pane.querySelector(".used-list")).toBeNull();
    expect(pane.querySelector('[data-action="used"]')?.textContent).toBe("See what this deck already uses");
  });

  it("names the elements it finds, and says which slides they are on", async () => {
    deckBase64 = await deckWithOneBox();
    const pane = await openAndAsk();
    const row = pane.querySelector(".used-row");
    expect(row?.querySelector(".used-name")?.textContent).toBe("One box");
    expect(row?.querySelector(".used-where")?.textContent).toBe("slide 1");
  });

  it("says the deck holds nothing rather than showing an empty space", async () => {
    // "Asked and empty" and "not asked" are different facts, and the pane has
    // to be able to tell the user which one it is.
    deckBase64 = await Pkg.open(await makeDeck([{ paragraphs: [["First"]] }])).then((p) => p.toBase64());
    const pane = await openAndAsk();
    expect(pane.querySelector(".used-head")?.textContent).toBe("Nothing from the library is in this deck yet");
  });

  it("goes to the slide when its number is clicked, and says so only once the host was seen there", async () => {
    deckBase64 = await deckWithOneBox();
    const pane = await openAndAsk();
    const go = pane.querySelector<HTMLButtonElement>('[data-action="jump"]');
    expect(go?.dataset["value"]).toBe("1");
    go?.click();
    await settle();
    // slideIdAt(0) answers "256" in this harness; that is the id the host was asked for.
    expect(host.selected).toEqual(["256"]);
    expect(document.getElementById("announcer")?.textContent).toBe("Slide 1");
    expect(pane.querySelector(".notice")).toBeNull();
  });

  it("claims nothing when the read-back names another slide", async () => {
    deckBase64 = await deckWithOneBox();
    host.selectAnswers = { supported: true, selected: ["999#1"] };
    const pane = await openAndAsk();
    pane.querySelector<HTMLButtonElement>('[data-action="jump"]')?.click();
    await settle();
    expect(pane.querySelector(".notice")?.textContent).toBe(
      "PowerPoint did not move to slide 1. Click slide 1 in the strip.",
    );
    expect(document.getElementById("announcer")?.textContent).not.toBe("Slide 1");
  });

  it("leaves the numbers as text on a host below PowerPointApi 1.5", async () => {
    deckBase64 = await deckWithOneBox();
    host.supports15 = false;
    const pane = await openAndAsk();
    expect(pane.querySelector('[data-action="jump"]')).toBeNull();
    expect(pane.querySelector(".used-where")?.textContent).toBe("slide 1");
  });

  it("opens every category when Open all is clicked, and remembers it", async () => {
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    const expanded = () =>
      [...pane.querySelectorAll('[data-action="category"]')].filter(
        (head) => head.getAttribute("aria-expanded") === "true",
      ).length;
    expect(expanded()).toBe(0);
    pane.querySelector<HTMLElement>('[data-action="open-all"]')?.click();
    await settle();
    // Both categories the stub library carries, and the control has gone.
    expect(expanded()).toBe(2);
    expect(pane.querySelector('[data-action="open-all"]')).toBeNull();
    // Written through to the browser's storage, like any other open category.
    const kept = JSON.parse(localStorage.getItem("ssf-slide-elements") ?? "{}") as { open?: string[] };
    expect(kept.open?.sort()).toEqual(["boxes", "stamps"]);
  });

  it("keeps saying it was never asked when the read fails", async () => {
    // Not an empty list: an empty list is a claim about the deck, and this is a
    // failure to look at it.
    deckBase64 = undefined;
    const pane = await openAndAsk();
    expect(pane.querySelector(".notice")?.textContent).toContain("could not be read");
    expect(pane.querySelector('[data-action="used"]')?.textContent).toBe("See what this deck already uses");
  });
});

describe("the two links out of the gear", () => {
  /** Open the pane with a library, open the gear, and click one of the links. */
  async function click(action: string): Promise<HTMLElement> {
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="gear"]') as HTMLElement).click();
    (pane.querySelector(`[data-action="${action}"]`) as HTMLElement).click();
    return pane;
  }

  it("sends the support page the build, the host and the platform — and nothing else", async () => {
    await click("report");
    // One URL, however many times it was asked for: every `openPane` in this
    // file leaves its own click listener on `document`, so a click can reach
    // several boots of the pane at once. What matters is that they all agree.
    expect(new Set(opened).size).toBe(1);
    const url = new URL(opened[0] as string);
    expect(url.pathname).toBe("/support.html");
    // The pane is served from the site, so the links go to the site it came
    // from rather than to an address written into the code.
    expect(url.origin).toBe(window.location.origin);
    expect([...url.searchParams.keys()].sort()).toEqual(["host", "platform"]);
    expect(url.searchParams.get("host")).toBe("PowerPoint");
    expect(url.searchParams.get("platform")).toBe("PC");
    // No build in a test runner: `__BUILD_STAMP__` is a define the bundler
    // replaces, and an absent one is left out rather than sent as "unknown".
  });

  it("opens the catalogue page, and closes the gear behind it", async () => {
    const pane = await click("catalogue");
    expect(opened[0]).toBe(`${window.location.origin}/catalogue.html`);
    expect(pane.querySelector(".gear-panel")).toBeNull();
  });

  it("says where to go when the host will not open a window, instead of doing nothing", async () => {
    // The failure that matters: a click that silently does nothing leaves the
    // user with no way to reach the page at all.
    externalOpens = false;
    const pane = await click("catalogue");
    const notice = pane.querySelector(".notice")?.textContent ?? "";
    expect(notice).toContain("would not open a browser window");
    expect(notice).toContain("/catalogue.html");
  });
});

describe("booting the pane", () => {
  it("draws the loading step first, with its one disabled primary", async () => {
    // The index is left HANGING on purpose: a mocked failure settles before the
    // import returns, so the loading screen would never be observable and this
    // case would silently be asserting about the problem screen instead.
    indexMode = "hang";
    const pane = await openPane();
    expect(pane.querySelector("h1")?.textContent).toBe("Loading the library");
    const button = pane.querySelector("button.primary") as HTMLButtonElement;
    expect(button.disabled).toBe(true);
  });

  it("says why when the host is below the floor, and draws no button", async () => {
    readiness = { ok: false, detail: "this PowerPoint is too old" };
    const pane = await openPane();
    expect(pane.querySelector("p.blocked")?.textContent).toBe("this PowerPoint is too old");
    expect(pane.querySelector("button")).toBeNull();
  });

  it("makes the live region once, outside the pane", async () => {
    await openPane();
    const regions = document.querySelectorAll("#announcer");
    expect(regions.length).toBe(1);
    const region = regions[0] as HTMLElement;
    expect(region.closest("#pane")).toBeNull();
    expect(region.getAttribute("role")).toBe("status");
    expect(region.getAttribute("aria-live")).toBe("polite");
  });

  it("shows the build in the header when the bundler stamped one", async () => {
    vi.stubGlobal("__BUILD_STAMP__", "abc1234");
    await openPane();
    const build = document.querySelector("header .build");
    expect(build?.textContent).toBe("abc1234");
    expect(build?.getAttribute("title")).toContain("abc1234");
  });

  it("shows no build when there is none to show", async () => {
    await openPane();
    expect(document.querySelector("header .build")).toBeNull();
  });

  it("follows PowerPoint's theme, not the browser's", async () => {
    await openPane("#1f1f1f");
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    await openPane("#ffffff");
    expect(document.documentElement.getAttribute("data-theme")).toBe("light");
  });

  it("leaves the theme alone when the host does not say", async () => {
    await openPane();
    expect(document.documentElement.getAttribute("data-theme")).toBeNull();
  });
});

describe("when the library does not arrive", () => {
  it("ends on a screen that says so and offers a retry, never a blank one", async () => {
    const pane = await openPane();
    await settle();
    expect(pane.querySelector("h1")?.textContent).toBe("The library did not load");
    const button = pane.querySelector("button.primary") as HTMLButtonElement;
    expect(button.textContent).toBe("Try again");
    expect(button.disabled).toBe(false);
    expect(pane.querySelector("p.blocked")?.textContent).toContain("the site did not answer");
  });
});

describe("when it does arrive", () => {
  it("shows the library, and admits it could not tell which slide the user is on", async () => {
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    expect(pane.querySelector("h1")?.textContent).toBe("Slide elements");
    // `currentSlide` answers undefined here, and the pane says so rather than
    // quietly aiming at the first slide.
    expect(pane.querySelector(".slide")?.textContent).toMatch(/did not say/i);
  });
});

describe("pressing a tile", () => {
  /**
   * The defect this exists for was invisible to every other test here and to
   * the shot audit, and it made the product do nothing at all.
   *
   * Most of a tile is the ghost DRAWING, which is an `<svg>`, and an SVG
   * element is not an `HTMLElement`. The click handler walked up from
   * `event.target` only when that target was an `HTMLElement`, so a click on
   * the picture — which is where a user clicks — resolved to no action and the
   * pane sat there. Every test passed because jsdom's `.click()` was being
   * called on the BUTTON, which is the one part of a tile a user is least
   * likely to hit.
   *
   * So this dispatches from the svg, the way a real pointer does. The host is
   * mocked, so the insert cannot finish — and that is the point. "The insert
   * was refused" is written in one place, inside the insert's own catch, so
   * seeing it is proof the click got that far.
   */
  it("reaches the insert when the click lands on the ghost drawing, not the button", async () => {
    indexMode = "ok";
    const pane = await openPane();
    await settle();

    pane.querySelector<HTMLElement>('[data-action="category"]')?.click();
    const ghost = pane.querySelector("svg.ghost");
    expect(ghost, "no tile drawing to click").not.toBeNull();
    expect(ghost).not.toBeInstanceOf(HTMLElement);

    ghost?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    for (let i = 0; i < 20 && !pane.querySelector(".outcome"); i++) await settle();

    expect(pane.querySelector(".outcome")?.textContent).toContain("The insert was refused");
  });

  it("still reaches it from the button itself", async () => {
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    pane.querySelector<HTMLElement>('[data-action="category"]')?.click();
    pane.querySelector<HTMLElement>('[data-action="tile"]')?.click();
    for (let i = 0; i < 20 && !pane.querySelector(".outcome"); i++) await settle();
    expect(pane.querySelector(".outcome")?.textContent).toContain("The insert was refused");
  });
});

describe("the stub keeps up with the module it stands in for", () => {
  /**
   * A double that is missing an export does not fail loudly.
   *
   * It fails as `undefined is not a function`, deep inside whichever path
   * happens to call it — and only if a test reaches that path at all. Three
   * exports were added to `src/office/powerpoint.js` and not to the stub
   * above, and every test here stayed green, because the one path that would
   * have called them returns early in a test runner.
   *
   * So the stub is held against the real module's own list of exports, read
   * off the source rather than imported: importing it would load Office.js.
   */
  it("stands in for every function the real module exports", async () => {
    const source = readFileSync("src/office/powerpoint.ts", "utf8");
    const exported = [...source.matchAll(/^export (?:async )?function (\w+)/gm)].map((m) => m[1] ?? "");
    expect(exported.length, "no exports found — the pattern has stopped matching").toBeGreaterThan(5);
    const stub = await import("../src/office/powerpoint.js");
    for (const name of exported) {
      expect(Object.keys(stub), `the stub has no ${name}`).toContain(name);
    }
  });
});

describe("the other insert target, on right-click", () => {
  async function openWithTiles(): Promise<HTMLElement> {
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    // The library's one category starts collapsed; open it so there is a tile.
    (pane.querySelector('[data-action="category"]') as HTMLElement).click();
    return pane;
  }

  const rightClick = (node: HTMLElement): boolean =>
    !node.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));

  it("opens on a tile, and takes the browser's own menu with it", async () => {
    const pane = await openWithTiles();
    const tile = pane.querySelector('[data-action="tile"]') as HTMLElement;
    // The event is CANCELLED — a browser menu over our own would be two menus.
    expect(rightClick(tile)).toBe(true);
    expect(pane.querySelector('[data-action="other-target"]')?.textContent).toBe("Insert as a new slide");
  });

  it("leaves the browser's menu alone everywhere else, so paste still works", async () => {
    const pane = await openWithTiles();
    const search = pane.querySelector('[data-action="search"]') as HTMLElement;
    expect(rightClick(search)).toBe(false);
    expect(pane.querySelector('[data-action="other-target"]')).toBeNull();
  });

  it("closes on Escape, before the search is cleared", async () => {
    // The Escape chain backs out of what was opened last. A menu that took the
    // search with it would cost the user their query for a menu they opened by
    // accident.
    const pane = await openWithTiles();
    const search = pane.querySelector('[data-action="search"]') as HTMLInputElement;
    // A query the fixture's one element still matches, so there is a tile to
    // right-click after the search narrows the list.
    search.value = "box";
    search.dispatchEvent(new Event("input", { bubbles: true }));
    rightClick(pane.querySelector('[data-action="tile"]') as HTMLElement);
    expect(pane.querySelector('[data-action="other-target"]')).not.toBeNull();

    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(pane.querySelector('[data-action="other-target"]')).toBeNull();
    expect((pane.querySelector('[data-action="search"]') as HTMLInputElement).value).toBe("box");
  });

  it("closes when the next click lands anywhere else", async () => {
    const pane = await openWithTiles();
    rightClick(pane.querySelector('[data-action="tile"]') as HTMLElement);
    (pane.querySelector('[data-action="gear"]') as HTMLElement).click();
    expect(pane.querySelector('[data-action="other-target"]')).toBeNull();
  });
});

describe("what the menu actually inserts", () => {
  async function tileAndMenu(): Promise<HTMLElement> {
    indexMode = "ok";
    deckBase64 = await Pkg.open(await makeDeck([{ paragraphs: [["First"]] }])).then((p) => p.toBase64());
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="category"]') as HTMLElement).click();
    return pane;
  }

  /** Let the insert run as far as it can with no host to talk to. */
  async function ran(): Promise<void> {
    for (let i = 0; i < 200 && spliced.length === 0; i++) await new Promise((r) => setTimeout(r, 5));
  }

  it("uses the OTHER target for that one insert, and leaves the setting alone", async () => {
    const pane = await tileAndMenu();
    const tile = pane.querySelector('[data-action="tile"]') as HTMLElement;
    tile.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    (pane.querySelector('[data-action="other-target"]') as HTMLElement).click();
    await ran();
    expect(spliced.map((s) => s.target)).toEqual(["new"]);
    // The gear still says what it said: this was one insert, not a setting.
    (pane.querySelector('[data-action="gear"]') as HTMLElement).click();
    expect(pane.querySelector('[data-action="target"][data-value="onto"]')?.getAttribute("aria-pressed")).toBe("true");
  });

  it("uses the setting for an ordinary click on the same tile", async () => {
    const pane = await tileAndMenu();
    (pane.querySelector('[data-action="tile"]') as HTMLElement).click();
    await ran();
    expect(spliced.map((s) => s.target)).toEqual(["onto"]);
  });
});

describe("what the pane remembers, and where", () => {
  /**
   * `docs/DESIGN.md` section 4. Two buckets: favourites and the first-run flag
   * per machine, everything about how you were reading the library per DECK.
   * Driven through the real `localStorage` jsdom gives, because what is under
   * test is which key each half lands under.
   */
  const A = "https://contoso-my.sharepoint.com/personal/me/Documents/Q4.pptx";
  const B = "https://contoso-my.sharepoint.com/personal/me/Documents/Q3.pptx";

  /** Open the pane on a deck and answer what the search box came back holding. */
  async function reopen(url: string | undefined): Promise<{ pane: HTMLElement; query: string }> {
    detachPanes();
    host.url = url;
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    return { pane, query: pane.querySelector<HTMLInputElement>('[data-action="search"]')?.value ?? "" };
  }

  /** Type into the open pane's search box. */
  async function search(pane: HTMLElement, query: string): Promise<void> {
    const box = pane.querySelector<HTMLInputElement>('[data-action="search"]') as HTMLInputElement;
    box.value = query;
    box.dispatchEvent(new Event("input", { bubbles: true }));
    await settle();
  }

  it("gives each deck its own search back, and gives a new deck none", async () => {
    // Section 4 asks for the search back, which nothing did before 2026-09-12.
    await search((await reopen(A)).pane, "boxes");
    expect((await reopen(A)).query).toBe("boxes");
    // A different deck is a different question, so it opens with a clear box.
    expect((await reopen(B)).query).toBe("");
    await search((await reopen(B)).pane, "stamp");
    // And neither has taken the other's.
    expect((await reopen(A)).query).toBe("boxes");
    expect((await reopen(B)).query).toBe("stamp");
  });

  it("gives each deck its own tags back", async () => {
    detachPanes();
    host.url = A;
    indexMode = "ok";
    let pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="tag"]') as HTMLElement).click();
    await settle();
    const picked = (p: HTMLElement): number => p.querySelectorAll('[data-action="tag"][aria-pressed="true"]').length;
    expect(picked(pane)).toBe(1);

    detachPanes();
    host.url = B;
    pane = await openPane();
    await settle();
    expect(picked(pane)).toBe(0);

    detachPanes();
    host.url = A;
    pane = await openPane();
    await settle();
    expect(picked(pane)).toBe(1);
  });

  it("keeps one deck's open categories out of another's", async () => {
    detachPanes();
    host.url = A;
    indexMode = "ok";
    let pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="open-all"]') as HTMLElement).click();
    await settle();

    // A different deck opens closed, whatever the first one did.
    detachPanes();
    host.url = B;
    pane = await openPane();
    await settle();
    const openIn = (p: HTMLElement): number =>
      [...p.querySelectorAll('[data-action="category"]')].filter((h) => h.getAttribute("aria-expanded") === "true")
        .length;
    expect(openIn(pane)).toBe(0);

    // And the first deck still has its own.
    detachPanes();
    host.url = A;
    pane = await openPane();
    await settle();
    expect(openIn(pane)).toBe(2);
  });

  it("keeps favourites and the first-run flag per machine, across decks", async () => {
    detachPanes();
    host.url = A;
    indexMode = "ok";
    let pane = await openPane();
    await settle();
    // Dismiss the coach marks and star one element, both on deck A.
    (pane.querySelector('[data-action="coached"]') as HTMLElement).click();
    await settle();
    (pane.querySelector('[data-action="category"]') as HTMLElement).click();
    (pane.querySelector('[data-action="star"]') as HTMLElement).click();
    await settle();

    detachPanes();
    host.url = B;
    pane = await openPane();
    await settle();
    // A statement about the library and about the person, not about a deck.
    expect(pane.querySelector(".coach")).toBeNull();
    expect(pane.textContent).toContain("Favourites");
  });

  it("writes the deck's half under a key of its own, and never the URL", async () => {
    host.url = A;
    indexMode = "ok";
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="open-all"]') as HTMLElement).click();
    await settle();
    const keys = Object.keys(localStorage);
    expect(keys).toContain("ssf-slide-elements");
    const deckKeys = keys.filter((k) => k !== "ssf-slide-elements");
    expect(deckKeys).toHaveLength(1);
    // A SharePoint path can name a client, a project or a person.
    expect(deckKeys[0]).toMatch(/^ssf-slide-elements:[0-9a-f]{8}$/);
    expect(localStorage.getItem(deckKeys[0] ?? "")).not.toContain("sharepoint");
    // The per-machine half holds no browsing state, and the deck's holds no
    // favourites: the split is real rather than two copies of everything.
    const machine = JSON.parse(localStorage.getItem("ssf-slide-elements") ?? "{}") as Record<string, unknown>;
    expect(Object.keys(machine).sort()).toEqual(["coached", "favourites"]);
    const deck = JSON.parse(localStorage.getItem(deckKeys[0] ?? "") ?? "{}") as { open?: string[] };
    expect(deck.open?.sort()).toEqual(["boxes", "stamps"]);
  });

  it("falls back to the one bucket when the host will not name a deck, and still remembers", async () => {
    // An unsaved deck. Everything in one place rather than nothing remembered.
    detachPanes();
    host.url = undefined;
    indexMode = "ok";
    let pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="open-all"]') as HTMLElement).click();
    await settle();
    expect(Object.keys(localStorage)).toEqual(["ssf-slide-elements"]);

    detachPanes();
    pane = await openPane();
    await settle();
    expect(
      [...pane.querySelectorAll('[data-action="category"]')].filter((h) => h.getAttribute("aria-expanded") === "true")
        .length,
    ).toBe(2);
  });
});

describe("moving the last insert to a new slide", () => {
  /**
   * A pane that has just finished an insert onto slide 1.
   *
   * Unlike the cases above, this one lets the insert RUN OUT: `host.current`
   * names the slide, so the pane gets past the refusal every other case in this
   * file stops at, and ends with a footer. `held` is what the stubbed splice
   * reports the destination already carried, which is the whole of the pane's
   * decision (`docs/DESIGN.md` section 6).
   */
  async function inserted(held: number, id = "one-box"): Promise<HTMLElement> {
    indexMode = "ok";
    host.current = { index: 0, id: "256" };
    host.held = held;
    deckBase64 = await Pkg.open(await makeDeck([{ paragraphs: [["First"]] }])).then((p) => p.toBase64());
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="category"]') as HTMLElement).click();
    (pane.querySelector(`[data-tile="${id}"], [data-action="tile"]`) as HTMLElement).click();
    for (let i = 0; i < 200 && spliced.length === 0; i++) await new Promise((r) => setTimeout(r, 5));
    await settle();
    return pane;
  }

  it("offers the move only when the slide already held something", async () => {
    expect((await inserted(1)).querySelector('[data-action="move"]')).not.toBeNull();
    expect((await inserted(0)).querySelector('[data-action="move"]')).toBeNull();
  });

  it("takes the insert back and makes it again as a new slide", async () => {
    const pane = await inserted(1);
    expect(spliced.map((s) => s.target)).toEqual(["onto"]);
    pane.querySelector<HTMLElement>('[data-action="move"]')?.click();
    for (let i = 0; i < 200 && spliced.length < 2; i++) await new Promise((r) => setTimeout(r, 5));
    await settle();
    // The second splice is the same element with the other target — an undo
    // and a fresh insert, not a second copy beside the first.
    expect(spliced.map((s) => s.target)).toEqual(["onto", "new"]);
    // And the undo really ran: a positional delete happened between the two.
    expect(host.removed.length).toBeGreaterThan(1);
  });

  it("stops at the undo when the undo does not work, rather than inserting a second copy", async () => {
    // The one outcome a user asking to MOVE something cannot have meant.
    const pane = await inserted(1);
    const before = spliced.length;
    host.refuseRemoval = true;
    pane.querySelector<HTMLElement>('[data-action="move"]')?.click();
    await settle();
    await settle();
    expect(spliced.length).toBe(before);
    expect(pane.querySelector(".outcome")?.textContent).toContain("Undo did not work");
  });

  it("withdraws the offer once the insert has simply been undone", async () => {
    // The offer goes because the history it needs went, which is `footerOf`'s
    // `undo > 0` — nothing clears `moveable`, deliberately. This is the
    // end-to-end half of the steps case that says the same thing.
    const pane = await inserted(1);
    pane.querySelector<HTMLElement>('[data-action="undo"]')?.click();
    for (let i = 0; i < 200 && host.removed.length < 2; i++) await new Promise((r) => setTimeout(r, 5));
    await settle();
    expect(pane.querySelector('[data-action="move"]')).toBeNull();
  });
});

describe("removing a part from every slide it is on", () => {
  /**
   * A deck with the library's one stamp on two slides, read by the pane.
   *
   * Built by INSERTING, through the real splice — this describe unmocks it,
   * because what is under test is a sequence of real removals and a stubbed
   * splice would leave nothing to remove.
   */
  async function deckWithStampOn(slides: number[]): Promise<string> {
    const { splice: realSplice } =
      await vi.importActual<typeof import("../src/core/splice/splice.js")>("../src/core/splice/splice.js");
    const { Pkg: RealPkg, harvest } =
      await vi.importActual<typeof import("../src/core/index.js")>("../src/core/index.js");
    const names = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Parameters<typeof harvest>[1]["names"];
    const lib = await harvest(await RealPkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx"))), {
      size: "16:9",
      names,
    });
    const el = lib.catalogue.elements.find((e) => e.id === "markeringer-1");
    if (!el) throw new Error("the library has no markeringer-1");
    let deck: string | Uint8Array = await makeDeck([
      { paragraphs: [["First"]] },
      { paragraphs: [["Second"]] },
      { paragraphs: [["Third"]] },
    ]);
    for (const slide of slides) {
      const report = await realSplice({
        deck,
        slide,
        element: {
          id: el.id,
          name: el.name,
          kind: el.kind,
          box: el.box,
          landing: el.landing,
          markup: el.markup,
        },
        options: { target: "onto", group: true, colours: "deck" },
        catalogue: { version: "v1", carried: lib.catalogue.carried, theme: lib.catalogue.theme },
        store: (path: string) => Promise.resolve(lib.parts.get(path)),
      });
      deck = report.base64;
    }
    return deck as string;
  }

  async function askedToRemove(): Promise<HTMLElement> {
    indexMode = "ok";
    deckBase64 = await deckWithStampOn([0]);
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="used"]') as HTMLElement).click();
    for (let i = 0; i < 300; i++) {
      if (pane.querySelector(".used-list")) break;
      await new Promise((r) => setTimeout(r, 5));
    }
    // The part's category, so its tile is on screen to carry the button.
    const stamps = [...pane.querySelectorAll<HTMLElement>('[data-action="category"]')].find(
      (c) => c.dataset["key"] === "stamps",
    );
    stamps?.click();
    (pane.querySelector('[data-action="remove"]') as HTMLElement).click();
    return pane;
  }

  /** Wait for a run to finish: the footer is what says it did. */
  async function ran(pane: HTMLElement): Promise<string> {
    for (let i = 0; i < 300; i++) {
      const said = pane.querySelector(".outcome")?.textContent;
      if (said) return said;
      await new Promise((r) => setTimeout(r, 5));
    }
    return "";
  }

  it("asks first, and removes nothing while the question is up", async () => {
    const pane = await askedToRemove();
    expect(pane.querySelector(".tile-ask")?.textContent).toContain("cannot undo");
    expect(host.cycles).toBe(0);
  });

  it("does nothing at all if the question is answered no", async () => {
    const pane = await askedToRemove();
    (pane.querySelector('[data-action="remove-cancel"]') as HTMLElement).click();
    expect(pane.querySelector(".tile-ask")).toBeNull();
    expect(host.cycles).toBe(0);
  });

  it("runs one insert-and-remove cycle per slide once it is answered yes", async () => {
    const pane = await askedToRemove();
    (pane.querySelector('[data-action="remove-go"]') as HTMLElement).click();
    expect(await ran(pane)).toContain("Removed from 1 slide");
    expect(host.cycles).toBe(1);
  });

  it("stops at the first step the deck's own size does not confirm", async () => {
    // The rule the insert path is built on: the DELTA is the evidence, and a
    // loop that pressed on past a step it could not verify would be editing a
    // deck whose shape it has already misread.
    host.refuseAt = 1;
    const pane = await askedToRemove();
    (pane.querySelector('[data-action="remove-go"]') as HTMLElement).click();
    const said = await ran(pane);
    expect(said).toContain("Removed from 0 of 1 slides");
    expect(said).toContain("as they were");
    expect(pane.querySelector(".outcome.by-hand")).not.toBeNull();
  });

  it("forgets what it knew about the deck afterwards, rather than showing a stale list", async () => {
    const pane = await askedToRemove();
    (pane.querySelector('[data-action="remove-go"]') as HTMLElement).click();
    await ran(pane);
    // Back to "never asked": the deck has changed under the pane, and the list
    // it read is about the deck as it was.
    expect(pane.querySelector('[data-action="used"]')?.textContent).toBe("See what this deck already uses");
  });
});

describe("a removal asks the deck it is about to change", () => {
  /** A deck of three slides carrying the part on the ones named, 1-based. */
  async function deckTagging(slides: number[]): Promise<string> {
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }, { paragraphs: [["Third"]] }]),
    );
    for (const n of slides) {
      const path = `ppt/slides/slide${n}.xml`;
      const doc = await pkg.doc(path);
      const shape = doc.getElementsByTagName("p:sp")[0] as unknown as Element;
      await writeShapeTags(pkg, path, shape, [
        [TAG_ELEMENT, "markeringer-1"],
        [TAG_CATALOGUE, "v1"],
      ]);
    }
    return pkg.toBase64();
  }

  it("re-derives which slides carry it, instead of trusting the list the question was asked about", async () => {
    // The list behind "Remove from N slides" comes from an EARLIER read. In
    // between, the user can add a slide, delete one, or drag them around — the
    // list is 1-based positions, and a position that has moved names a
    // different slide. Here the pane reads a deck carrying the part on two
    // slides, and the deck it is then handed carries it on one.
    indexMode = "ok";
    deckBase64 = await deckTagging([1, 2]);
    const pane = await openPane();
    await settle();
    (pane.querySelector('[data-action="used"]') as HTMLElement).click();
    for (let i = 0; i < 300; i++) {
      if (pane.querySelector(".used-list")) break;
      await new Promise((r) => setTimeout(r, 5));
    }
    expect(pane.querySelector(".used-where")?.textContent).toBe("slides 1 and 2");

    // The deck moves on under the pane.
    deckBase64 = await deckTagging([1]);

    const stamps = [...pane.querySelectorAll<HTMLElement>('[data-action="category"]')].find(
      (c) => c.dataset["key"] === "stamps",
    );
    stamps?.click();
    expect(pane.querySelector('[data-action="remove"]')?.textContent).toBe("Remove from 2 slides");
    (pane.querySelector('[data-action="remove"]') as HTMLElement).click();
    (pane.querySelector('[data-action="remove-go"]') as HTMLElement).click();
    let said = "";
    for (let i = 0; i < 300; i++) {
      said = pane.querySelector(".outcome")?.textContent ?? "";
      if (said) break;
      await new Promise((r) => setTimeout(r, 5));
    }
    // One slide carries it now, so one cycle runs and the footer counts what
    // the DECK said. Trusting the list gives "Removed from 1 of 2 slides" and
    // an attempt against a slide that holds nothing of ours.
    expect(said).toContain("Removed from 1 slide");
    expect(said).not.toContain("of 2");
    expect(host.cycles).toBe(1);
  });
});
