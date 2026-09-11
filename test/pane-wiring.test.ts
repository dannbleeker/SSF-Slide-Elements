/**
 * @vitest-environment jsdom
 */
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";

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
  // Every call the pane can make, so a missing export cannot be mistaken for a
  // pane that decided not to make it. The last case in this file holds this
  // list against the real module, because the claim above was untrue for three
  // exports before anyone noticed.
  slideCount: () => Promise.resolve(3),
  countReaching: () => Promise.resolve(3),
  readDeck: () => Promise.reject(new Error("no deck in a test runner")),
  currentSlide: () => Promise.resolve(undefined),
  selectedShape: () => Promise.resolve(undefined),
  slideIdAt: () => Promise.resolve(undefined),
  onSlideChange: () => Promise.resolve(false),
  insertPackage: () => Promise.resolve(undefined),
  removeSlideAt: () => Promise.resolve(undefined),
}));

vi.mock("../src/pane/catalogue.js", async () => {
  const actual = await vi.importActual<typeof import("../src/pane/catalogue.js")>("../src/pane/catalogue.js");
  return {
    ...actual,
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
                  categories: [{ key: "boxes", name: "White boxes" }],
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

afterEach(() => {
  vi.unstubAllGlobals();
  document.documentElement.removeAttribute("data-theme");
  readiness = { ok: true, detail: "fine" };
  indexMode = "fail";
  window.localStorage.clear();
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
