/**
 * @vitest-environment jsdom
 */
import { afterEach, describe, expect, it, vi } from "vitest";

/**
 * The pane's entry point, wired.
 *
 * `pane-render.test.ts` calls `render` directly, so `main.ts` — the only file
 * that touches Office.js — would otherwise run nowhere in the suite. It runs
 * here with `Office` stubbed to the two things it reads: `onReady`, which calls
 * back at once, and `context.officeTheme`, which the theme stamp reads.
 */
type Readiness = { ok: boolean; detail: string };
let readiness: Readiness = { ok: true, detail: "fine" };

vi.mock("../src/office/powerpoint.js", () => ({
  ready: () => readiness,
}));

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

afterEach(() => {
  vi.unstubAllGlobals();
  document.documentElement.removeAttribute("data-theme");
  readiness = { ok: true, detail: "fine" };
});

describe("booting the pane", () => {
  it("draws the start step with its one disabled primary", async () => {
    const pane = await openPane();
    expect(pane.querySelector("h1")?.textContent).toBe("Start here");
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
