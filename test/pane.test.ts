/**
 * @vitest-environment jsdom
 *
 * The picker's behaviour, in the one environment the suite has for it.
 *
 * jsdom has no layout and no colour, so nothing here can say the pane LOOKS
 * right — that is what `scripts/pane-shots.mjs` is for. What it can pin is
 * which elements a search finds, what an empty result says, and whether a
 * section with nothing under it is drawn at all.
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { clampName, countLine, esc, fold, grouped, matches, renderPicker, statusLine } from "../src/pane/render.js";
import type { CatalogueIndex } from "../src/core/catalogue/types.js";

const index = JSON.parse(
  readFileSync(join(process.cwd(), "public", "catalogue", "index.json"), "utf8"),
) as CatalogueIndex;

describe("search", () => {
  it("shows everything for an empty query", () => {
    expect(matches(index, "").length).toBe(index.elements.length);
  });

  /**
   * The library's names assume the section is visible: "2 vertikale" exists in
   * several sections, and somebody searching for "streger" means the ones under
   * "Kun streger" even though those names never say so.
   */
  it("matches a section's name as well as an element's", () => {
    const hits = matches(index, "streger");
    expect(hits.length).toBeGreaterThan(0);
    const sections = new Set(hits.map((h) => h.section));
    expect(sections.has("kun-streger")).toBe(true);
  });

  it("finds a Danish name typed without its accents", () => {
    expect(matches(index, "gra kasser").length).toBeGreaterThan(0);
    expect(fold("Grå")).toBe("graa");
  });

  it("narrows on each extra term rather than widening", () => {
    const one = matches(index, "hvid").length;
    const two = matches(index, "hvid checkliste").length;
    expect(two).toBeGreaterThan(0);
    expect(two).toBeLessThan(one);
  });

  it("answers nothing for a query that matches nothing", () => {
    expect(matches(index, "zzzzzz")).toEqual([]);
  });
});

describe("grouping", () => {
  it("drops a section whose elements were all filtered out", () => {
    const shown = matches(index, "trekant");
    const groups = grouped(index, shown);
    // A section header with nothing under it reads as one that failed to load.
    expect(groups.every((g) => g.items.length > 0)).toBe(true);
    expect(groups.length).toBeLessThan(index.sections.length);
  });

  it("keeps the deck's own section order", () => {
    const names = grouped(index, index.elements).map((g) => g.name);
    expect(names).toEqual(index.sections.map((s) => s.name));
  });
});

describe("the count line", () => {
  it("says the total when nothing is filtered", () => {
    expect(countLine(98, 98)).toBe("98 elements");
  });
  it("says the fraction when something is", () => {
    expect(countLine(4, 98)).toBe("4 of 98");
  });
  it("says nothing matches rather than 0 of 98", () => {
    expect(countLine(0, 98)).toBe("nothing matches");
  });
  it("does not pluralise one", () => {
    expect(countLine(1, 1)).toBe("1 element");
  });
});

describe("rendering", () => {
  it("draws a card per element, carrying its id", () => {
    document.body.innerHTML = renderPicker(index, "trekant");
    const cards = document.querySelectorAll(".card");
    expect(cards.length).toBe(matches(index, "trekant").length);
    expect(cards[0]?.getAttribute("data-id")).toBeTruthy();
  });

  it("puts the preview inside the card", () => {
    document.body.innerHTML = renderPicker(index, "trekant");
    expect(document.querySelector(".card .thumb svg")).not.toBeNull();
  });

  it("says so when nothing matches, quoting what was typed", () => {
    document.body.innerHTML = renderPicker(index, "zzzzzz");
    expect(document.querySelector(".empty")?.textContent).toContain("zzzzzz");
    expect(document.querySelectorAll(".card").length).toBe(0);
  });

  /**
   * Every name on this screen came out of a file. None of the library's do
   * anything interesting, but the escape is the thing being tested, not the
   * library.
   */
  it("escapes a name that carries markup", () => {
    expect(esc('<img src=x onerror="1">')).toBe("&lt;img src=x onerror=&quot;1&quot;&gt;");
  });

  it("keeps what was typed in the search box after a redraw", () => {
    document.body.innerHTML = renderPicker(index, "hvid");
    expect((document.getElementById("q") as HTMLInputElement).value).toBe("hvid");
  });
});

describe("label truncation", () => {
  /**
   * Done in the string rather than in CSS. `-webkit-line-clamp` computed to an
   * inert `flow-root` here and hard-clipped the text mid-line instead, and the
   * pane runs in a WebView whose engine version this repo does not choose. A
   * character budget is the same everywhere and is visible to jsdom, where a
   * CSS clamp is not.
   */
  it("leaves a short name exactly as it is", () => {
    expect(clampName("En kasse")).toBe("En kasse");
  });

  it("cuts a long name at a word boundary and marks it", () => {
    const long = "Tabel med store hvid kasse på mørk baggrund – 5 rækker, 7 kolonner";
    const out = clampName(long);
    expect(out.length).toBeLessThanOrEqual(47);
    expect(out.endsWith("…")).toBe(true);
    expect(out).not.toMatch(/ …$/);
    expect(long.startsWith(out.slice(0, -1))).toBe(true);
  });

  it("cuts mid-word rather than returning almost nothing", () => {
    // No space in the budget at all: the alternative is an ellipsis and
    // three characters, which names nothing.
    const out = clampName("Averyverylongsinglewordwithnospacesatallinit", 20);
    expect(out.length).toBe(21);
  });

  it("keeps the full name on the card for hover", () => {
    document.body.innerHTML = renderPicker(index, "");
    const card = Array.from(document.querySelectorAll(".card")).find(
      (c) => (c.getAttribute("title") ?? "").length > 46,
    );
    expect(card).toBeDefined();
    expect(card!.getAttribute("title")).not.toContain("…");
  });
});

describe("status", () => {
  it("names the element it is working on", () => {
    expect(statusLine({ phase: "splicing", name: "Trekant" })).toContain("Trekant");
    expect(statusLine({ phase: "done", name: "Trekant" })).toContain("Trekant");
  });
  it("says nothing when idle", () => {
    expect(statusLine({ phase: "idle" })).toBe("");
  });
  it("passes a failure's own detail through", () => {
    expect(statusLine({ phase: "failed", detail: "no room" })).toBe("no room");
  });
});
