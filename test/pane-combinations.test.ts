/**
 * @vitest-environment jsdom
 */
/**
 * The pane's state, combined.
 *
 * Every other pane test picks a state a person would be in. This one picks
 * states nobody would think to: the gear open while a removal is being
 * confirmed, a card open on a busy pane with a stale snapshot, an element in
 * Favourites AND Recent AND its category at once. It renders each and asks the
 * three things that must be true of ANY of them — it does not throw, there is
 * one primary control, and nothing that belongs to one tile is drawn on two.
 *
 * That last one is why this file exists. Swept over all 262,144 combinations of
 * eighteen flags, it found the right-click menu being drawn on every copy of an
 * element rather than on the tile that was clicked — a state no hand-written
 * case had put together, because it needs Recent to hold the element the menu
 * is open on.
 *
 * Kept at singles, pairs and triples rather than the full power set: that is
 * 1,000 renders instead of 262,144, it runs in a blink, and the defect it was
 * written for is a PAIR (a menu plus a list holding the same element). The full
 * sweep is a scratch file away when something smells.
 */
import { describe, expect, it } from "vitest";
import type { Element } from "../src/core/catalogue/types.js";
import { render } from "../src/pane/render.js";
import { EMPTY, stepFor, type Library, type PaneState } from "../src/pane/steps.js";

function element(over: Partial<Element> & { id: string }): Element {
  return {
    key: over.key ?? over.id,
    name: over.name ?? over.id,
    category: over.category ?? { key: "boxes", name: "White boxes" },
    slide: 1,
    kind: "slide",
    box: { x: 0.1, y: 0.2, w: 0.5, h: 0.5 },
    landing: "layout",
    shapes: 1,
    tags: ["boxes"],
    markup: { xml: "", rels: [], parts: [] },
    ...over,
  };
}

const stamps = { key: "stamps", name: "Stamps and labels" };
const LIBRARY: Library = {
  size: "16:9",
  width: 12192000,
  height: 6858000,
  version: "v1",
  categories: [{ key: "boxes", name: "White boxes" }, stamps],
  elements: [
    element({ id: "one-box", name: "One box" }),
    element({ id: "flow-1", name: "Flow, 1 box", run: { key: "Flow, N boxes", noun: "boxes", count: 1 } }),
    element({ id: "flow-2", name: "Flow, 2 boxes", run: { key: "Flow, N boxes", noun: "boxes", count: 2 } }),
    element({ id: "approved", name: "Approved stamp", kind: "part", landing: "top-right", category: stamps }),
  ],
};

/** Each flag is a thing that can be on or off independently of the others. */
const FLAGS = [
  ["library", { library: LIBRARY }],
  ["problem", { problem: "offline" }],
  ["busy", { busy: true, chosen: "one-box" }],
  ["gear", { gear: true }],
  ["coached", { coached: true }],
  ["query", { query: "box", open: ["boxes", "stamps"] }],
  ["tags", { tags: ["boxes"] }],
  ["category", { category: "boxes" }],
  ["previewing", { previewing: "one-box", open: ["boxes"] }],
  ["menuFor", { menuFor: "one-box", open: ["boxes"] }],
  ["removing", { removing: { id: "approved", slides: [2, 5], done: 0 }, open: ["stamps"] }],
  [
    "used",
    {
      used: [
        { element: "approved", slides: [2, 5] },
        { element: "gone", slides: [1] },
      ],
    },
  ],
  ["reading", { reading: true }],
  ["onSlide", { slide: 2, onSlide: { slide: 2, boxes: [{ x: 0.1, y: 0.1, w: 0.3, h: 0.2 }] } }],
  ["deck", { deck: { width: 10692000, height: 7560000 } }],
  [
    "outcome",
    { outcome: { ok: false, byHand: true, name: "One box", detail: "did not work" }, undo: 1, recent: ["one-box"] },
  ],
  ["notice", { notice: "something happened" }],
  // The flag the duplicated menu needed: one element in three lists at once.
  ["everywhere", { favourites: ["one-box", "approved"], recent: ["one-box"], open: ["boxes", "stamps"] }],
  ["borrowed", { library: { ...LIBRARY, borrowed: "4:3 library, scaled to A4 slides." } }],
] as const;

describe("fuzz: every combination of the pane's optional state", () => {
  /** Every mask with at most three flags on: the shapes a pair-wise defect hides in. */
  function masks(upTo: number): number[] {
    const out: number[] = [];
    const bits = FLAGS.length;
    for (let mask = 0; mask < 1 << bits; mask += 1) {
      let on = 0;
      for (let bit = 0; bit < bits; bit += 1) if (mask & (1 << bit)) on += 1;
      if (on <= upTo) out.push(mask);
    }
    return out;
  }

  it("renders any of them without throwing, with one primary and nothing drawn twice", () => {
    const root = document.createElement("div");
    document.body.append(root);
    const broken: string[] = [];
    const all = masks(3);
    for (const mask of all) {
      let state: PaneState = { ...EMPTY };
      const on: string[] = [];
      for (let bit = 0; bit < FLAGS.length; bit += 1) {
        if (!(mask & (1 << bit))) continue;
        const [name, patch] = FLAGS[bit] as [string, Partial<PaneState>];
        on.push(name);
        state = { ...state, ...patch };
      }
      try {
        render(root, state, stepFor(state));
        const primaries = root.querySelectorAll("button.primary").length;
        if (primaries > 1) broken.push(`${on.join("+")}: ${primaries} primary buttons`);
        const cards = root.querySelectorAll(".card").length;
        if (cards > 1) broken.push(`${on.join("+")}: ${cards} preview cards`);
        const menus = root.querySelectorAll(".tile-menu").length;
        if (menus > 1) broken.push(`${on.join("+")}: ${menus} menus at once`);
      } catch (e) {
        broken.push(`${on.join("+")}: threw ${e instanceof Error ? e.message : String(e)}`);
      }
      if (broken.length > 8) break;
    }
    // The vacuity guard: a sample that shrank to nothing would pass forever.
    expect(all.length).toBeGreaterThan(900);
    expect(broken.slice(0, 8)).toEqual([]);
  }, 300000);
});
