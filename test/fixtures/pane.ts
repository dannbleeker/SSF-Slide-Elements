/**
 * The library the pane's decision tests are all asked about.
 *
 * Five elements: two plain boxes, a two-member sized run, and a stamp that is a
 * PART rather than a whole slide — which is the distinction most of the pane's
 * rules turn on. Small on purpose. A rule that needs a hundred elements to show
 * itself is a rule nobody can read the test for.
 *
 * Shared rather than copied, because `pane-steps`, `pane-search`, `pane-card`
 * and `pane-used` all ask about the same library and four copies of it would
 * drift. It followed the split of `steps.ts` out of one file on 2026-09-12.
 */
import type { Element } from "../../src/core/catalogue/types.js";
import { EMPTY, type Library, type PaneState } from "../../src/pane/steps.js";

export function element(over: Partial<Element> & { id: string }): Element {
  return {
    key: over.key ?? over.id,
    name: over.name ?? over.id,
    category: over.category ?? { key: "boxes", name: "White boxes" },
    slide: over.slide ?? 1,
    kind: over.kind ?? "slide",
    box: over.box ?? { x: 0.1, y: 0.2, w: 0.5, h: 0.5 },
    landing: over.landing ?? "layout",
    shapes: over.shapes ?? 1,
    tags: over.tags ?? [],
    markup: over.markup ?? { xml: "", rels: [], parts: [] },
    ...over,
  };
}

export const LIBRARY: Library = {
  size: "16:9",
  width: 12192000,
  height: 6858000,
  version: "v1",
  categories: [
    { key: "boxes", name: "White boxes" },
    { key: "stamps", name: "Stamps and labels" },
  ],
  elements: [
    element({ id: "one-box", name: "One box", tags: ["boxes", "white"] }),
    element({ id: "two-boxes", name: "Two boxes", tags: ["boxes"] }),
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
    element({
      id: "approved",
      name: "Approved stamp",
      key: "Godkendt",
      category: { key: "stamps", name: "Stamps and labels" },
      landing: "top-right",
      kind: "part",
      tags: ["stamp"],
    }),
  ],
};

export const browsing: PaneState = { ...EMPTY, library: LIBRARY };
