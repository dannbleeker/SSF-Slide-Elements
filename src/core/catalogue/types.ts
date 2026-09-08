/**
 * The catalogue: what the harvest reads out of a library deck, and what the
 * pane and the splice read back.
 *
 * Everything here is data. The rules that produce it are in `harvest.ts`,
 * the boxes in `boxes.ts`, the sizes in `runs.ts`, the tags in `tags.ts`;
 * `docs/DESIGN.md` sections 2 and 3 are the prose they implement.
 */

/** The two slide sizes the library is authored in. */
export type SlideSize = "16:9" | "4:3";

/** A box in fractions of the slide: 0..1 on both axes. */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Where an element lands on the destination slide (`docs/DESIGN.md` section 5).
 *
 * - `layout`: a whole-slide element, placed relative to the destination's
 *   title placeholder and body.
 * - `top-right`: a stamp or a label.
 * - `cursor`: a marker, a flowchart shape or an icon, at the selected shape or
 *   the slide centre.
 * - `as-authored`: a part wider than half the slide, left where it sits.
 */
export type Landing = "layout" | "top-right" | "cursor" | "as-authored";

/** An element that comes in several sizes: which run it belongs to and its count. */
export interface SizeRun {
  /** The English name with the count replaced by `N`, shared by every member. */
  key: string;
  /** What the count counts: boxes, rows, lines, levels, columns, areas, items. */
  noun: string;
  count: number;
}

/** One relationship an element's markup names, resolved to a package path. */
export interface MarkupRel {
  id: string;
  type: string;
  /** The package path (`ppt/media/image3.png`), or a URL for an external target. */
  target: string;
  external: boolean;
}

/** The markup of one element and what it depends on, enough for the splice to rebuild it in another deck. */
export interface Markup {
  /** The element's top-level shapes, serialised, in slide order. */
  xml: string;
  /** Every relationship the xml names, by id. */
  rels: MarkupRel[];
  /** Every package part reachable from those relationships (media, charts, embeddings, diagrams, tags), by path. */
  parts: string[];
}

export interface Element {
  /** Stable across deck reorders and both sizes: a slug of the key. */
  id: string;
  /** The deck's Danish title (a whole-slide element) or part name (a shape on a collection slide). */
  key: string;
  /** The English name from the names file. */
  name: string;
  category: { key: string; name: string };
  /** 1-based position in the deck, for the harvest report and the preview cut. */
  slide: number;
  kind: "slide" | "part";
  box: Box;
  landing: Landing;
  /** Top-level shapes the element consists of: 1 for a part. */
  shapes: number;
  tags: string[];
  run?: SizeRun;
  markup: Markup;
}

export interface Catalogue {
  size: SlideSize;
  /** The deck's slide size in EMU. */
  width: number;
  height: number;
  categories: { key: string; name: string }[];
  elements: Element[];
}

/** The names a locale adds to the deck: categories and elements, keyed by the deck's Danish titles. */
export interface Names {
  categories: Record<string, string>;
  names: Record<string, string>;
}

/** What the harvest hands back: the catalogue and the bytes of every part the elements reach. */
export interface Harvest {
  catalogue: Catalogue;
  parts: Map<string, Uint8Array | string>;
}
