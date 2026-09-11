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
  /**
   * The carried parts this element is the FIRST to reach, by path.
   *
   * A manifest of what this element contributes to the shared part store, not
   * the element's own closure — and the difference bites. Parts are collected
   * once per deck, so an element that shares a picture with an earlier one
   * lists nothing for it while still naming it in `rels`: measured on the
   * committed library, 14 of the 16:9 elements' relationship targets and 50 of
   * the 4:3 elements' are absent from their own `parts`.
   *
   * **So the splice resolves what to copy from `rels`, and walks each part's
   * own relationships from there.** This list is what the harvest writes to
   * disk under `<size>/parts/`, and nothing should read it as a dependency set.
   */
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
  /**
   * For a rotated part, the UNROTATED frame `box` was computed from, and the
   * angle in degrees clockwise.
   *
   * Only the preview cut needs it: `box` is the rotated extent, so its corners
   * hold whatever the slide has behind the element, and masking them away needs
   * the frame back (`docs/DESIGN.md` section 3). Absent when the element is not
   * rotated, which is all but the owner's stamps.
   */
  rotation?: { deg: number; frame: Box };
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
  /**
   * What this deck's `<a:schemeClr>` names resolve to: the library theme's
   * twelve colours, plus the four the master's colour map redirects.
   *
   * The switch behind the gear (`docs/DESIGN.md` section 7) pins an element's
   * colours to these when the user asks for "As in the library". Without the
   * map there is nothing to pin them TO — the theme part is in the library
   * deck, which the pane never sees — so the harvest carries it.
   *
   * One map per SIZE, not per element, and the harvest refuses a deck whose
   * slides do not agree on a theme rather than picking one. Measured on the
   * committed library, 2026-09-11: one master and one theme in each deck, and
   * since the 4:3 deck was re-themed the same day, the same colours in both —
   * so an element pinned in one size is pinned to what it is in the other.
   * `test/colours.test.ts` asserts that equality rather than the hexes, and
   * `docs/DESIGN.md` section 3 says what the two decks disagreed about.
   */
  theme: Record<string, string>;
  /**
   * Every carried part in this size's store, with the content type the library
   * deck declared for it.
   *
   * The splice writes these parts into somebody else's package, and a part with
   * no content type declared there is a package PowerPoint refuses outright.
   * Carried from the source rather than inferred from the extension: a table
   * mapping `.bin` to an embedded object and `.emf` to a picture is right for
   * today's library and silently wrong for the first family the owner adds.
   *
   * On the CATALOGUE rather than on an element, because the store is shared:
   * two elements using the same picture must not disagree about what it is.
   */
  carried: Record<string, string>;
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
