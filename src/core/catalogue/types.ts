/**
 * What the harvest produces and the pane consumes.
 *
 * Split in two on purpose. The INDEX is small enough to load at startup and
 * hold in memory: it is everything the picker draws. The PAYLOAD is the markup
 * and the bytes, one file per element, fetched only when somebody actually
 * inserts one. A library of a hundred elements carrying embedded workbooks and
 * icon SVGs is megabytes; the index for the same library is tens of kilobytes.
 */

/** Every measurement here is an EMU, the unit the file format uses. */
export interface Box {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/**
 * One insertable thing.
 *
 * `id` is derived from the source slide and is stable across a re-harvest as
 * long as the slide keeps its position in its section — see `elementId`. It is
 * what the pane remembers as a favourite and what a payload file is named
 * after, so a change to how it is derived invalidates both.
 */
export interface ElementEntry {
  id: string;
  /** The slide's title placeholder, verbatim. The name a human gave it. */
  name: string;
  /** The id of the section this belongs to. */
  section: string;
  /** Where this came from, for the harvest report and for bug reports. */
  source: { slide: number; part: string };
  /** The element's bounding box on its source slide, as authored. */
  bounds: Box;
  /** How many top-level shapes travel. One is common; twenty-three happens. */
  shapes: number;
  /** Kinds present, for the picker's badges: text, table, chart, picture, group. */
  kinds: string[];
  /** An inline SVG preview, drawn from the harvested geometry. */
  preview: string;
  /** Bytes the payload will cost, so the pane can warn before a slow fetch. */
  payloadBytes: number;
}

export interface SectionEntry {
  id: string;
  /** The divider slide's title. The category name, as authored. */
  name: string;
  source: { slide: number };
}

/** The light file, loaded once when the pane opens. */
export interface CatalogueIndex {
  /** Bumped when the shape of these files changes. The pane refuses a stranger. */
  version: 1;
  /** Which library deck this came from, and when. */
  library: { name: string; slideSize: { cx: number; cy: number }; harvestedAt: string };
  sections: SectionEntry[];
  elements: ElementEntry[];
}

/** A part that travels with an element: markup, or bytes as base64. */
export interface PayloadPart {
  path: string;
  contentType: string;
  /** Exactly one of these. Markup is kept as text so it stays diffable. */
  text?: string;
  base64?: string;
}

/**
 * The heavy file, one per element.
 *
 * `shapes` is a list of serialised top-level members of the source slide's
 * `<p:spTree>` — `<p:sp>`, `<p:grpSp>`, `<p:graphicFrame>`, `<p:pic>`,
 * `<p:cxnSp>` or an `<mc:AlternateContent>` wrapping one. They are stored as
 * TEXT rather than as a parsed tree because that is what survives a JSON round
 * trip without a schema, and because a diff of the harvested output is then
 * something a person can read.
 *
 * `rels` maps the relationship ids those shapes name to the parts in this
 * payload. The splice rewrites every one of them, because the ids are only
 * unique within the slide they came from.
 */
export interface ElementPayload {
  version: 1;
  id: string;
  shapes: string[];
  rels: { rId: string; type: string; path: string; external?: boolean }[];
  parts: PayloadPart[];
  bounds: Box;
}
