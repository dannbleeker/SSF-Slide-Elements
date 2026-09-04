/**
 * Read a library deck and produce the catalogue.
 *
 * The whole design principle is that the LIBRARY DECK IS THE AUTHORING
 * SURFACE. Adding an element is drawing it on a slide and pushing; there is no
 * sidecar file to keep in sync, no naming convention to remember, and nothing
 * to hand-edit. Everything this file needs is already in the deck that
 * PowerPoint saved.
 *
 * Three signals carry it, all read from the deck as it stands:
 *
 * - **The LAYOUT says what a slide is for.** A slide on the divider layout
 *   opens a section; a slide on the content layout is an element; the cover is
 *   neither. Measured against the real library: 1 cover, 10 dividers, 98
 *   elements, no exceptions.
 * - **The TITLE placeholder is the name.** Of a section on a divider, of an
 *   element on a content slide.
 * - **A PLACEHOLDER is furniture and everything else is content.** Also
 *   measured: every one of the 98 content slides carries exactly two
 *   placeholders, `title` and `sldNum`, and 677 non-placeholder shapes between
 *   them. The title is the element's name, not part of it, and the slide number
 *   belongs to the library deck.
 *
 * `LayoutRoles` names the layouts rather than hard-coding the Danish ones, so a
 * library in another language needs a config line and not a patch.
 */
import { Pkg } from "../pptx/pkg.js";
import {
  A_NS,
  MC_NS,
  P_NS,
  child,
  childElements,
  element,
  elements,
  relationshipIdsIn,
  serializeXml,
} from "../pptx/xml.js";
import { FOLLOWED, REL_TYPE, TRAVELS_WITH_ELEMENT } from "../pptx/parts.js";
import type { Box, CatalogueIndex, ElementEntry, ElementPayload, PayloadPart, SectionEntry } from "./types.js";

/**
 * Which layout means what.
 *
 * Matched against the layout's `<p:cSld name="…">` — the name PowerPoint shows
 * in the layout gallery, which is what somebody maintaining the library
 * actually sees. Case-insensitive because that is a difference nobody intends.
 */
export interface LayoutRoles {
  /** Opens a section. Its title is the section name. */
  divider: string[];
  /** Holds an element. Its title is the element name. */
  content: string[];
  /** Neither — a cover, a back page, a note to the maintainer. */
  ignore: string[];
}

/** The roles as the shipped library uses them. */
export const DEFAULT_ROLES: LayoutRoles = {
  divider: ["Sektionsadskillelse", "Section Header", "Section Divider"],
  content: ["Kun titel", "Title Only"],
  ignore: ["Frontpage", "Title Slide", "Front Page"],
};

/** The top-level members of a shape tree that are shapes rather than bookkeeping. */
const SHAPE_TAGS = new Set(["sp", "grpSp", "graphicFrame", "pic", "cxnSp"]);

/** What went wrong, or what was skipped and why. The harvest reports, never guesses. */
export interface HarvestNote {
  slide: number;
  level: "skipped" | "warning";
  detail: string;
}

export interface HarvestResult {
  index: CatalogueIndex;
  payloads: ElementPayload[];
  notes: HarvestNote[];
}

/**
 * A stable id for an element.
 *
 * Derived from the SECTION and the element's NAME rather than from the slide
 * number, because slide numbers move: inserting one element at the top of a
 * section would otherwise renumber every element below it, and every id the
 * pane has remembered as a favourite would point at the wrong thing.
 *
 * A name that repeats within a section gets a numeric suffix, which is the one
 * case where an id is position-dependent. The harvest reports it as a warning:
 * two elements with the same name in the same section is almost always a
 * mistake in the deck, and the id is only stable while the order holds.
 */
export function elementId(section: string, name: string, seen: Map<string, number>): { id: string; repeated: boolean } {
  const base = `${slug(section)}--${slug(name)}`;
  const n = (seen.get(base) ?? 0) + 1;
  seen.set(base, n);
  return n === 1 ? { id: base, repeated: false } : { id: `${base}-${n}`, repeated: true };
}

/**
 * A name reduced to something that can be a file name and a URL segment.
 *
 * Danish letters are transliterated rather than dropped: "Hvid kasse på vægt"
 * losing its vowels would produce `hvid-kasse-p-vgt`, which is not a name
 * anybody can match back to a slide when they are reading a bug report. Every
 * element in the shipped library is Danish, so this is the common case, not an
 * edge one.
 */
export function slug(text: string): string {
  const folded = text
    .toLowerCase()
    .replace(/æ/g, "ae")
    .replace(/ø/g, "oe")
    .replace(/å/g, "aa")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
  const out = folded.replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return out || "unnamed";
}

/** The text of a shape, runs joined in document order. */
function textOf(shape: Element): string {
  return elements(shape, A_NS, "t")
    .map((t) => t.textContent ?? "")
    .join("")
    .trim();
}

/** The `<p:ph>` on a shape, if it is a placeholder. */
function placeholderOf(shape: Element): Element | undefined {
  const nv = child(shape, P_NS, "nvSpPr") ?? child(shape, P_NS, "nvPicPr") ?? child(shape, P_NS, "nvGraphicFramePr");
  if (!nv) return undefined;
  const nvPr = child(nv, P_NS, "nvPr");
  return nvPr ? child(nvPr, P_NS, "ph") : undefined;
}

/**
 * Where a top-level shape sits, in slide coordinates.
 *
 * `<a:xfrm>` is read from the shape's OWN properties, never by descending —
 * a group's children each carry one in the group's child coordinate space, and
 * the first `<a:off>` found by a descendant search on a `<p:grpSp>` belongs to
 * the group itself only by luck of document order.
 *
 * Undefined when the shape states no transform, which for a top-level shape
 * means it inherits from a placeholder. Those are furniture here and are
 * filtered out before this is asked.
 */
export function boundsOf(shape: Element): Box | undefined {
  const spPr = child(shape, P_NS, "spPr") ?? child(shape, P_NS, "grpSpPr") ?? child(shape, P_NS, "xfrm") ?? undefined;
  const xfrm = spPr ? (spPr.localName === "xfrm" ? spPr : child(spPr, A_NS, "xfrm")) : undefined;
  if (!xfrm) return undefined;
  const off = child(xfrm, A_NS, "off");
  const ext = child(xfrm, A_NS, "ext");
  const x = Number(off?.getAttribute("x"));
  const y = Number(off?.getAttribute("y"));
  const cx = Number(ext?.getAttribute("cx"));
  const cy = Number(ext?.getAttribute("cy"));
  if (![x, y, cx, cy].every(Number.isFinite)) return undefined;
  return { x, y, cx, cy };
}

/** The smallest box containing all of them. */
export function unionBounds(boxes: Box[]): Box {
  if (boxes.length === 0) return { x: 0, y: 0, cx: 0, cy: 0 };
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (const b of boxes) {
    x0 = Math.min(x0, b.x);
    y0 = Math.min(y0, b.y);
    x1 = Math.max(x1, b.x + b.cx);
    y1 = Math.max(y1, b.y + b.cy);
  }
  return { x: x0, y: y0, cx: x1 - x0, cy: y1 - y0 };
}

/**
 * What a shape IS, for the picker's badge.
 *
 * A `<p:graphicFrame>` is the interesting one: table, chart and SmartArt all
 * arrive as one, and which it is shows only in the `<a:graphicData uri="…">`
 * inside it.
 */
export function kindOf(shape: Element): string {
  switch (shape.localName) {
    case "pic":
      return "picture";
    case "grpSp":
      return "group";
    case "cxnSp":
      return "connector";
    case "graphicFrame": {
      const data = element(shape, A_NS, "graphicData");
      const uri = data?.getAttribute("uri") ?? "";
      if (uri.includes("/table")) return "table";
      if (uri.includes("/chart")) return "chart";
      if (uri.includes("/diagram")) return "diagram";
      return "graphic";
    }
    case "AlternateContent":
      return "chart";
    default:
      return textOf(shape) ? "text" : "shape";
  }
}

/** The layout a slide is built on, by the name PowerPoint shows for it. */
async function layoutNameOf(pkg: Pkg, slidePath: string): Promise<string | undefined> {
  for (const rel of await pkg.rels(slidePath)) {
    if (rel.getAttribute("Type") !== REL_TYPE.slideLayout) continue;
    const target = await pkg.relTarget(slidePath, rel.getAttribute("Id") ?? "");
    if (!target || target.external) continue;
    const doc = await pkg.doc(target.path);
    const cSld = element(doc, P_NS, "cSld");
    return cSld?.getAttribute("name") ?? undefined;
  }
  return undefined;
}

function roleOf(layout: string | undefined, roles: LayoutRoles): "divider" | "content" | "ignore" | "unknown" {
  if (layout === undefined) return "unknown";
  const want = layout.trim().toLowerCase();
  const has = (list: string[]): boolean => list.some((n) => n.trim().toLowerCase() === want);
  if (has(roles.divider)) return "divider";
  if (has(roles.content)) return "content";
  if (has(roles.ignore)) return "ignore";
  return "unknown";
}

/**
 * Gather every part an element's markup reaches, following relationships
 * transitively.
 *
 * Transitively because a chart is not one part: the `<p:graphicFrame>` names
 * `chart1.xml`, which names the workbook behind it, its colours and its style,
 * and any picture drawn on top of it. Stopping at the first hop produces an
 * element whose chart opens as an empty frame.
 *
 * `TRAVELS_WITH_ELEMENT` is the gate, and it is an allowlist for a reason a
 * comment in `parts.ts` gives at length: relationship targets come out of a
 * file, and a deck can be sent to somebody.
 */
async function gatherParts(
  pkg: Pkg,
  ownerPart: string,
  rIds: Set<string>,
  notes: HarvestNote[],
  slide: number,
): Promise<{ rels: ElementPayload["rels"]; parts: PayloadPart[] }> {
  const rels: ElementPayload["rels"] = [];
  const parts = new Map<string, PayloadPart>();
  const visited = new Set<string>();

  /** Read a part into the payload once, as text when it is XML and bytes otherwise. */
  const take = async (path: string): Promise<PayloadPart | undefined> => {
    const already = parts.get(path);
    if (already) return already;
    const contentType = (await pkg.contentTypeOf(path)) ?? "application/octet-stream";
    // XML parts are kept as text so the harvested output diffs readably;
    // everything else is bytes, and bytes have to be base64 to survive JSON.
    const part: PayloadPart = contentType.includes("xml")
      ? { path, contentType, text: await pkg.text(path) }
      : { path, contentType, base64: toBase64(await pkg.bytes(path)) };
    parts.set(path, part);
    return part;
  };

  /**
   * Everything this part reaches, and everything those reach.
   *
   * A part's own `.rels` travels with it unchanged: the ids inside it are only
   * meaningful relative to that part, and that part is copied intact. Only the
   * ids the SLIDE's shapes name have to be rewritten, and those are the `top`
   * ones handled by the caller.
   */
  const follow = async (path: string): Promise<void> => {
    if (visited.has(path)) return;
    visited.add(path);
    const relsPath = Pkg.relsPathFor(path);
    if (!pkg.has(relsPath)) return;
    await take(relsPath);
    for (const rel of await pkg.rels(path)) {
      const target = await pkg.relTarget(path, rel.getAttribute("Id") ?? "");
      if (!target || target.external) continue;
      if (!TRAVELS_WITH_ELEMENT.test(target.path)) {
        notes.push({ slide, level: "warning", detail: `refused to carry ${target.path}, reached from ${path}` });
        continue;
      }
      await take(target.path);
      await follow(target.path);
    }
  };

  for (const rId of rIds) {
    const target = await pkg.relTarget(ownerPart, rId);
    if (!target) {
      notes.push({ slide, level: "warning", detail: `the slide names ${rId}, which its relationships do not define` });
      continue;
    }
    // An external target is a URL — a hyperlink, or a linked picture. It is
    // recorded so the splice can recreate the relationship, and nothing is
    // copied, because there is nothing in the package to copy.
    if (target.external) {
      rels.push({ rId, type: target.type, path: target.path, external: true });
      continue;
    }
    // Relationships the slide owns rather than the shapes — its layout, its
    // notes, its own tags — describe the LIBRARY's slide and must not travel.
    if (!FOLLOWED.has(target.type)) continue;
    if (!TRAVELS_WITH_ELEMENT.test(target.path)) {
      notes.push({
        slide,
        level: "warning",
        detail: `refused to carry ${target.path}, which is not a part an element may own`,
      });
      continue;
    }
    rels.push({ rId, type: target.type, path: target.path });
    await take(target.path);
    await follow(target.path);
  }

  return { rels, parts: [...parts.values()] };
}

/** Base64 without Node's Buffer, so the same code runs in the pane. */
export function toBase64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    s += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
  }
  return btoa(s);
}

/**
 * Harvest a library deck.
 *
 * Never throws on a slide it cannot read: the note goes into `notes` and the
 * rest of the library is still produced. A library is a hundred slides
 * maintained by hand, and one bad slide taking the whole catalogue with it is
 * how a build gate gets switched off.
 */
export async function harvest(
  pkg: Pkg,
  opts: {
    libraryName: string;
    roles?: LayoutRoles;
    preview?: (shapes: Element[], bounds: Box, slide: { cx: number; cy: number }) => string;
  },
): Promise<HarvestResult> {
  const roles = opts.roles ?? DEFAULT_ROLES;
  const notes: HarvestNote[] = [];
  const sections: SectionEntry[] = [];
  const entries: ElementEntry[] = [];
  const payloads: ElementPayload[] = [];
  const seenIds = new Map<string, number>();
  const slideSize = await pkg.slideSize();
  const slides = await pkg.slidePaths();

  let current: SectionEntry | undefined;

  for (const [i, slidePath] of slides.entries()) {
    const slideNo = i + 1;
    const layout = await layoutNameOf(pkg, slidePath);
    const role = roleOf(layout, roles);

    if (role === "ignore") continue;
    if (role === "unknown") {
      notes.push({
        slide: slideNo,
        level: "skipped",
        detail: `layout ${layout === undefined ? "(none)" : `"${layout}"`} has no role — add it to LayoutRoles to include or ignore this slide`,
      });
      continue;
    }

    const doc = await pkg.doc(slidePath);
    const tree = element(doc, P_NS, "spTree");
    if (!tree) {
      notes.push({ slide: slideNo, level: "skipped", detail: "the slide has no shape tree" });
      continue;
    }

    // The title placeholder names the thing, whichever kind of slide it is.
    let title = "";
    const content: Element[] = [];
    for (const node of childElements(tree)) {
      const local = node.localName ?? "";
      const isAlternate = local === "AlternateContent" && node.namespaceURI === MC_NS;
      if (!SHAPE_TAGS.has(local) && !isAlternate) continue;
      const ph = placeholderOf(node);
      if (ph) {
        const type = ph.getAttribute("type") ?? "body";
        if (type === "title" || type === "ctrTitle") title = textOf(node);
        // Every placeholder is furniture: the title names the element and the
        // slide number belongs to the library deck. Measured across the whole
        // shipped library — see this file's header.
        continue;
      }
      content.push(node);
    }

    if (role === "divider") {
      if (!title) {
        notes.push({ slide: slideNo, level: "skipped", detail: "a divider slide with no title cannot name a section" });
        continue;
      }
      current = { id: slug(title), name: title, source: { slide: slideNo } };
      sections.push(current);
      continue;
    }

    if (!title) {
      notes.push({
        slide: slideNo,
        level: "skipped",
        detail: "a content slide with no title has no name to show in the picker",
      });
      continue;
    }
    if (content.length === 0) {
      notes.push({
        slide: slideNo,
        level: "skipped",
        detail: `"${title}" has a title and nothing else — there is no element to insert`,
      });
      continue;
    }
    if (!current) {
      notes.push({
        slide: slideNo,
        level: "skipped",
        detail: `"${title}" comes before any divider slide, so it belongs to no section`,
      });
      continue;
    }

    const boxes = content.map(boundsOf).filter((b): b is Box => b !== undefined);
    if (boxes.length !== content.length) {
      notes.push({
        slide: slideNo,
        level: "warning",
        detail: `"${title}" has ${content.length - boxes.length} shape(s) that state no transform; they travel but do not count towards its bounds`,
      });
    }
    const bounds = unionBounds(boxes);

    const rIds = new Set<string>();
    for (const node of content) for (const id of relationshipIdsIn(node)) rIds.add(id);
    const { rels, parts } = await gatherParts(pkg, slidePath, rIds, notes, slideNo);

    // Whether the name repeated is answered by the counter, never by looking at
    // the id: half the library's names END IN A DIGIT ("…med checkliste på 4"),
    // so a regex for a numeric suffix reports every one of them as a duplicate
    // and the real duplicates are lost in the noise.
    const { id, repeated } = elementId(current.id, title, seenIds);
    if (repeated) {
      notes.push({
        slide: slideNo,
        level: "warning",
        detail: `"${title}" repeats a name already used in "${current.name}" — the picker will show two identical entries, and this one's id depends on slide order`,
      });
    }

    const shapes = content.map((n) => serializeXml(n));
    const payload: ElementPayload = { version: 1, id, shapes, rels, parts, bounds };
    payloads.push(payload);

    entries.push({
      id,
      name: title,
      section: current.id,
      source: { slide: slideNo, part: slidePath },
      bounds,
      shapes: content.length,
      kinds: [...new Set(content.map(kindOf))].sort(),
      preview: opts.preview ? opts.preview(content, bounds, slideSize) : "",
      payloadBytes: JSON.stringify(payload).length,
    });
  }

  return {
    index: {
      version: 1,
      library: { name: opts.libraryName, slideSize, harvestedAt: new Date().toISOString() },
      sections,
      elements: entries,
    },
    payloads,
    notes,
  };
}
