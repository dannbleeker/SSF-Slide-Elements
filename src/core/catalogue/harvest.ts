/**
 * The harvest: a library deck in, a catalogue out.
 *
 * The deck is the authoring surface (`docs/DESIGN.md` sections 2 and 3):
 *
 * - A heading slide (a title and nothing else) starts a category.
 * - Every other slide under it is one element, named by its title, unless the
 *   category's heading slide carries the collection marker in its notes, in
 *   which case every top-level shape on the slide is an element of its own,
 *   the slide's title is their category, and each is named by its own text.
 *   A slide's own notes can override the marker either way.
 * - Layout chrome (title, footer, slide number, date) and empty placeholders
 *   are never part of an element; a shape entirely off the slide is not either.
 * - The English name of every key comes from the names file, and a key without
 *   one fails the harvest, all of them listed at once.
 *
 * Pure: reads through `Pkg`, touches nothing, and knows nothing of Office.js.
 */
import { XMLSerializer } from "@xmldom/xmldom";
import { Pkg } from "../pptx/pkg.js";
import { A_NS, P_NS, PKG_REL_NS, R_NS, elements, element } from "../pptx/xml.js";
import { boxOf, offSlide, rounded, topLevelShapes, union } from "./boxes.js";
import { sizeRuns } from "./runs.js";
import { tagsFor } from "./tags.js";
import { paragraphsOf, partName, placeholderType, slug, textOf, titleOf } from "./text.js";
import type {
  Box,
  Catalogue,
  Element as CatalogueElement,
  Harvest,
  Landing,
  Markup,
  MarkupRel,
  Names,
  SlideSize,
} from "./types.js";

/** The marker a heading slide's notes carry to make the slides under it collection slides; a slide's own notes can override. */
const PER_SHAPE = /^\s*SSF:\s*(?:ét|et|one)\s+element\s+(?:pr\.?|per)\s+(?:figur|shape)\s*$/im;
const PER_SLIDE = /^\s*SSF:\s*(?:ét|et|one)\s+element\s+(?:pr\.?|per)\s+(?:dias|slide)\s*$/im;

/** Placeholder types that are layout chrome, never content. */
const CHROME = new Set(["title", "ctrTitle", "ftr", "sldNum", "dt"]);

/** The parts an element may carry with it: what a slide owns through its graphics. Layouts, masters, themes and notes are the destination's. */
const CARRIED = /^ppt\/(charts|diagrams|drawings|embeddings|media|tags)\//;

export class HarvestError extends Error {
  constructor(
    message: string,
    readonly problems: string[],
  ) {
    super(message);
    this.name = "HarvestError";
  }
}

/** Layout chrome, or a placeholder with nothing in it: not content. A table or a picture in a content placeholder is content. */
function isChrome(shape: Element): boolean {
  const type = placeholderType(shape);
  if (type === undefined) return false;
  if (CHROME.has(type)) return true;
  const hasText = textOf(shape).length > 0;
  const hasGraphic = elements(shape, A_NS, "graphic").length > 0 || elements(shape, A_NS, "blip").length > 0;
  return !(hasText || hasGraphic);
}

/** Every `r:*` attribute value under an element: the relationship ids its markup names. */
function relIdsIn(el: Element): Set<string> {
  const out = new Set<string>();
  const walk = (node: Element): void => {
    for (let i = 0; i < node.attributes.length; i++) {
      const attr = node.attributes.item(i);
      if (attr && attr.namespaceURI === R_NS && attr.value) out.add(attr.value);
    }
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes.item(i);
      if (child && child.nodeType === 1) walk(child as Element);
    }
  };
  walk(el);
  return out;
}

async function relationshipsOf(pkg: Pkg, part: string): Promise<MarkupRel[]> {
  const path = Pkg.relsPathFor(part);
  if (!pkg.has(path)) return [];
  const doc = await pkg.doc(path);
  const out: MarkupRel[] = [];
  for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
    const id = rel.getAttribute("Id");
    const type = rel.getAttribute("Type");
    const target = rel.getAttribute("Target");
    if (!id || !type || !target) continue;
    const external = (rel.getAttribute("TargetMode") ?? "") === "External";
    out.push({ id, type, target: external ? target : pkg.resolved(part, target), external });
  }
  return out;
}

/** The notes text of a slide, or "" when it has no notes page. */
async function notesOf(pkg: Pkg, slidePath: string): Promise<string> {
  const notes = (await pkg.relatedParts(slidePath)).find((p) => p.startsWith("ppt/notesSlides/"));
  if (!notes) return "";
  const doc = await pkg.doc(notes);
  return elements(doc, P_NS, "sp")
    .filter((sp) => (placeholderType(sp) ?? "body") === "body")
    .flatMap((sp) => paragraphsOf(sp))
    .join("\n");
}

/** Every carried part reachable from a set of relationship targets, through their own relationships, deduplicated. */
async function reachableParts(
  pkg: Pkg,
  starts: string[],
  seen: Set<string>,
  collect: Map<string, Uint8Array | string>,
): Promise<string[]> {
  const out: string[] = [];
  const queue = starts.filter((p) => CARRIED.test(p));
  while (queue.length) {
    const part = queue.shift() as string;
    if (seen.has(part) || !pkg.has(part)) continue;
    seen.add(part);
    out.push(part);
    collect.set(part, /\.(xml|rels)$/.test(part) ? await pkg.text(part) : await pkg.bytes(part));
    const rels = Pkg.relsPathFor(part);
    if (pkg.has(rels)) collect.set(rels, await pkg.text(rels));
    for (const next of await pkg.relatedParts(part)) if (CARRIED.test(next)) queue.push(next);
  }
  return out;
}

async function markupFor(
  pkg: Pkg,
  slidePath: string,
  shapes: Element[],
  seen: Set<string>,
  collect: Map<string, Uint8Array | string>,
): Promise<Markup> {
  const serializer = new XMLSerializer();
  const xml = shapes.map((s) => serializer.serializeToString(s as never)).join("");
  const ids = new Set<string>();
  for (const s of shapes) for (const id of relIdsIn(s)) ids.add(id);
  const rels = (await relationshipsOf(pkg, slidePath)).filter((r) => ids.has(r.id));
  const parts = await reachableParts(
    pkg,
    rels.filter((r) => !r.external).map((r) => r.target),
    seen,
    collect,
  );
  // A part reached from the slide's relationships but not by the CARRIED rule is the destination's business
  // (a layout, a notes page); the splice never copies it, so it is not listed.
  return { xml, rels, parts };
}

function landingFor(kind: "slide" | "part", collectionTitle: string, box: Box): Landing {
  if (kind === "slide") return "layout";
  if (box.w > 0.5) return "as-authored";
  return /stempl/i.test(collectionTitle) ? "top-right" : "cursor";
}

export interface HarvestOptions {
  size: SlideSize;
  names: Names;
}

/**
 * Read a library deck into a catalogue.
 *
 * Throws `HarvestError` with every problem found when a key or a category has
 * no English name, two elements share a key, or a slide with content has no
 * title: the deck is the owner's to fix and one run should list it all.
 */
export async function harvest(pkg: Pkg, options: HarvestOptions): Promise<Harvest> {
  const pres = await pkg.doc("ppt/presentation.xml");
  const sldSz = element(pres, P_NS, "sldSz");
  const width = Number(sldSz?.getAttribute("cx") ?? 0);
  const height = Number(sldSz?.getAttribute("cy") ?? 0);
  if (!(width > 0 && height > 0))
    throw new HarvestError("the deck states no slide size", ["ppt/presentation.xml has no <p:sldSz>"]);

  const problems: string[] = [];
  const parts = new Map<string, Uint8Array | string>();
  const seen = new Set<string>();
  const categories: { key: string; name: string }[] = [];
  const elementsOut: CatalogueElement[] = [];
  const keysSeen = new Map<string, number>();

  const category = (key: string): { key: string; name: string } => {
    const existing = categories.find((c) => c.key === key);
    if (existing) return existing;
    const name = options.names.categories[key];
    if (name === undefined) problems.push(`category "${key}" has no English name in the names file`);
    const made = { key, name: name ?? key };
    categories.push(made);
    return made;
  };
  const nameOf = (key: string, slide: number): string => {
    const name = options.names.names[key];
    if (name === undefined) problems.push(`"${key}" (slide ${slide}) has no English name in the names file`);
    const before = keysSeen.get(key);
    if (before !== undefined) problems.push(`"${key}" is the key of both slide ${before} and slide ${slide}`);
    keysSeen.set(key, slide);
    return name ?? key;
  };

  let heading: { key: string; perShape: boolean } | undefined;
  const slidePaths = await pkg.slidePaths();
  for (let i = 0; i < slidePaths.length; i++) {
    const slidePath = slidePaths[i] as string;
    const slideNo = i + 1;
    const doc = await pkg.doc(slidePath);
    const title = titleOf(doc);
    const content: { shape: Element; box: Box | undefined }[] = [];
    for (const shape of topLevelShapes(doc)) {
      if (isChrome(shape)) continue;
      const box = boxOf(shape, width, height);
      if (box && offSlide(box)) continue;
      content.push({ shape, box });
    }
    const notes = await notesOf(pkg, slidePath);
    if (content.length === 0) {
      // A heading, when it has a title. A slide with nothing on it at all is skipped.
      if (title) heading = { key: title, perShape: PER_SHAPE.test(notes) };
      continue;
    }
    if (!heading) continue; // the cover, and anything before the first heading
    if (!title) {
      problems.push(`slide ${slideNo} has content but no title, so it has no key`);
      continue;
    }
    const perShape = PER_SLIDE.test(notes) ? false : PER_SHAPE.test(notes) ? true : heading.perShape;
    if (perShape) {
      const cat = category(title);
      const used = new Set<string>();
      const nameless = { count: 0 };
      for (const { shape, box } of content) {
        const key = partName(shape, title, used, nameless);
        const b = rounded(box ?? { x: 0, y: 0, w: 1, h: 1 });
        elementsOut.push({
          id: slug(key),
          key,
          name: nameOf(key, slideNo),
          category: cat,
          slide: slideNo,
          kind: "part",
          box: b,
          landing: landingFor("part", title, b),
          shapes: 1,
          tags: tagsFor(key, title, true),
          markup: await markupFor(pkg, slidePath, [shape], seen, parts),
        });
      }
      continue;
    }
    const cat = category(heading.key);
    const b = rounded(union(content.map((c) => c.box).filter((x): x is Box => x !== undefined)));
    elementsOut.push({
      id: slug(title),
      key: title,
      name: nameOf(title, slideNo),
      category: cat,
      slide: slideNo,
      kind: "slide",
      box: b,
      landing: "layout",
      shapes: content.length,
      tags: tagsFor(title, heading.key),
      markup: await markupFor(
        pkg,
        slidePath,
        content.map((c) => c.shape),
        seen,
        parts,
      ),
    });
  }

  const ids = new Map<string, string>();
  for (const el of elementsOut) {
    const other = ids.get(el.id);
    if (other !== undefined && other !== el.key)
      problems.push(`"${el.key}" and "${other}" both slug to the id "${el.id}"`);
    ids.set(el.id, el.key);
  }
  if (problems.length)
    throw new HarvestError(`the ${options.size} deck cannot be harvested: ${problems.length} problem(s)`, problems);

  const runs = sizeRuns(elementsOut.filter((e) => e.kind === "slide").map((e) => e.name));
  for (const el of elementsOut) {
    const run = runs.get(el.name);
    if (run && el.kind === "slide") el.run = run;
  }

  const catalogue: Catalogue = { size: options.size, width, height, categories, elements: elementsOut };
  return { catalogue, parts };
}
