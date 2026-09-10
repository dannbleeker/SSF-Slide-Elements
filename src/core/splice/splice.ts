/**
 * An element into the slide the user is on, in the file.
 *
 * This is the whole mechanism `docs/DESIGN.md` section 6 describes, and it is
 * one shape: **rebuild the destination slide with the element on it, hand
 * PowerPoint the rebuild, and take the original away.** Nothing is drawn
 * through the shape collection — `docs/BACKLOG.md` lists that under "Rejected —
 * do not re-propose", with the sibling rounds that rejected it.
 *
 * The package handed back holds the user's ENTIRE deck with one slide added and
 * every other slide UNLISTED: its part is still in the zip, its relationship
 * still points at it, but no `<p:sldId>` names it. PowerPoint for the web was
 * asked directly whether it accepts that, on 2026-09-10, because the whole cost
 * of the insert turns on it — `docs/host-answers/` holds the sheets and
 * `docs/PROBE.md` question 1 the reading. It does: both prunings landed exactly
 * one slide, so the host reads the slide list and ignores a part nothing lists.
 * The expensive alternative was removing every other slide properly, orphans
 * walked, on a deck that can be fifty megabytes.
 *
 * The second measured fact this rests on is question 1b. An insert of the
 * deck's OWN bytes added no master under either formatting option (1 → 1),
 * while a foreign fixture deck added one (1 → 2). This package is built from
 * the user's own deck, so the master, the layout and the theme the rebuilt
 * slide points at are the ones already in it, and the insert does not grow the
 * deck by a master. Building the slide from the library deck instead would have.
 */
import type { MarkupRel } from "../catalogue/types.js";
import { cloneSlide } from "../pptx/clone.js";
import { framesOf, slideSize } from "../pptx/layout.js";
import { Pkg } from "../pptx/pkg.js";
import { REL_TYPE } from "../pptx/parts.js";
import { TAG_CATALOGUE, TAG_ELEMENT, writeShapeTags } from "../pptx/tags.js";
import { A_NS, PKG_REL_NS, P_NS, R_NS, child, children, element, elements } from "../pptx/xml.js";
import { carry, type PartStore } from "./carry.js";
import { authored, moveFrom, place, type Box, type Landing, type Rect } from "./landing.js";
import {
  applyMove,
  emptyBodyPlaceholders,
  groupShapes,
  groupable,
  highestShapeId,
  parseFragment,
  renumber,
  repoint,
  slideShapes,
  topLevel,
  unplaceholder,
} from "./shapes.js";

const PRESENTATION = "ppt/presentation.xml";

/** What the pane knows about the element it is inserting. */
export interface SpliceElement {
  /** The catalogue id, written into the shape's tag so the deck can be read back. */
  id: string;
  /** The English name, which becomes the shape's name in the selection pane. */
  name: string;
  kind: "slide" | "part";
  box: Box;
  landing: Landing;
  /** A marker wraps the selected shape rather than merely sitting on it. */
  wraps?: boolean;
  markup: { xml: string; rels: MarkupRel[] };
}

export interface SpliceOptions {
  /** `docs/DESIGN.md` section 7: onto this slide, or as a new slide after it. */
  target: "onto" | "new";
  /** Section 7 again: multi-shape elements land as one group, or loose. */
  group: boolean;
}

export interface SpliceRequest {
  /** The user's deck, as `getFileAsync` handed it over. */
  deck: Uint8Array | string;
  /** Which slide the user is on, counting from zero. */
  slide: number;
  element: SpliceElement;
  options: SpliceOptions;
  catalogue: { version: string; carried: Record<string, string> };
  store: PartStore;
  /** The selected shape's rectangle, when the host could name one. */
  selection?: Rect;
}

/**
 * What the splice did, in numbers the pane can show and the suite can assert.
 *
 * The base64 is the package; everything else is evidence. `docs/DESIGN.md`
 * section 6 makes the deck delta the evidence an insert worked, never the
 * absence of an error, and a report that only carried the bytes would leave the
 * pane with nothing to say afterwards.
 */
export interface SpliceReport {
  base64: string;
  /** How many slides the user's deck holds. The splice does not change it. */
  deckSlides: number;
  /** The rebuilt slide inside the package: the one slide the package lists. */
  slidePath: string;
  /** Where the element landed, in EMU on the destination slide. */
  landed: Rect;
  /** How many top-level shapes were added: one when they were grouped. */
  shapes: number;
  grouped: boolean;
  /** Carried parts copied into the package. */
  parts: number;
  /** Empty content placeholders taken off the rebuilt slide. */
  placeholders: number;
}

/**
 * Leave exactly one slide listed in the deck's own order.
 *
 * The `<p:sldId>` entries go and the relationships stay, which is the "unlisted"
 * arm of probe question 1 rather than the "pruned" one — both landed a single
 * slide on the web, and this is the half that touches least.
 */
async function keepOnly(pkg: Pkg, slidePath: string): Promise<number> {
  const pres = await pkg.doc(PRESENTATION);
  const list = element(pres, P_NS, "sldIdLst");
  if (!list) throw new Error("ssf-slide-elements: this deck's presentation.xml has no <p:sldIdLst>");
  let removed = 0;
  for (const sldId of elements(list, P_NS, "sldId")) {
    const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
    const target = rId ? await pkg.relTarget(PRESENTATION, rId) : undefined;
    if (target === slidePath) continue;
    sldId.parentNode?.removeChild(sldId);
    removed += 1;
  }
  // The one slide that survives is the one the caller named. A package listing
  // nothing at all would be handed to `insertSlidesFromBase64` as a deck with
  // no slides in it, and the host's answer to that is not something any round
  // has measured — so it is refused here, where the cause is still nameable.
  if (!element(pres, P_NS, "sldIdLst")?.getElementsByTagNameNS(P_NS, "sldId").length) {
    throw new Error(`ssf-slide-elements: ${slidePath} is not in this deck's slide order, so nothing would be inserted`);
  }
  return removed;
}

/**
 * The deck with exactly one of its own slides listed, and nothing changed.
 *
 * What Undo hands back to PowerPoint. Taking an insert back that landed "onto
 * this slide" means putting the REPLACED slide back, and the only way to put a
 * slide into a deck is to hand the host a package containing it — so this is
 * the same unlisting trick the splice itself uses, over the bytes the pane read
 * before the insert.
 *
 * Nothing is rebuilt and nothing is cloned: the slide that comes back is the
 * user's own, byte for byte, which is the whole point of an undo.
 */
export async function onlySlide(deck: Uint8Array | string, slide: number): Promise<{ base64: string; path: string }> {
  const pkg = await Pkg.open(deck);
  const paths = await pkg.slidePaths();
  const path = paths[slide];
  if (path === undefined) {
    throw new Error(`ssf-slide-elements: slide ${slide + 1} is not in this deck, which has ${paths.length} slide(s)`);
  }
  await keepOnly(pkg, path);
  return { base64: await pkg.toBase64(), path };
}

/** The shape tree of a slide part. */
async function spTreeOf(pkg: Pkg, slidePath: string): Promise<Element> {
  const doc = await pkg.doc(slidePath);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  const spTree = cSld ? child(cSld, P_NS, "spTree") : undefined;
  if (!spTree) throw new Error(`ssf-slide-elements: ${slidePath} has no <p:spTree>`);
  return spTree;
}

/**
 * Empty a cloned slide, so "as a new slide" starts from the layout rather than
 * from what the user happened to be looking at.
 *
 * The clone is the basis rather than a slide built from nothing, because that
 * is what guarantees the new slide sits on the same layout and master as the
 * one it follows — and a slide whose layout is not in the deck is a slide
 * PowerPoint has to invent a design for.
 *
 * Placeholders stay and are emptied; everything else goes. The notes page goes
 * too: a new slide carrying the previous slide's speaker notes is a surprise
 * nobody asked for, and the part it points at is left in the package as an
 * orphan, which `scripts/package-integrity.mjs` treats as weight rather than
 * damage and which the insert never reaches, because nothing lists it.
 */
async function blank(pkg: Pkg, slidePath: string): Promise<void> {
  const spTree = await spTreeOf(pkg, slidePath);
  for (const shape of slideShapes(spTree)) {
    const isPlaceholder = placeholderIn(shape);
    if (!isPlaceholder) {
      shape.parentNode?.removeChild(shape);
      continue;
    }
    const txBody = child(shape, P_NS, "txBody");
    if (!txBody) continue;
    // Every paragraph but the first goes, and the first is emptied of runs. A
    // `<p:txBody>` with no `<a:p>` at all is schema-invalid; PowerPoint reports
    // the file as damaged and names nothing.
    const paragraphs = children(txBody, A_NS, "p");
    for (const p of paragraphs.slice(1)) p.parentNode?.removeChild(p);
    const first = paragraphs[0];
    if (!first) continue;
    for (const node of Array.from(first.childNodes)) {
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      // `<a:pPr>` and `<a:endParaRPr>` are the paragraph's own formatting and
      // are kept, so the empty placeholder still looks like itself. Runs go.
      if (el.namespaceURI === A_NS && (el.localName === "pPr" || el.localName === "endParaRPr")) continue;
      first.removeChild(el);
    }
  }

  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return;
  const rels = await pkg.doc(relsPath);
  for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
    if (rel.getAttribute("Type") === REL_TYPE.notesSlide) rel.parentNode?.removeChild(rel);
  }
}

/** Whether a top-level shape claims a placeholder on its slide. */
function placeholderIn(shape: Element): boolean {
  for (const node of Array.from(shape.childNodes)) {
    if (node.nodeType !== 1) continue;
    const nvPr = child(node as Element, P_NS, "nvPr");
    if (nvPr && child(nvPr, P_NS, "ph")) return true;
  }
  return false;
}

/**
 * Splice an element into a deck and answer the package to hand PowerPoint.
 *
 * The order of the steps is not arrangeable. Ids are renumbered against the
 * destination slide BEFORE anything is appended to it, so the numbering sees
 * the slide as it was; the parts are carried and the references repointed
 * BEFORE the shapes are adopted, so a failure to find a part leaves the user's
 * deck untouched rather than half-written; and the tags are written LAST,
 * because a tag is a relationship on the slide and the shape has to be in the
 * slide to carry the reference.
 */
export async function splice(request: SpliceRequest): Promise<SpliceReport> {
  const pkg = await Pkg.open(request.deck);
  const paths = await pkg.slidePaths();
  const deckSlides = paths.length;
  if (deckSlides === 0) throw new Error("ssf-slide-elements: this deck has no slides to insert into");
  if (!Number.isInteger(request.slide) || request.slide < 0 || request.slide >= deckSlides) {
    throw new Error(
      `ssf-slide-elements: slide ${request.slide + 1} is not in this deck, which has ${deckSlides} slide(s)`,
    );
  }
  const source = paths[request.slide] as string;

  const size = await slideSize(pkg);
  const frames = await framesOf(pkg, source);

  const rebuilt = await cloneSlide(pkg, source);
  if (request.options.target === "new") await blank(pkg, rebuilt);

  const spTree = await spTreeOf(pkg, rebuilt);
  const fragment = parseFragment(request.element.markup.xml);
  const tops = topLevel(fragment);
  if (tops.length === 0) throw new Error(`ssf-slide-elements: the element "${request.element.id}" has no shapes`);

  // A library shape can claim to BE the slide's body. On a destination that
  // already has one, two shapes claim the same index and PowerPoint resolves it
  // by inheriting geometry from the layout — the element jumps to wherever the
  // layout's placeholder is and the landing computed below is thrown away.
  unplaceholder(tops);

  let nextId = highestShapeId(spTree) + 1;
  nextId = renumber(fragment, nextId);

  const carried = await carry({
    pkg,
    owner: rebuilt,
    rels: request.element.markup.rels,
    types: request.catalogue.carried,
    store: request.store,
  });
  repoint(fragment, carried.ids);

  // The BOX is the reference frame, not the union of the shapes' own
  // rectangles, and the two are not the same thing: the catalogue's box is a
  // rotated shape's rotated extent (`docs/DESIGN.md` section 3), which is the
  // ink the user sees. Landing the ink where the rule says is the promise;
  // landing the unrotated frame there would put a 29° stamp off the edge.
  const from = authored(request.element.box, size);
  const landed = place({
    box: request.element.box,
    landing: request.element.landing,
    slide: size,
    frames,
    ...(request.selection ? { selection: request.selection } : {}),
    ...(request.element.wraps ? { wraps: true } : {}),
  });
  applyMove(tops, from, moveFrom(from, landed));

  // The group takes the next id after the shapes it wraps, which `renumber`
  // has just told us. Nothing needs an id after this one — a second element is
  // a second call to `splice`, which reads the slide again — so the counter is
  // spent here rather than advanced for a caller that does not exist.
  const wanted = request.options.group && groupable(tops);
  if (wanted) groupShapes(fragment, tops, request.element.name, nextId);

  const doc = spTree.ownerDocument;
  if (!doc) throw new Error(`ssf-slide-elements: ${rebuilt} belongs to no document`);
  const added: Element[] = [];
  for (const shape of topLevel(fragment)) {
    const imported = doc.importNode(shape, true);
    spTree.appendChild(imported);
    added.push(imported);
  }

  // `docs/DESIGN.md` section 6: a whole-slide element removes the "Click to add
  // text" ghosts it lands over. The TITLE stays, and a placeholder the user has
  // typed into is content rather than a ghost — `emptyBodyPlaceholders` is
  // where both of those live.
  let placeholders = 0;
  if (request.element.kind === "slide") {
    for (const ghost of emptyBodyPlaceholders(spTree)) {
      ghost.parentNode?.removeChild(ghost);
      placeholders += 1;
    }
  }

  for (const shape of added) {
    await writeShapeTags(pkg, rebuilt, shape, [
      [TAG_ELEMENT, request.element.id],
      [TAG_CATALOGUE, request.catalogue.version],
    ]);
  }

  await keepOnly(pkg, rebuilt);

  return {
    base64: await pkg.toBase64(),
    deckSlides,
    slidePath: rebuilt,
    landed,
    shapes: added.length,
    grouped: wanted,
    parts: carried.parts.size,
    placeholders,
  };
}
